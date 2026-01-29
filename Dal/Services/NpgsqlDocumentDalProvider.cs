using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Npgsql;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json.Nodes;

namespace Analitics6400.Dal.Services;

public sealed class NpgsqlDocumentDalProvider : IDocumentProvider
{
    private readonly string _connectionString;

    public NpgsqlDocumentDalProvider(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string not found");
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(int? limit = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync(ct);

        var sql = new StringBuilder();
        sql.Append("""
        SELECT "Id", "DocumentSchemaId", "Published", "IsArchived",
               "Version", "JsonData", "IsCanForValidate", "ChangedDateUtc"
        FROM "documents"
        ORDER BY "Id"
        """);

        if (limit.HasValue)
        {
            sql.AppendLine();
            sql.Append("LIMIT @limit");
        }


        await using var cmd = new NpgsqlCommand(sql.ToString(), conn);

        if (limit.HasValue)
        {
            cmd.Parameters.AddWithValue("@limit", limit.Value);
        }

        await using var reader = await cmd.ExecuteReaderAsync(
            CommandBehavior.SequentialAccess, ct);

        int returned = 0;

        while (await reader.ReadAsync(ct))
        {
            yield return new DocumentDtoModel
            {
                Id = reader.GetGuid(0),
                DocumentSchemaId = reader.GetGuid(1),
                Published = reader.GetDateTime(2),
                IsArchived = reader.GetBoolean(3),
                Version = reader.GetDouble(4),
                JsonData = JsonNode.Parse(reader.GetString(5)) as JsonObject,
                IsCanForValidate = reader.GetBoolean(6),
                ChangedDateUtc = reader.GetDateTime(7)
            };

            returned++;

            if (limit.HasValue && returned >= limit.Value)
            {
                yield break;
            }
        }
    }
}