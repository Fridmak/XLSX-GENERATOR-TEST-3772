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

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
    int? limit = null,
    [EnumeratorCancellation] CancellationToken ct = default)
    {
        await using var conn = new NpgsqlConnection(_connectionString);
        await conn.OpenAsync(ct);

        const int batchSize = 1000;
        int offset = 0;
        int returned = 0;

        while (!ct.IsCancellationRequested)
        {
            var sql = $"""
        SELECT "Id", "DocumentSchemaId", "Published", "IsArchived",
               "Version", "JsonData", "IsCanForValidate", "ChangedDateUtc"
        FROM "documents"
        ORDER BY "Id"
        LIMIT @BatchSize OFFSET @Offset
        """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.CommandType = CommandType.Text;

            // Добавляем параметры напрямую в команду
            cmd.Parameters.AddWithValue("@BatchSize", batchSize);
            cmd.Parameters.AddWithValue("@Offset", offset);

            await using var reader = await cmd.ExecuteReaderAsync(
                CommandBehavior.SequentialAccess, ct);

            var hasRows = false;
            while (await reader.ReadAsync(ct))
            {
                hasRows = true;

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
                offset++;

                if (limit.HasValue && returned >= limit.Value)
                {
                    yield break;
                }
            }

            if (!hasRows)
                break;
        }
    }

}
