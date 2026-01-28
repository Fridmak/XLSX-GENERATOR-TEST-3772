using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Microsoft.EntityFrameworkCore;
using System.Runtime.CompilerServices;
using System.Text;

namespace Analitics6400.Dal.Services;

public sealed class DocumentDalProvider : IDocumentProvider
{
    private readonly DocumentDbContext _db;

    public DocumentDalProvider(DocumentDbContext db)
    {
        _db = db;
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(int? limit = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        var query = _db.Documents
            .AsNoTracking()
            .OrderBy(d => d.Id)
            .Select(d => new DocumentDtoModel
            {
                Id = d.Id,
                DocumentSchemaId = d.DocumentSchemaId,
                Published = d.Published,
                IsArchived = d.IsArchived,
                Version = d.Version,
                JsonData = d.JsonData,
                IsCanForValidate = d.IsCanForValidate,
                ChangedDateUtc = d.ChangedDateUtc
            });


        if (limit.HasValue)
        {
            query = query.Take(limit.Value);
        }

        int returned = 0;

        await foreach (var doc in query.AsAsyncEnumerable().WithCancellation(ct))
        {
            yield return doc;

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;
        }
    }

}
