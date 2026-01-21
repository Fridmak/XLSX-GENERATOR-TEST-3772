using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Microsoft.EntityFrameworkCore;

namespace Analitics6400.Dal.Services;

public sealed class DocumentDalProvider : IDocumentProvider
{
    private readonly DocumentDbContext _db;

    public DocumentDalProvider(DocumentDbContext db)
    {
        _db = db;
    }

    public IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(CancellationToken ct = default)
    {
        return _db.Documents
            .AsNoTracking()
            .OrderBy(d => d.Id)
            .Select(d => new DocumentDtoModel
            {
                Id = d.Id,
                DocumentSchemaId = d.DocumentSchemaId,
                Published = d.Published,
                IsArchived = d.IsArchived,
                Version = d.Version,
                IsCanForValidate = d.IsCanForValidate,
                ChangedDateUtc = d.ChangedDateUtc
            })
            .AsAsyncEnumerable();
    }
}
