using Analitics6400.Logic.Models;

namespace Analitics6400.Dal.Services.Interfaces;

public interface IDocumentProvider
{
    IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(CancellationToken ct = default);
}
