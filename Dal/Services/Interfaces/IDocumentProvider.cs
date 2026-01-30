using Analitics6400.Logic.Models;
using System.Runtime.CompilerServices;

namespace Analitics6400.Dal.Services.Interfaces;

public interface IDocumentProvider
{
    IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
        int? limit = null,
        CancellationToken ct = default);
}
