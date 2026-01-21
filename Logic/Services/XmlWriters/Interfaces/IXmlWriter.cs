using Analitics6400.Logic.Services.XmlWriters.Models;

namespace Analitics6400.Logic.Services.XmlWriters.Interfaces;

public interface IXmlWriter
{
     Task<byte[]> GenerateAsync<T>(
         IAsyncEnumerable<T> rows,
         IReadOnlyList<ExcelColumn> columns,
         CancellationToken ct = default);
}
