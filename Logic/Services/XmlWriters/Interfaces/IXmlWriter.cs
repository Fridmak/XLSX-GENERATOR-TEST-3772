using Analitics6400.Logic.Services.XmlWriters.Models;

namespace Analitics6400.Logic.Services.XmlWriters.Interfaces;

public interface IXmlWriter
{

    string Extension { get; }
    Task GenerateAsync<T>(
         IAsyncEnumerable<T> rows,
         IReadOnlyList<ExcelColumn> columns,
         string output,
         CancellationToken ct = default);
}
