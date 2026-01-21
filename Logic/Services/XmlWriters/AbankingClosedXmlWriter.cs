using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;

namespace Analitics6400.Logic.Services.XmlWriters;

public class AbankingClosedXmlWriter : IXmlWriter
{
    public Task<byte[]> GenerateAsync<T>(IAsyncEnumerable<T> rows, IReadOnlyList<ExcelColumn> columns, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }
}
