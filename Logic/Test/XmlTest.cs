using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Analitics6400.Logic.Services.XmlWriters;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using Analitics6400.Logic.Test.Interfaces;

namespace Analitics6400.Logic.Test;

public sealed class XmlTest<T> : IXmlTest where T : IXmlWriter
{
    private readonly IXmlWriter _writer;
    private readonly IDocumentProvider _documentProvider;
    public string Name => nameof(T); 

    public XmlTest(IEnumerable<IXmlWriter> writer, IDocumentProvider documentProvider)
    {
        _writer = writer.Where(x => x is T).First();
        _documentProvider = documentProvider;
    }

    public async Task RunAsync()
    {
        var columns = new List<ExcelColumn>
        {
            new("Id", typeof(Guid)),
            new("DocumentSchemaId", typeof(Guid?)),
            new("Published", typeof(DateTime?)),
            new("IsArchived", typeof(bool)),
            new("Version", typeof(double)),
            new("JsonData", typeof(string)),
            new("IsCanForValidate", typeof(bool)),
            new("ChangedDateUtc", typeof(DateTime?))
        };

        var rows = _documentProvider.GetDocumentsAsync();

        var bytes = await _writer.GenerateAsync(rows, columns);

        var fileName = $"DocumentsReport_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx";
        await File.WriteAllBytesAsync(fileName, bytes);

        Console.WriteLine($"Report generated: {fileName}, size: {bytes.Length} bytes");
    }

}
