using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Test.Interfaces;

namespace Analitics6400.Logic.Test;

public sealed class XmlTest<T> : IXmlTest where T : IXmlWriter
{
    public string Name => nameof(T);

    private readonly IXmlWriter _writer;

    public XmlTest(IEnumerable<IXmlWriter> writer)
    {
        _writer = writer.Where(x => x is T).First();
    }

    public async Task RunAsync()
    {
        await _writer.GenerateAsync();
    }
}
