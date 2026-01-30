using Analitics6400.Logic.Models;

namespace Analitics6400.Logic.Test.Interfaces;

public interface IXmlTest
{
    public string Name { get; }
    public Task RunAsync(CancellationToken ct = default);
}
