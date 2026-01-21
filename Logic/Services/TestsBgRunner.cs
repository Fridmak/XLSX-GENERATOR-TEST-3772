using Analitics6400.Logic.Test.Interfaces;

namespace Analitics6400.Logic.Services;

public class TestsBgRunner : BackgroundService
{
    private readonly IEnumerable<IXmlTest> _xmlTestRunners;

    public TestsBgRunner(IEnumerable<IXmlTest> xmlTestRunners)
    {
        _xmlTestRunners = xmlTestRunners;
    }
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        foreach (var testRunner in _xmlTestRunners)
        {
            await testRunner.RunAsync();

            Console.WriteLine($"Test {testRunner.Name}");
        }
    }
}
