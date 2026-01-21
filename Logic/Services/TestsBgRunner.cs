using Analitics6400.Logic.Test.Interfaces;

namespace Analitics6400.Logic.Services;

public class TestsBgRunner : BackgroundService
{
    private readonly IServiceScopeFactory _scopeFactory;

    public TestsBgRunner(IServiceScopeFactory scopeFactory)
    {
        _scopeFactory = scopeFactory;
    }
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        using var scope = _scopeFactory.CreateScope();

        var testRunners = scope.ServiceProvider.GetServices<IXmlTest>();

        foreach (var testRunner in testRunners)
        {
            await testRunner.RunAsync();

            Console.WriteLine($"Test {testRunner.Name}");
        }
    }
}
