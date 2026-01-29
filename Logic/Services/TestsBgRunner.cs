using Analitics6400.Logic.Test.Interfaces;

public class TestsBgRunner : BackgroundService
{
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly ILogger<TestsBgRunner> _logger;

    public TestsBgRunner(
        IServiceScopeFactory scopeFactory,
        ILogger<TestsBgRunner> logger)
    {
        _scopeFactory = scopeFactory;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        using var scope = _scopeFactory.CreateScope();
        var testRunners = scope.ServiceProvider.GetServices<IXmlTest>();

        foreach (var testRunner in testRunners)
        {
            await testRunner.RunAsync(ct: stoppingToken);
            _logger.LogInformation($"Test {testRunner.Name} completed.");
        }
    }
}