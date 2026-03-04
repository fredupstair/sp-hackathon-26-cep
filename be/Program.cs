using CepFunctions.Functions;
using CepFunctions.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureServices((ctx, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Singletons – reuse HttpClient and token credentials
        services.AddSingleton<GraphClient>();
        services.AddSingleton<SharePointClient>();

        // Scoped per invocation
        services.AddScoped<PointsEngine>();
        services.AddScoped<BadgeEngine>();
        services.AddScoped<TeamsNotifier>();

        // Register OrchestratorTimer so it can be injected into AdminApi
        services.AddTransient<OrchestratorTimer>();
    })
    .Build();

host.Run();
