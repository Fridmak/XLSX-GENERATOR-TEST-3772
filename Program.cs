using Analitics6400.Dal;
using Analitics6400.Dal.Services;
using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Analitics6400.Logic.Seed;
using Analitics6400.Logic.Services;
using Analitics6400.Logic.Services.XmlWriters;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Test;
using Analitics6400.Logic.Test.Interfaces;
using Microsoft.EntityFrameworkCore;
using Npgsql;

public class Programm
{
    public static async Task Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        ConfigureDataProviders(builder);

        builder.Services.AddTransient<DocumentSeeder>();

        ConfigureTestServices(builder);

        builder.Services.AddHostedService<TestsBgRunner>();

        SetSettings(builder);

        var app = builder.Build();

        //await StartFillingDbWithDocuments(app, 50_000);

        app.Run();
    }

    private static void ConfigureDataProviders(WebApplicationBuilder builder)
    {
        var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

        var dataSourceBuilder = new NpgsqlDataSourceBuilder(connectionString);
        dataSourceBuilder.EnableDynamicJson();
        var dataSource = dataSourceBuilder.Build();

        builder.Services.AddDbContext<DocumentDbContext>(options =>
        {
            options.UseNpgsql(dataSource);
        });


        //builder.Services.AddTransient<IDocumentProvider, DocumentDalProvider>();
        //builder.Services.AddTransient<IDocumentProvider, NpgsqlDocumentDalProvider>();
        builder.Services.AddTransient<IDocumentProvider>(sp => new MockDocumentDalProvider(documentCount: 100_000)); // внутри настройка распределения документов
    }

    private static void ConfigureTestServices(WebApplicationBuilder builder)
    {
        builder.Services.AddTransient<IXmlWriter, AbankingCsvWriter>();
        builder.Services.AddTransient<IXmlWriter, AbankingOpenXmlWriter>();
        builder.Services.AddTransient<IXmlWriter, AbankingClosedXmlWriter>();

        // Выбираем нужный:
        builder.Services.AddTransient<IXmlTest, XmlTest<AbankingOpenXmlWriter>>();
        //builder.Services.AddTransient<IXmlTest, XmlTest<AbankingClosedXmlWriter>>();
        //builder.Services.AddTransient<IXmlTest, XmlTest<AbankingCsvWriter>>(); // Полностю рабочий и корректный
    }

    private static void SetSettings(WebApplicationBuilder builder)
    {
        builder.Services.Configure<XmlTestSettings>(
            builder.Configuration.GetSection(nameof(XmlTestSettings)));
    }

    #region Заполнение бд документами
    /// <summary>
    /// Начать заполнение бд тестовыми данными
    /// </summary>
    /// <param name="app"></param>
    /// <returns></returns>
    private static async Task StartFillingDbWithDocuments(WebApplication app, int documentsCount)
    {
        using (var scope = app.Services.CreateScope()) // Генерация в бд
        {
            var db = scope.ServiceProvider.GetRequiredService<DocumentDbContext>();
            db.Database.Migrate();

            var seeder = scope.ServiceProvider.GetRequiredService<DocumentSeeder>();
            await seeder.SeedAsync(totalDocuments: documentsCount);
        }
    }
    #endregion
}
