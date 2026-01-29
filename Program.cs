using Analitics6400.Dal;
using Analitics6400.Dal.Services;
using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Seed;
using Analitics6400.Logic.Services;
using Analitics6400.Logic.Services.XmlWriters;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Test;
using Analitics6400.Logic.Test.Interfaces;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Microsoft.EntityFrameworkCore;
using Npgsql;

var builder = WebApplication.CreateBuilder(args);

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

var dataSourceBuilder = new NpgsqlDataSourceBuilder(connectionString);
dataSourceBuilder.EnableDynamicJson();
var dataSource = dataSourceBuilder.Build();

builder.Services.AddDbContext<DocumentDbContext>(options =>
{
    options.UseNpgsql(dataSource);
});

builder.Services.AddTransient<DocumentSeeder>();
builder.Services.AddTransient<IDocumentProvider, DocumentDalProvider>();
//builder.Services.AddTransient<IDocumentProvider, NpgsqlDocumentDalProvider>();

//builder.Services.AddTransient<IDocumentProvider>(sp =>
//    new MockDocumentDalProvider(documentCount: 100000));

//builder.Services.AddTransient<IXmlTest, XmlTest<AbankingOpenXmlWriter>>();
//builder.Services.AddTransient<IXmlTest, XmlTest<AbankingClosedXmlWriter>>();
//builder.Services.AddTransient<IXmlTest, XmlTest<AbankingCsvWriter>>();
builder.Services.AddTransient<IXmlTest, XmlTest<AbankingClosedXmlWriterBenchmark>>();

builder.Services.AddTransient<IXmlWriter, AbankingCsvWriter>();
builder.Services.AddTransient<IXmlWriter, AbankingOpenXmlWriter>();
builder.Services.AddTransient<IXmlWriter, AbankingClosedXmlWriter>();
builder.Services.AddTransient<IXmlWriter, AbankingClosedXmlWriterBenchmark>();

builder.Services.AddHostedService<TestsBgRunner>();

var app = builder.Build();

//using (var scope = app.Services.CreateScope()) // Генерация в бд
//{
//    var db = scope.ServiceProvider.GetRequiredService<DocumentDbContext>();
//    db.Database.Migrate();

//    var seeder = scope.ServiceProvider.GetRequiredService<DocumentSeeder>();
//    await seeder.SeedAsync(totalDocuments: 17000);
//}

app.Run();