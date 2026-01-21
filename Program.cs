using Analitics6400.Dal;
using Analitics6400.Logic.Seed;
using Analitics6400.Logic.Services;
using Analitics6400.Logic.Services.XmlWriters;
using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Test;
using Analitics6400.Logic.Test.Interfaces;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<DocumentDbContext>(options =>
{
    options.UseNpgsql(builder.Configuration.GetConnectionString("DefaultConnection"));
});

builder.Services.AddScoped<DocumentSeeder>();

builder.Services.AddTransient<IXmlTest, XmlTest<AbankingOpenXmlWriter>>();
builder.Services.AddTransient<IXmlTest, XmlTest<AbankingClosedXmlWriter>>();

builder.Services.AddTransient<IXmlWriter, AbankingOpenXmlWriter>();
builder.Services.AddTransient<IXmlWriter, AbankingClosedXmlWriter>();

builder.Services.AddHostedService<TestsBgRunner>();

var app = builder.Build();

//using (var scope = app.Services.CreateScope())
//{
//    var db = scope.ServiceProvider.GetRequiredService<DocumentDbContext>();
//    db.Database.Migrate();

//    var seeder = scope.ServiceProvider.GetRequiredService<DocumentSeeder>();
//    await seeder.SeedAsync();
//}

app.Run();