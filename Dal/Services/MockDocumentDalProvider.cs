using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Analitics6400.Dal.Services;

/// <summary>
/// Реалистичный мок-провайдер документов для тестов генерации Excel.
/// Генерирует документы с разными размерами и структурами.
/// </summary>
public sealed class MockDocumentDalProvider : IDocumentProvider, IDisposable
{
    private readonly int _documentCount;
    private readonly Random _random;
    private readonly List<string> _documentTypes = new()
    {
        "Invoice", "Contract", "Report", "Application", "Order",
        "Payment", "Receipt", "Statement", "Certificate", "License"
    };

    // Более реалистичное распределение размеров
    // 70% - маленькие (<10KB), 25% - средние (10-100KB), 5% - большие (>100KB)
    private readonly (int Min, int Max, double Probability)[] _sizeDistribution = new[]
    {
        (100, 10_000, 0.40),    // Маленькие
        (10_001, 100_000, 0.45), // Средние
        (100_001, 1_000_000, 0.15) // Большие
    };

    public MockDocumentDalProvider(int documentCount = 1000, int randomSeed = 123456789)
    {
        _documentCount = documentCount;
        _random = new(Seed: randomSeed);
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
        int? limit = null,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        int returned = 0;
        var documentCount = limit ?? _documentCount;

        for (int i = 0; i < documentCount; i++)
        {
            ct.ThrowIfCancellationRequested();

            // Создаем документ с различной структурой
            var jsonObject = GenerateRealisticJson(i);

            // Сериализуем для проверки размера
            var jsonString = jsonObject.ToJsonString();
            var currentSize = Encoding.UTF8.GetByteCount(jsonString);

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow.AddDays(-_random.Next(0, 365)),
                Version = Math.Round(1.0 + _random.NextDouble() * 9, 1), // Версии 1.0-10.0
                IsArchived = _random.NextDouble() < 0.1, // 10% архивных
                JsonData = jsonObject,
                IsCanForValidate = _random.NextDouble() < 0.8, // 80% можно валидировать
                ChangedDateUtc = DateTime.UtcNow.AddHours(-_random.Next(0, 720)) // До 30 дней назад
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            // Регулярно освобождаем поток для асинхронности
            if (i % 100 == 0)
                await Task.Yield();
        }
    }

    private JsonObject GenerateRealisticJson(int index)
    {
        var docType = _documentTypes[_random.Next(_documentTypes.Count)];
        var (minSize, maxSize, _) = GetRandomSizeCategory();

        // Базовый объект
        var obj = new JsonObject
        {
            ["documentId"] = Guid.NewGuid().ToString(),
            ["type"] = docType,
            ["number"] = $"DOC-{DateTime.UtcNow:yyyyMMdd}-{index:00000}",
            ["createdAt"] = DateTime.UtcNow.ToString("O"),
            ["status"] = _random.NextDouble() < 0.8 ? "Active" : "Archived"
        };

        // Добавляем различные поля в зависимости от типа
        switch (docType)
        {
            case "Invoice":
                AddInvoiceData(obj);
                break;
            case "Contract":
                AddContractData(obj);
                break;
            case "Report":
                AddReportData(obj);
                break;
            default:
                AddGenericData(obj);
                break;
        }

        // Добавляем дополнительные поля для увеличения размера (если нужно)
        var currentJson = obj.ToJsonString();
        var currentSize = Encoding.UTF8.GetByteCount(currentJson);

        if (currentSize < minSize)
        {
            AddAdditionalData(obj, minSize - currentSize);
        }

        return obj;
    }

    private void AddInvoiceData(JsonObject obj)
    {
        obj["totalAmount"] = Math.Round(_random.NextDouble() * 10000, 2);
        obj["currency"] = _random.NextDouble() < 0.7 ? "RUB" : "USD";
        obj["vendor"] = $"Vendor_{_random.Next(1000)}";

        var items = new JsonArray();
        int itemCount = _random.Next(1, 50);

        for (int i = 0; i < itemCount; i++)
        {
            items.Add(new JsonObject
            {
                ["product"] = $"Product_{_random.Next(1000)}",
                ["quantity"] = _random.Next(1, 100),
                ["price"] = Math.Round(_random.NextDouble() * 100, 2),
                ["total"] = Math.Round(_random.NextDouble() * 1000, 2)
            });
        }

        obj["items"] = items;
    }

    private void AddContractData(JsonObject obj)
    {
        obj["startDate"] = DateTime.UtcNow.ToString("O");
        obj["endDate"] = DateTime.UtcNow.AddDays(_random.Next(30, 365)).ToString("O");
        obj["parties"] = new JsonArray
        {
            new JsonObject { ["name"] = $"Party_A_{_random.Next(100)}", ["role"] = "Supplier" },
            new JsonObject { ["name"] = $"Party_B_{_random.Next(100)}", ["role"] = "Client" }
        };

        // Длинное текстовое поле
        var clauses = new StringBuilder();
        int clauseCount = _random.Next(5, 50);
        for (int i = 0; i < clauseCount; i++)
        {
            clauses.AppendLine($"Clause {i + 1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit. ");
        }
        obj["terms"] = clauses.ToString();
    }

    private void AddReportData(JsonObject obj)
    {
        obj["period"] = $"{DateTime.UtcNow:yyyy-MM}";
        obj["generatedBy"] = $"User_{_random.Next(100)}";

        var metrics = new JsonObject();
        for (int i = 0; i < _random.Next(10, 100); i++)
        {
            metrics[$"metric_{i}"] = Math.Round(_random.NextDouble() * 1000, 2);
        }
        obj["metrics"] = metrics;
    }

    private void AddGenericData(JsonObject obj)
    {
        obj["description"] = $"Generic document {Guid.NewGuid()}";

        var metadata = new JsonObject();
        int metaCount = _random.Next(5, 20);
        for (int i = 0; i < metaCount; i++)
        {
            metadata[$"field_{i}"] = $"value_{_random.Next(1000)}";
        }
        obj["metadata"] = metadata;
    }

    private void AddAdditionalData(JsonObject obj, int additionalBytesNeeded)
    {
        if (additionalBytesNeeded <= 0) return;

        // Добавляем большое текстовое поле
        var largeField = new StringBuilder();
        int approxCharsNeeded = additionalBytesNeeded / 2; // Примерно 2 байта на символ UTF-8

        while (largeField.Length < approxCharsNeeded)
        {
            largeField.Append("Sample text for padding. ");
        }

        obj["additionalData"] = largeField.ToString();

        // Или добавляем массив
        if (_random.NextDouble() < 0.5)
        {
            var array = new JsonArray();
            int itemsNeeded = additionalBytesNeeded / 100; // ~100 байт на элемент

            for (int i = 0; i < itemsNeeded; i++)
            {
                array.Add(new JsonObject
                {
                    ["id"] = Guid.NewGuid().ToString(),
                    ["value"] = _random.NextDouble(),
                    ["timestamp"] = DateTime.UtcNow.ToString("O")
                });
            }

            obj["extraItems"] = array;
        }
    }

    private (int Min, int Max, double Probability) GetRandomSizeCategory()
    {
        var r = _random.NextDouble();
        double cumulative = 0;

        foreach (var category in _sizeDistribution)
        {
            cumulative += category.Probability;
            if (r <= cumulative)
            {
                return category;
            }
        }

        return _sizeDistribution[0];
    }

    public void Dispose()
    {
        // Нет неуправляемых ресурсов
    }
}
