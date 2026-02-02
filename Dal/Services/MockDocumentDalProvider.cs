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

    private readonly (int Min, int Max, double Probability)[] _sizeDistribution = new[]
    {
        (1_000, 40_000, 0.50),
        (40_001, 800_000, 0.30),
        (800_001, 1_000_000, 0.20)
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

            var targetSize = GetTargetSize();
            var jsonObject = GenerateJsonOfExactSize(targetSize, i);

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow.AddDays(-_random.Next(0, 365)),
                Version = Math.Round(1.0 + _random.NextDouble() * 9, 1),
                IsArchived = _random.NextDouble() < 0.1,
                JsonData = jsonObject,
                IsCanForValidate = _random.NextDouble() < 0.8,
                ChangedDateUtc = DateTime.UtcNow.AddHours(-_random.Next(0, 720))
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            if (i % 100 == 0)
                await Task.Yield();
        }
    }

    private int GetTargetSize()
    {
        var r = _random.NextDouble();
        double cumulative = 0;

        foreach (var (min, max, probability) in _sizeDistribution)
        {
            cumulative += probability;
            if (r <= cumulative)
            {
                return _random.Next(min, max + 1);
            }
        }

        return _random.Next(1_000, 1_000_001);
    }

    private JsonObject GenerateJsonOfExactSize(int targetSizeBytes, int index)
    {
        var docType = _documentTypes[_random.Next(_documentTypes.Count)];
        var obj = CreateBaseJsonObject(docType, index);

        // Получаем текущий размер
        var currentJson = obj.ToJsonString();
        var currentSize = Encoding.UTF8.GetByteCount(currentJson);

        // Добавляем данные пока не достигнем нужного размера
        if (currentSize < targetSizeBytes)
        {
            AddPaddingData(obj, targetSizeBytes - currentSize);
        }
        else if (currentSize > targetSizeBytes)
        {
            // Если перебор, создаем новый меньший объект
            obj = CreateMinimalJsonObject(docType, index);
            currentJson = obj.ToJsonString();
            currentSize = Encoding.UTF8.GetByteCount(currentJson);

            if (currentSize < targetSizeBytes)
            {
                AddPaddingData(obj, targetSizeBytes - currentSize);
            }
        }

        // Финальная проверка
        var finalJson = obj.ToJsonString();
        var finalSize = Encoding.UTF8.GetByteCount(finalJson);

        // Если все еще не совпадает, добавляем/убираем padding
        if (finalSize != targetSizeBytes)
        {
            AdjustJsonSize(obj, targetSizeBytes, finalSize);
        }

        return obj;
    }

    private JsonObject CreateBaseJsonObject(string docType, int index)
    {
        return new JsonObject
        {
            ["documentId"] = Guid.NewGuid().ToString(),
            ["type"] = docType,
            ["number"] = $"DOC-{DateTime.UtcNow:yyyyMMdd}-{index:00000}",
            ["createdAt"] = DateTime.UtcNow.ToString("O"),
            ["status"] = _random.NextDouble() < 0.8 ? "Active" : "Archived"
        };
    }

    private JsonObject CreateMinimalJsonObject(string docType, int index)
    {
        var obj = CreateBaseJsonObject(docType, index);
        obj["data"] = "minimal";
        return obj;
    }

    private void AddPaddingData(JsonObject obj, int bytesNeeded)
    {
        if (bytesNeeded <= 0) return;

        // Создаем строку нужного размера
        var paddingKey = $"padding_{Guid.NewGuid():N}";

        // Рассчитываем размер ключа и кавычек
        var keySize = Encoding.UTF8.GetByteCount($"\"{paddingKey}\":\"");
        var closingSize = Encoding.UTF8.GetByteCount("\"");
        var commaSize = obj.Count > 0 ? 1 : 0; // запятая между элементами

        var valueSize = bytesNeeded - keySize - closingSize - commaSize;

        if (valueSize <= 0)
        {
            // Если места мало, просто добавляем маленькое значение
            obj[paddingKey] = "x";
            return;
        }

        // Создаем строку точно нужного размера
        var paddingValue = GeneratePaddingString(valueSize);
        obj[paddingKey] = paddingValue;
    }

    private string GeneratePaddingString(int exactByteSize)
    {
        // Используем ASCII символы для точного контроля размера
        const string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        var builder = new StringBuilder(exactByteSize);

        // Заполняем предсказуемыми данными
        for (int i = 0; i < exactByteSize; i++)
        {
            builder.Append(chars[i % chars.Length]);
        }

        return builder.ToString();
    }

    private void AdjustJsonSize(JsonObject obj, int targetSize, int currentSize)
    {
        var json = obj.ToJsonString();
        var difference = targetSize - currentSize;

        if (difference == 0) return;

        if (difference > 0)
        {
            // Нужно добавить байты
            var paddingKey = $"size_adjust_{Math.Abs(difference)}";

            // Учитываем размер нового поля
            var keyPart = $"\"{paddingKey}\":\"";
            var keySize = Encoding.UTF8.GetByteCount(keyPart);
            var closingSize = Encoding.UTF8.GetByteCount("\"");
            var commaSize = 1;

            var availableForValue = difference - keySize - closingSize - commaSize;

            if (availableForValue > 0)
            {
                obj[paddingKey] = new string('x', availableForValue);
            }
            else
            {
                // Если места совсем мало, добавляем в существующее поле
                var existingKey = obj.First().Key;
                var existingValue = obj[existingKey]?.ToString() ?? "";
                obj[existingKey] = existingValue + new string('x', difference);
            }
        }
        else
        {
            // Нужно убрать байты (difference отрицательное)
            difference = Math.Abs(difference);
            var lastKey = obj.Last().Key;
            var lastValue = obj[lastKey]?.ToString() ?? "";

            if (lastValue.Length > difference)
            {
                obj[lastKey] = lastValue.Substring(0, lastValue.Length - difference);
            }
            else
            {
                obj.Remove(lastKey);
            }
        }
    }

    public void Dispose()
    {
        // Нет неуправляемых ресурсов
    }
}