using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;
using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;

namespace Analitics6400.Dal.Services;

/// <summary>
/// Простой мок-провайдер документов для тестов.
/// Генерирует документы с JSON случайного размера.
/// </summary>
public sealed class MockDocumentDalProvider : IDocumentProvider, IDisposable
{
    private readonly int _documentCount;
    private readonly Random _random = new();

    private const int MinSize = 10_000;       // 10 КБ
    private const int MaxSize = 1_000_000;    // 1 МБ
    private const double Lambda = 1.0 / 50_000; // средний размер ~50 КБ

    public MockDocumentDalProvider(int documentCount = 1000)
    {
        _documentCount = documentCount;
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
        int? limit = null,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        int returned = 0;

        for (int i = 0; i < _documentCount; i++)
        {
            ct.ThrowIfCancellationRequested();

            var jsonObject = GenerateRandomJson();

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow,
                Version = 1.0,
                IsArchived = false,
                JsonData = jsonObject,
                IsCanForValidate = true,
                ChangedDateUtc = DateTime.UtcNow
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            await Task.Yield();
        }
    }

    private JsonObject GenerateRandomJson()
    {
        int targetSize = GetExponentialSize();

        var obj = new JsonObject
        {
            ["documentId"] = Guid.NewGuid().ToString(),
            ["createdAt"] = DateTime.UtcNow.ToString("O"),
            ["items"] = new JsonArray()
        };

        int estimatedItemSize = 200; // приблизительный размер одного элемента
        int itemsCount = Math.Max(1, (targetSize - 500) / estimatedItemSize);

        var items = (JsonArray)obj["items"]!;
        for (int i = 0; i < itemsCount; i++)
        {
            items.Add(new JsonObject
            {
                ["id"] = Guid.NewGuid().ToString(),
                ["name"] = $"Item_{i}",
                ["value"] = _random.NextDouble() * 1000,
                ["timestamp"] = DateTime.UtcNow.AddSeconds(-_random.Next(0, 86400)).ToString("O")
            });
        }

        return obj;
    }

    private int GetExponentialSize()
    {
        double u = _random.NextDouble();
        double size = -Math.Log(1 - u) / Lambda;

        if (size < MinSize) size = MinSize;
        if (size > MaxSize) size = MaxSize;

        return (int)size;
    }

    public void Dispose() { /* Нет неуправляемых ресурсов */ }
}