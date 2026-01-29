using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;
using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;

namespace Analitics6400.Dal.Services;

/// <summary>
/// Mock провайдер документов для тестов.
/// Возвращает одинаковые документы с JSONData случайного размера с экспоненциальным распределением.
/// Большие JSON (~1 МБ) встречаются редко.
/// </summary>
public sealed class MockDocumentDalProvider : IDocumentProvider
{
    private readonly int _documentCount;
    private readonly Random _random = new();
    private readonly JsonObject _jsonTemplate;
    private readonly string _largeJsonString;

    private const int MinSize = 10_000;       // 10 КБ
    private const int MaxSize = 1_000_000;    // 1 МБ
    private const double Lambda = 1.0 / 50_000; // средний размер ~50 КБ

    public MockDocumentDalProvider(int documentCount = 1000)
    {
        _documentCount = documentCount;

        // Создаем шаблон JsonObject
        _jsonTemplate = new JsonObject
        {
            ["id"] = Guid.NewGuid().ToString(),
            ["timestamp"] = DateTime.UtcNow.ToString("O"),
            ["data"] = new JsonObject
            {
                ["value1"] = "test",
                ["value2"] = 123,
                ["nested"] = new JsonObject
                {
                    ["field"] = "nestedValue"
                }
            }
        };

        // Создаем большую строку для копирования
        _largeJsonString = CreateLargeJsonString();
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
        int? limit = null,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        int returned = 0;

        for (int i = 0; i < _documentCount; i++)
        {
            ct.ThrowIfCancellationRequested();

            // экспоненциальное распределение
            int jsonSize = GetExponentialSize();

            // Создаем JsonObject нужного размера
            JsonObject jsonObject = CreateJsonObject(jsonSize);

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow,
                IsArchived = false,
                Version = 1.0,
                JsonData = jsonObject,
                IsCanForValidate = true,
                ChangedDateUtc = DateTime.UtcNow
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            await Task.Yield(); // чтобы IAsyncEnumerable был асинхронным
        }
    }

    private JsonObject CreateJsonObject(int targetSize)
    {
        var result = new JsonObject();

        // Добавляем базовые поля
        result["documentId"] = Guid.NewGuid().ToString();
        result["createdAt"] = DateTime.UtcNow.ToString("O");
        result["index"] = _random.Next(1, 10000);

        // Создаем массив данных для увеличения размера
        var dataArray = new JsonArray();
        result["items"] = dataArray;

        // Рассчитываем сколько элементов нужно добавить
        // Средний размер одного элемента ~200 байт
        int itemsNeeded = (targetSize - 500) / 200;
        if (itemsNeeded < 1) itemsNeeded = 1;

        for (int i = 0; i < itemsNeeded; i++)
        {
            var item = new JsonObject
            {
                ["id"] = Guid.NewGuid().ToString(),
                ["name"] = $"Item_{i}",
                ["value"] = _random.NextDouble() * 1000,
                ["timestamp"] = DateTime.UtcNow.AddSeconds(-_random.Next(0, 86400)).ToString("O"),
                ["metadata"] = new JsonObject
                {
                    ["category"] = $"Category{_random.Next(1, 10)}",
                    ["tags"] = new JsonArray
                    {
                        $"tag{_random.Next(1, 10)}",
                        $"tag{_random.Next(11, 20)}",
                        $"tag{_random.Next(21, 30)}"
                    }
                }
            };

            dataArray.Add(item);

            // Проверяем текущий размер
            if (GetJsonSize(result) >= targetSize)
                break;
        }

        return result;
    }

    private int GetJsonSize(JsonObject jsonObject)
    {
        // Простой способ оценить размер - сериализовать в строку
        return jsonObject.ToJsonString().Length;
    }

    private string CreateLargeJsonString()
    {
        // Создаем очень большой JsonObject и сериализуем его
        var largeObject = new JsonObject();
        var largeArray = new JsonArray();

        for (int i = 0; i < 1000; i++)
        {
            var item = new JsonObject
            {
                ["index"] = i,
                ["data"] = new string('x', 1000), // строка из 1000 символов
                ["values"] = new JsonArray
                {
                    _random.Next(),
                    _random.NextDouble(),
                    _random.NextDouble(),
                    _random.NextDouble()
                }
            };
            largeArray.Add(item);
        }

        largeObject["data"] = largeArray;
        return largeObject.ToJsonString();
    }

    private int GetExponentialSize()
    {
        // Генератор экспоненциального распределения
        double u = _random.NextDouble();
        double size = -Math.Log(1 - u) / Lambda; // формула экспоненциального распределения

        // Ограничиваем размер [MinSize, MaxSize]
        if (size < MinSize) size = MinSize;
        if (size > MaxSize) size = MaxSize;

        return (int)size;
    }

    // Если нужен более контролируемый подход с ресайклингом буфера
    public void Dispose()
    {
        // Ничего не делаем, так как нет неуправляемых ресурсов
    }
}