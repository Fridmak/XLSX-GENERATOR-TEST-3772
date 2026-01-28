using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;

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
    private readonly byte[] _jsonTemplateBytes;
    private readonly byte[] _largeJsonBuffer;

    private const int MinSize = 10_000;       // 10 КБ
    private const int MaxSize = 1_000_000;    // 1 МБ
    private const double Lambda = 1.0 / 50_000; // средний размер ~50 КБ

    public MockDocumentDalProvider(int documentCount = 1000)
    {
        _documentCount = documentCount;
        _jsonTemplateBytes = Encoding.UTF8.GetBytes("{\"key\":\"value\"}");

        // Создаем один раз большой буфер максимального размера
        _largeJsonBuffer = ArrayPool<byte>.Shared.Rent(MaxSize);
        FillLargeBuffer();
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

            // Используем предзаполненный буфер
            string json = GetJsonString(jsonSize);

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow,
                IsArchived = false,
                Version = 1.0,
                JsonData = json,
                IsCanForValidate = true,
                ChangedDateUtc = DateTime.UtcNow
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            await Task.Yield(); // чтобы IAsyncEnumerable был асинхронным
        }
    }

    private void FillLargeBuffer()
    {
        int position = 0;
        while (position < MaxSize - _jsonTemplateBytes.Length)
        {
            Buffer.BlockCopy(_jsonTemplateBytes, 0, _largeJsonBuffer, position, _jsonTemplateBytes.Length);
            position += _jsonTemplateBytes.Length;
        }
    }

    private string GetJsonString(int size)
    {
        // Просто берем нужное количество байт из предзаполненного буфера
        return Encoding.UTF8.GetString(_largeJsonBuffer, 0, size);
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

    // Освобождаем ресурсы при необходимости
    public void Dispose()
    {
        ArrayPool<byte>.Shared.Return(_largeJsonBuffer);
    }
}