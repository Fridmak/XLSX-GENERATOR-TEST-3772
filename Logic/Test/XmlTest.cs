using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using Analitics6400.Logic.Test.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml;

namespace Analitics6400.Logic.Test;

public sealed class XmlTest<T> : IXmlTest where T : IXmlWriter
{
    private readonly IXmlWriter _writer;
    private readonly IDocumentProvider _documentProvider;
    private readonly ILogger<XmlTest<T>> _logger;
    private readonly XmlTestSettings _settings;
    public string Name => nameof(T); 

    public XmlTest(IEnumerable<IXmlWriter> writer, IDocumentProvider documentProvider, ILogger<XmlTest<T>> logger, IOptions<XmlTestSettings> options)
    {
        _writer = writer.Where(x => x is T).First();
        _documentProvider = documentProvider;
        _logger = logger;
        _settings = options.Value;
    }

    #region Разогрев

    private async Task StartUp()
    {
        var columns = new List<ExcelColumn>
        {
            new("Id", typeof(Guid)),
            new("DocumentSchemaId", typeof(Guid?)),
            new("Published", typeof(DateTime?)),
            new("IsArchived", typeof(bool)),
            new("Version", typeof(double)),
            new("JsonData", typeof(string)),
            new("IsCanForValidate", typeof(bool)),
            new("ChangedDateUtc", typeof(DateTime?))
        };

        var rows = _documentProvider.GetDocumentsAsync(limit: 2000);

        _logger.LogInformation("Разогрев генератора без записи в файл...");

        var nullStream = new NullStream();
        var stopwatch = Stopwatch.StartNew();

        await _writer.GenerateAsync(rows, columns, nullStream);

        stopwatch.Stop();
        _logger.LogInformation($"Разогрев завершен. Время: {stopwatch.Elapsed.TotalSeconds:F2} секунд");
    }

    #endregion

    public async Task RunAsync(CancellationToken ct = default)
    {
        //await StartUp(); // Не для openxml. Надо для него переделать

        var columns = new List<ExcelColumn>
        {
            new("Id", typeof(Guid)),
            new("DocumentSchemaId", typeof(Guid?)),
            new("Published", typeof(DateTime?)),
            new("IsArchived", typeof(bool)),
            new("Version", typeof(double)),
            new("JsonData", typeof(JsonObject)),
            new("IsCanForValidate", typeof(bool)),
            new("ChangedDateUtc", typeof(DateTime?))
        };

        _logger.LogInformation("Начало генерации CSV");
        var stopwatch = Stopwatch.StartNew();

        int rowCount = 0;
        try
        {
            var rows = _documentProvider.GetDocumentsAsync(_settings.DocumentsLimit, ct);

            var fileName = $"DocumentsReport_{DateTime.UtcNow:yyyyMMdd_HHmmss}.{_writer.Extension}";
            await using var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize:  64 * 1024);

            _logger.LogInformation($"Вызов генератора {_writer.ToString()} для {nameof(fileName)}: {fileName}");

            async IAsyncEnumerable<DocumentDtoModel> CountRows(
                IAsyncEnumerable<DocumentDtoModel> source,
                CancellationToken ct = default)
            {
                var rowCount = 0;
                var maxSize = 0L;
                var totalSize = 0L;
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = false
                };

                await foreach (var row in source.WithCancellation(ct))
                {
                    long currentRowSize = 0;
                    if (row.JsonData != null)
                    {
                        var countingStream = new CountingStream();
                        using var writer = new Utf8JsonWriter(countingStream, new JsonWriterOptions { Indented = false });
                        row.JsonData.WriteTo(writer);
                        writer.Flush();

                        currentRowSize = countingStream.LengthWritten;
                    }

                    maxSize = Math.Max(maxSize, currentRowSize);
                    totalSize += currentRowSize;
                    rowCount++;

                    if (rowCount % 1000 == 0)
                    {
                        var meanSize = rowCount > 0 ? totalSize / rowCount : 0;

                        _logger.LogInformation($"Обработано строк: {rowCount}");
                        _logger.LogInformation($"Максимальный размер JSON: {maxSize} bytes");
                        _logger.LogInformation($"Средний размер JSON: {meanSize} bytes");

                        _logger.LogDebug($"Текущий размер: {currentRowSize} bytes");
                    }

                    yield return row;
                }

                // Финальная статистика
                if (rowCount > 0)
                {
                    var finalMeanSize = totalSize / rowCount;

                    _logger.LogInformation($"=== ФИНАЛЬНАЯ СТАТИСТИКА ===");
                    _logger.LogInformation($"Всего строк: {rowCount}");
                    _logger.LogInformation($"Максимальный размер: {maxSize} bytes ({maxSize / 1024.0:F2} KB)");
                    _logger.LogInformation($"Средний размер: {finalMeanSize} bytes ({finalMeanSize / 1024.0:F2} KB)");
                    _logger.LogInformation($"Общий объем: {totalSize} bytes ({totalSize / 1024.0 / 1024.0:F2} MB)");
                }
            }

            await _writer.GenerateAsync(CountRows(rows), columns, fs, ct);

            _logger.LogWarning(Process.GetCurrentProcess().WorkingSet64.ToString());

            _logger.LogWarning(GC.GetTotalMemory(false).ToString());

            _logger.LogWarning(GC.GetTotalMemory(true).ToString());


            stopwatch.Stop();
            _logger.LogInformation($"Файл сгенерирован: {fileName}");
            _logger.LogInformation($"Всего обработано строк: {rowCount}");
            _logger.LogInformation($"Время выполнения: {stopwatch.Elapsed.TotalSeconds:F2} секунд");
        }
        catch (OperationCanceledException)
        {
            stopwatch.Stop();
            _logger.LogWarning("Генерация была отменена пользователем");
            _logger.LogInformation($"Обработано строк до отмены: {rowCount}");
            _logger.LogInformation($"Время выполнения до отмены: {stopwatch.Elapsed.TotalSeconds:F2} секунд");
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex, "Ошибка при генерации CSV");
            _logger.LogInformation($"Обработано строк до ошибки: {rowCount}");
            _logger.LogInformation($"Время выполнения до ошибки: {stopwatch.Elapsed.TotalSeconds:F2} секунд");
            throw;
        }
    }
}
