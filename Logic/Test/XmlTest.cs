using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using Analitics6400.Logic.Test.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Text.Json.Nodes;
using System.Xml;

namespace Analitics6400.Logic.Test;

public sealed class XmlTest<T> : IXmlTest where T : IXmlWriter
{
    private readonly IXmlWriter _writer;
    private readonly IDocumentProvider _documentProvider;
    private readonly ILogger<XmlTest<T>> _logger;
    public string Name => nameof(T); 

    public XmlTest(IEnumerable<IXmlWriter> writer, IDocumentProvider documentProvider, ILogger<XmlTest<T>> logger)
    {
        _writer = writer.Where(x => x is T).First();
        _documentProvider = documentProvider;
        _logger = logger;
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
        //var columns = new List<ExcelColumn>
        //{
        //    new("Id", typeof(Guid)),
        //    new("DocumentSchemaId", typeof(Guid?)),
        //    new("Published", typeof(DateTime?)),
        //    new("IsArchived", typeof(bool)),
        //    new("Version", typeof(double)),
        //    new("IsCanForValidate", typeof(bool)),
        //    new("ChangedDateUtc", typeof(DateTime?))
        //};

        _logger.LogInformation("Начало генерации CSV");
        var stopwatch = Stopwatch.StartNew();

        int rowCount = 0;
        try
        {
            var rows = _documentProvider.GetDocumentsAsync(limit: null, ct);

            var fileName = $"DocumentsReport_{DateTime.UtcNow:yyyyMMdd_HHmmss}.{_writer.Extension}";
            await using var fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None, bufferSize: 1024 * 1024);

            _logger.LogInformation($"Вызов генератора {nameof(T)} для {nameof(fileName)}: {fileName}");

            async IAsyncEnumerable<TItem> CountRows<TItem>(IAsyncEnumerable<TItem> source)
            {
                await foreach (var row in source.WithCancellation(ct))
                {
                    rowCount++;
                    if (rowCount % 1000 == 0)
                        _logger.LogInformation($"Обработано строк: {rowCount}");

                    yield return row;
                }
            }

            await _writer.GenerateAsync(CountRows(rows), columns, fs, ct);

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