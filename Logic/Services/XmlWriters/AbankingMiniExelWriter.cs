using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingMiniExcelWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type, string), Func<object, object?>[]> _accessors = new();

    public string Extension => "xlsx";

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        string output,
        CancellationToken ct = default)
    {
        var accessors = GetAccessors<T>(columns);

        var config = new OpenXmlConfiguration
        {
            FastMode = true,
            AutoFilter = false,
            TableStyles = TableStyles.None,
            EnableSharedStringCache = true,
            SharedStringCacheSize = 50 * 1024 * 1024,
            IgnoreEmptyRows = true,
            EnableWriteNullValueCell = true,
            WriteEmptyStringAsNull = false
        };

        // MiniExcel работает с IEnumerable, конвертируем IAsyncEnumerable в IEnumerable батчами
        var ind = 0;
        var lastSheetName = $"Sheet{ind}";
        await foreach (var (batch, sheetName) in GetBatches(rows, columns, accessors, ct))
        {
            {
                // Первый батч - создаем файл
                if (output.Position == 0)
                {
                    await MiniExcel.SaveAsAsync(output, batch,
                        printHeader: true,
                        sheetName: sheetName,
                        excelType: ExcelType.XLSX,
                        configuration: config,
                        cancellationToken: ct);
                }
                else
                {
                    output.Seek(0, SeekOrigin.Begin);

                    var isNewSheet = lastSheetName != sheetName;

                    // Последующие батчи - добавляем
                    await MiniExcel.InsertAsync(output, batch,
                        sheetName: sheetName,
                        excelType: ExcelType.XLSX,
                        configuration: config,
                        printHeader: isNewSheet,
                        cancellationToken: ct);
                }
                ind++;

                await Task.Yield();
            }
        }
    }

    private async IAsyncEnumerable<(List<Dictionary<string, object>> Batch, string SheetName)> GetBatches<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Func<object, object?>[] accessors,
        [EnumeratorCancellation] CancellationToken ct)
    {

        var sheetIndex = 0; 
        var batch = new List<Dictionary<string, object>>();
        var headerNames = columns.Select(c => c.Header ?? c.Name).ToArray();

        await foreach (var row in rows.WithCancellation(ct))
        {
            // Получаем значения всех колонок
            var columnValues = new string[columns.Count];
            var columnChunks = new int[columns.Count];
            var maxChunks = 1;
            var stringsUsed = 1;

            for (int i = 0; i < columns.Count; i++)
            {
                var value = accessors[i](row);
                var text = value switch
                {
                    null => string.Empty,
                    string s => s,
                    _ => JsonSerializer.Serialize(value, JsonSerializerOptions.Default)
                };

                columnValues[i] = text;
                columnChunks[i] = (text.Length + XmlConstants.MaxCellTextLength - 1) / XmlConstants.MaxCellTextLength;

                if (columnChunks[i] > maxChunks)
                {
                    maxChunks = columnChunks[i];
                }
            }

            // Генерируем строки
            for (int chunk = 0; chunk < maxChunks; chunk++)
            {
                var rowData = new Dictionary<string, object>();

                for (int i = 0; i < columns.Count; i++)
                {
                    var header = headerNames[i];
                    var text = columnValues[i];

                    if (chunk == 0)
                    {
                        // Первая строка - все значения (обрезанные до 32К если нужно)
                        rowData[header] = text.Length > XmlConstants.MaxCellTextLength
                            ? text.Substring(0, XmlConstants.MaxCellTextLength)
                            : text;
                    }
                    else
                    {
                        // Для последующих строк - ТОЛЬКО если есть продолжение текста
                        int offset = chunk * XmlConstants.MaxCellTextLength;
                        if (offset < text.Length)
                        {
                            // Эта колонка имеет продолжение - пишем его
                            int length = Math.Min(XmlConstants.MaxCellTextLength, text.Length - offset);
                            rowData[header] = text.Substring(offset, length);
                        }
                        else
                        {
                            // Все остальные колонки - пустые
                            rowData[header] = string.Empty;
                        }
                    }
                }

                batch.Add(rowData);
            }
            stringsUsed += maxChunks;

            if (stringsUsed >= XmlConstants.ExcelMaxRows - 100)
            {
                sheetIndex += 1;
                stringsUsed = 0;
                yield return (batch, $"Sheet{sheetIndex}");
                batch = new List<Dictionary<string, object>>();
            }

            if (batch.Count >= XmlConstants.BatchSize)
            {
                stringsUsed = 0;
                yield return (batch, $"Sheet{sheetIndex}");
                batch = new List<Dictionary<string, object>>();
            }
        }

        if (batch.Count > 0)
        {
            yield return (batch, $"Sheet{sheetIndex}");
        }
    }

    private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
    {
        var key = (typeof(T), string.Join("|", columns.Select(c => c.Name)));

        return _accessors.GetOrAdd(key, _ =>
        {
            var props = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

            return columns.Select(col =>
                props.TryGetValue(col.Name, out var prop)
                    ? (Func<object, object?>)(obj => prop.GetValue(obj))
                    : static _ => null
            ).ToArray();
        });
    }
}