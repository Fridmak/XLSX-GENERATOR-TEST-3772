using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Models;
using ClosedXML.Excel;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingClosedXmlWriterForeign : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();
    private readonly ILogger<AbankingClosedXmlWriterForeign> _logger;

    public string Extension => ".xlsx";

    public AbankingClosedXmlWriterForeign(ILogger<AbankingClosedXmlWriterForeign> logger)
    {
        _logger = logger;
    }

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        CancellationToken ct = default)
    {
        if (columns.Count == 0)
            throw new ArgumentException("Columns cannot be empty", nameof(columns));

        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");
        WriteHeader(ws, columns);

        var accessors = GetAccessors<T>(columns);
        int currentRow = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            bool compense = false;
            int maxCellRowOffset = 0;

            for (int c = 0; c < columns.Count; c++)
            {
                object? value = accessors[c](row!);

                // Если это не строка, сериализуем JSON/объект
                if (value != null && value.GetType() != typeof(string))
                {
                    try
                    {
                        value = JsonSerializer.Serialize(value);
                    }
                    catch
                    {
                        value = value.ToString();
                    }
                }

                string text = value?.ToString() ?? string.Empty;

                int rowOffset = 0;

                // Разбиваем длинные строки на несколько ячеек вниз
                while (text.Length > 0)
                {
                    int chunkSize = Math.Min(XmlConstants.MaxCellTextLength, text.Length);
                    string chunk = text.Substring(0, chunkSize);

                    if (text.Length > chunkSize)
                        chunk += XmlConstants.JsonOverflowNotice;

                    ws.Cell(currentRow + rowOffset, c + 1).Value = chunk;

                    text = text.Length > chunkSize ? text.Substring(chunkSize) : string.Empty;
                    rowOffset++;
                    compense = rowOffset > 1;
                }

                if (rowOffset > maxCellRowOffset)
                    maxCellRowOffset = rowOffset;
            }

            // Смещаем строку на наибольший offset
            currentRow += Math.Max(1, maxCellRowOffset);

            if (compense)
            {
                currentRow--; // как в твоей логике
            }

            if (currentRow % 1000 == 0)
                await Task.Yield();
        }

        workbook.SaveAs(output);
    }

    private static void WriteHeader(IXLWorksheet ws, IReadOnlyList<ExcelColumn> columns)
    {
        for (int c = 0; c < columns.Count; c++)
            ws.Cell(1, c + 1).Value = columns[c].Header;
    }

    private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
    {
        var key = (typeof(T), ColumnsKey: string.Join("|", columns.Select(c => c.Name)));

        return _propertyAccessors.GetOrAdd(key, _ =>
        {
            var props = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .ToDictionary(p => p.Name, p => p);

            return columns.Select(col =>
                    props.TryGetValue(col.Name, out var prop)
                        ? (Func<object, object?>)(obj => prop.GetValue(obj))
                        : (_ => null))
                .ToArray();
        });
    }
}
