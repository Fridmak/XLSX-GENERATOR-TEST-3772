using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Models;
using ClosedXML.Excel;
using System.Collections.Concurrent;
using System.Reflection;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingClosedXmlWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();
    private readonly ILogger<AbankingClosedXmlWriter> _logger;

    public string Extension => ".xlsx";

    public AbankingClosedXmlWriter(ILogger<AbankingClosedXmlWriter> logger)
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
        int sheetId = 1;
        var ws = workbook.AddWorksheet(GetSheetName(sheetId));
        WriteHeader(ws, columns);

        var accessors = GetAccessors<T>(columns);
        int rowIndex = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            int maxRowInObject = rowIndex;

            for (int c = 0; c < columns.Count; c++)
            {
                var value = accessors[c](row!);
                var text = value?.ToString() ?? string.Empty;

                WriteValueInChunks(ws, c + 1, text, ref maxRowInObject, ref sheetId, workbook, columns);
            }

            rowIndex = maxRowInObject;

            if (rowIndex % 1000 == 0)
                await Task.Yield();
        }

        workbook.SaveAs(output);
    }

    private static void WriteValueInChunks(IXLWorksheet ws, int column, string text, ref int rowIndex, ref int sheetId, XLWorkbook workbook, IReadOnlyList<ExcelColumn> columns)
    {
        int maxLen = XmlConstants.MaxCellTextLength;
        int start = 0;

        while (start < text.Length)
        {
            int length = Math.Max(0, Math.Min(maxLen, text.Length - start));
            string chunk = text.Substring(start, length);

            if (rowIndex > XmlConstants.ExcelMaxRows)
            {
                sheetId++;
                ws = workbook.AddWorksheet(GetSheetName(sheetId));
                WriteHeader(ws, columns);
                rowIndex = 2;
            }

            ws.Cell(rowIndex, column).SetValue(chunk);
            rowIndex++;
            start += length;
        }
    }

    private static string GetSheetName(int sheetId) => $"Sheet{sheetId}";

    private static void WriteHeader(IXLWorksheet ws, IReadOnlyList<ExcelColumn> columns)
    {
        for (int c = 0; c < columns.Count; c++)
        {
            ws.Cell(1, c + 1).SetValue(columns[c].Header);
        }
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
