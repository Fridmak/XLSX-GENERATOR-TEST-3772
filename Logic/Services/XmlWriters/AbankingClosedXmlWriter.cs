using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Models;
using ClosedXML.Excel;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;

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
        int currentRow = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            int maxRowOffset = 1; // Максимальное количество строк, которое займёт этот объект

            // Собираем для каждой колонки куски текста
            var chunksPerColumn = new string[columns.Count][];
            for (int c = 0; c < columns.Count; c++)
            {
                object? value = accessors[c](row!);

                // Преобразуем объекты в безопасные строки
                string text;
                if (value == null)
                    text = string.Empty;
                else if (value is string s)
                    text = s;
                else
                {
                    try { text = JsonSerializer.Serialize(value); }
                    catch { text = value.ToString() ?? string.Empty; }
                }

                // Разбиваем на куски по лимиту
                chunksPerColumn[c] = SplitIntoChunks(text, XmlConstants.MaxCellTextLength);
                if (chunksPerColumn[c].Length > maxRowOffset)
                    maxRowOffset = chunksPerColumn[c].Length;
            }

            // Записываем куски по всем колонкам, выравнивая по строкам
            for (int offset = 0; offset < maxRowOffset; offset++)
            {
                // Если превысили лимит строк Excel, создаём новый лист
                if (currentRow > XmlConstants.ExcelMaxRows)
                {
                    sheetId++;
                    ws = workbook.AddWorksheet(GetSheetName(sheetId));
                    WriteHeader(ws, columns);
                    currentRow = 2;
                }

                for (int c = 0; c < columns.Count; c++)
                {
                    string chunkText = offset < chunksPerColumn[c].Length ? chunksPerColumn[c][offset] : string.Empty;
                    ws.Cell(currentRow, c + 1).SetValue(chunkText);
                }

                currentRow++;
            }

            // Чтобы UI/CPU не блокировался при огромных данных
            if (currentRow % 1000 == 0)
                await Task.Yield();
        }

        workbook.SaveAs(output);
    }

    // Разделение текста на куски
    private static string[] SplitIntoChunks(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text))
            return new[] { string.Empty };

        var chunks = new List<string>();
        int start = 0;
        while (start < text.Length)
        {
            int len = Math.Min(maxLength, text.Length - start);
            chunks.Add(text.Substring(start, len));
            start += len;
        }
        return chunks.ToArray();
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
