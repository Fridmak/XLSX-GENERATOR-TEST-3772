using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingOpenXmlWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type, string), Func<object, object?>[]> _accessors = new();
    private readonly ILogger<AbankingOpenXmlWriter> _logger;

    public string Extension => "xlsx";

    public AbankingOpenXmlWriter(ILogger<AbankingOpenXmlWriter> logger)
    {
        _logger = logger;
    }

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        CancellationToken ct = default)
    {
        using var document = SpreadsheetDocument.Create(
            output,
            SpreadsheetDocumentType.Workbook,
            autoSave: true);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var accessors = GetAccessors<T>(columns);

        uint sheetId = 1;
        uint rowIndex = 1;
        long sheetSize = 0; // нужен для предотвращения ошибки Stream was too long при переносе в zip xml документа

        var worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
        var writer = CreateSheetWriter(worksheetPart);

        WriteHeader(writer, columns);
        rowIndex = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            var split = SplitRow(row, accessors, columns, out int maxChunks);
            long estimated = EstimateRowSize(split);

            if (sheetSize + estimated > XmlConstants.MaxSheetBytes ||
                rowIndex + maxChunks > XmlConstants.ExcelMaxRows)
            {
                CloseSheet(writer, worksheetPart);

                sheetId++;
                rowIndex = 1;
                sheetSize = 0;

                worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
                writer = CreateSheetWriter(worksheetPart);
                WriteHeader(writer, columns);
                rowIndex = 2;
            }

            for (int chunk = 0; chunk < maxChunks; chunk++)
            {
                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int c = 0; c < columns.Count; c++)
                {
                    var text = chunk < split[c].Length
                        ? split[c][chunk]
                        : ReadOnlyMemory<char>.Empty;

                    WriteCell(writer, c + 1, rowIndex, text.Span);
                }

                writer.WriteEndElement(); // Row
                rowIndex++;
            }

            sheetSize += estimated;

            if (rowIndex % 1000 == 0)
            {
                await Task.Yield();
            }
        }

        CloseSheet(writer, worksheetPart);
    }

    // ---------------- helpers ----------------

    private static ReadOnlyMemory<char>[][] SplitRow<T>(
        T row,
        Func<object, object?>[] accessors,
        IReadOnlyList<ExcelColumn> columns,
        out int maxChunks)
    {
        var result = new ReadOnlyMemory<char>[columns.Count][];
        maxChunks = 1;

        for (int i = 0; i < columns.Count; i++)
        {
            var value = accessors[i](row);
            var text = value switch
            {
                null => string.Empty,
                string s => s,
                _ => JsonSerializer.Serialize(value, new JsonSerializerOptions { WriteIndented = false })
            };

            var chunks = SplitText(text).ToArray();
            result[i] = chunks;

            if (chunks.Length > maxChunks)
                maxChunks = chunks.Length;
        }

        return result;
    }

    private static IEnumerable<ReadOnlyMemory<char>> SplitText(string text)
    {
        for (int i = 0; i < text.Length; i += XmlConstants.MaxCellTextLength)
        {
            int len = Math.Min(XmlConstants.MaxCellTextLength, text.Length - i);
            yield return text.AsMemory(i, len);
        }
    }

    private static void WriteCell(
        OpenXmlWriter writer,
        int columnIndex,
        uint rowIndex,
        ReadOnlySpan<char> text)
    {
        writer.WriteStartElement(new Cell
        {
            CellReference = GetColumnName(columnIndex) + rowIndex,
            DataType = CellValues.InlineString
        });

        writer.WriteStartElement(new InlineString());
        writer.WriteStartElement(new Text { Space = SpaceProcessingModeValues.Preserve });

        for (int i = 0; i < text.Length; i += XmlConstants.TextChunkSize)
        {
            int len = Math.Min(XmlConstants.TextChunkSize, text.Length - i);
            writer.WriteString(new string(text.Slice(i, len)));
        }

        writer.WriteEndElement(); // Text
        writer.WriteEndElement(); // InlineString
        writer.WriteEndElement(); // Cell
    }

    private static long EstimateRowSize(ReadOnlyMemory<char>[][] chunks)
    {
        long size = 64;

        foreach (var col in chunks)
            foreach (var chunk in col)
                size += 128 + chunk.Length;

        return size;
    }

    private static WorksheetPart CreateWorksheetPart(
        WorkbookPart workbookPart,
        Sheets sheets,
        uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = $"Sheet{sheetId}"
        });

        return worksheetPart;
    }

    private static OpenXmlWriter CreateSheetWriter(WorksheetPart worksheetPart)
    {
        var writer = OpenXmlWriter.Create(worksheetPart);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());
        return writer;
    }

    private static void CloseSheet(OpenXmlWriter writer, WorksheetPart worksheetPart)
    {
        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
        writer.Close();
    }

    private static void WriteHeader(OpenXmlWriter writer, IReadOnlyList<ExcelColumn> columns)
    {
        writer.WriteStartElement(new Row { RowIndex = 1 });

        for (int i = 0; i < columns.Count; i++)
            WriteCell(writer, i + 1, 1, (columns[i].Header ?? "").AsSpan());

        writer.WriteEndElement();
    }

    private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
    {
        var key = (typeof(T), string.Join("|", columns.Select(c => c.Name)));

        return _accessors.GetOrAdd(key, _ =>
        {
            var props = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .ToDictionary(p => p.Name, StringComparer.Ordinal);

            return columns.Select(col =>
                props.TryGetValue(col.Name, out var prop)
                    ? (Func<object, object?>)prop.GetValue
                    : static _ => null
            ).ToArray();
        });
    }

    private static string GetColumnName(int columnIndex)
    {
        string result = "";
        while (columnIndex > 0)
        {
            columnIndex--;
            result = (char)('A' + columnIndex % 26) + result;
            columnIndex /= 26;
        }
        return result;
    }
}
