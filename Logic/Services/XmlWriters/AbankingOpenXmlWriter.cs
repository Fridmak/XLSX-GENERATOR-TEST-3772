using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

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
            SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var accessors = GetAccessors<T>(columns);

        uint sheetId = 1;
        uint rowIndex = 1;
        long sheetBytes = 0; // нужен для предотвращения ошибки Stream was too long при переносе в zip xml документа

        var worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
        var writer = CreateSheetWriter(worksheetPart);

        WriteHeader(writer, columns);
        rowIndex = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            // 1. Получаем строки ячеек (БЕЗ чанков)
            var values = GetRowValues(row, accessors, columns, out int maxChunks);

            // 2. Проверка лимитов
            if (rowIndex + maxChunks > XmlConstants.ExcelMaxRows)
            {
                CloseSheet(writer, worksheetPart);
                return;
            }

            // 3. Пишем строки
            for (int chunk = 0; chunk < maxChunks; chunk++)
            {
                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int c = 0; c < columns.Count; c++)
                {
                    WriteCellChunk(
                        writer,
                        values[c],
                        chunk);
                }

                writer.WriteEndElement(); // Row
                rowIndex++;
            }

            if (rowIndex % 10000 == 0)
                await Task.Yield();
        }


        CloseSheet(writer, worksheetPart);
    }

    // ---------------- helpers ----------------

    private static string[] GetRowValues<T>(
        T row,
        Func<object, object?>[] accessors,
        IReadOnlyList<ExcelColumn> columns,
        out int maxChunks)
    {
        var result = new string[columns.Count];
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

            result[i] = text;

            int chunks = (text.Length + XmlConstants.MaxCellTextLength - 1)
                         / XmlConstants.MaxCellTextLength;

            if (chunks > maxChunks)
                maxChunks = chunks;
        }

        return result;
    }

    private static void WriteCellChunk(
        OpenXmlWriter writer,
        string text,
        int chunkIndex)
    {
        int max = XmlConstants.MaxCellTextLength;
        int offset = chunkIndex * max;

        if (offset >= text.Length)
        {
            writer.WriteStartElement(new Cell());
            writer.WriteEndElement();
            return;
        }

        int len = Math.Min(max, text.Length - offset);

        writer.WriteStartElement(new Cell
        {
            DataType = CellValues.InlineString
        });

        writer.WriteStartElement(new InlineString());
        writer.WriteStartElement(new Text());

        const int subChunk = 4096; // 🔑 ключевая строка
        for (int i = 0; i < len; i += subChunk)
        {
            int partLen = Math.Min(subChunk, len - i);
            writer.WriteString(text.AsSpan(offset + i, partLen).ToString());
        }
        //writer.WriteString(text);

        writer.WriteEndElement(); // Text
        writer.WriteEndElement(); // InlineString
        writer.WriteEndElement(); // Cell
    }

    private static void WriteCell(
        OpenXmlWriter writer,
        int columnIndex,
        uint rowIndex,
        ReadOnlySpan<char> text)
    {
        writer.WriteStartElement(new Cell
        {
            DataType = CellValues.InlineString
        });

        writer.WriteStartElement(new InlineString());
        writer.WriteStartElement(new Text { Space = SpaceProcessingModeValues.Preserve });

        writer.WriteString(text.ToString());

        writer.WriteEndElement(); // Text
        writer.WriteEndElement(); // InlineString
        writer.WriteEndElement(); // Cell
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
}

