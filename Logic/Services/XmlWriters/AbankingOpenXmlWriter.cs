using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingOpenXmlWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();
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
        if (columns == null || columns.Count == 0)
            throw new ArgumentException("Columns cannot be empty", nameof(columns));

        using var document = SpreadsheetDocument.Create(
            output,
            SpreadsheetDocumentType.Workbook,
            autoSave: true);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;
        uint rowIndex = 1;

        var worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
        var writer = CreateSheetWriter(worksheetPart);

        WriteHeader(writer, columns);
        rowIndex = 2;

        var accessors = GetAccessors<T>(columns);

        await foreach (var row in rows.WithCancellation(ct))
        {
            var columnChunks = new List<string[]>(columns.Count);
            int maxChunks = 1;

            for (int c = 0; c < columns.Count; c++)
            {
                var value = accessors[c](row);
                string text = value switch
                {
                    null => string.Empty,
                    string s => s,
                    _ => JsonSerializer.Serialize(value)
                };

                var chunks = SplitByExcelLimit(text).ToArray();
                columnChunks.Add(chunks);

                if (chunks.Length > maxChunks)
                    maxChunks = chunks.Length;
            }

            if (rowIndex + (uint)(maxChunks - 1) > XmlConstants.ExcelMaxRows)
            {
                CloseSheet(writer);
                sheetId++;
                rowIndex = 1;
                worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
                writer = CreateSheetWriter(worksheetPart);
                WriteHeader(writer, columns);
                rowIndex = 2;
            }

            for (int chunkIdx = 0; chunkIdx < maxChunks; chunkIdx++)
            {
                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int c = 0; c < columns.Count; c++)
                {
                    var chunks = columnChunks[c];
                    var value = chunkIdx < chunks.Length ? chunks[chunkIdx] : string.Empty;
                    WriteCellInline(writer, c + 1, rowIndex, value);
                }

                writer.WriteEndElement(); // Row
                rowIndex++;
            }

            columnChunks.Clear();

            if (rowIndex % 1000 == 0)
                await Task.Yield();
        }

        CloseSheet(writer);
        workbookPart.Workbook.Save();
    }

    private static IEnumerable<string> SplitByExcelLimit(string text)
    {
        if (string.IsNullOrEmpty(text))
            yield return string.Empty;

        const int MaxCellTextLength = 32_767;

        for (int i = 0; i < text.Length; i += MaxCellTextLength)
        {
            int length = Math.Min(MaxCellTextLength, text.Length - i);
            yield return text.Substring(i, length);
        }
    }

    private static void WriteCellInline(OpenXmlWriter writer, int columnIndex, uint rowIndex, string text)
    {
        writer.WriteStartElement(new Cell
        {
            CellReference = GetColumnName(columnIndex) + rowIndex,
            DataType = CellValues.InlineString
        });

        writer.WriteStartElement(new InlineString());
        writer.WriteElement(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        writer.WriteEndElement(); // InlineString

        writer.WriteEndElement(); // Cell
    }

    private static WorksheetPart CreateWorksheetPart(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
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

    private static void CloseSheet(OpenXmlWriter writer)
    {
        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
        writer.Close();
    }

    private static void WriteHeader(OpenXmlWriter writer, IReadOnlyList<ExcelColumn> columns)
    {
        writer.WriteStartElement(new Row { RowIndex = 1 });

        for (int i = 0; i < columns.Count; i++)
            WriteCellInline(writer, i + 1, 1, columns[i].Header ?? string.Empty);

        writer.WriteEndElement(); // Row
    }

    private static string GetColumnName(int columnIndex)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string result = "";

        while (columnIndex > 0)
        {
            int rem = (columnIndex - 1) % 26;
            result = letters[rem] + result;
            columnIndex = (columnIndex - 1) / 26;
        }

        return result;
    }

    private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
    {
        var key = (typeof(T), string.Join("|", columns.Select(c => c.Name)));

        return _propertyAccessors.GetOrAdd(key, _ =>
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
