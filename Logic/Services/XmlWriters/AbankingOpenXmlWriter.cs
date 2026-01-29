using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;

namespace Analitics6400.Logic.Services.XmlWriters
{
    public sealed class AbankingOpenXmlWriter : IXmlWriter
    {
        private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();
        private readonly ILogger<AbankingOpenXmlWriter> _logger;

        public string Extension => ".xlsx";

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
            if (columns.Count == 0)
            {
                throw new ArgumentException("Columns cannot be empty", nameof(columns));
            }

            // Создаем документ
            using var document = SpreadsheetDocument.Create(
                output,
                SpreadsheetDocumentType.Workbook,
                autoSave: false);

            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            uint sheetId = 1;
            uint rowIndex = 2; // первая строка - заголовок

            // создаем первый лист
            var worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
            var writer = OpenXmlWriter.Create(worksheetPart);

            try
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                // пишем заголовок
                WriteHeader(writer, columns);

                var accessors = GetAccessors<T>(columns);

                await foreach (var row in rows.WithCancellation(ct))
                {
                    // если превышен лимит строк, создаем новый лист
                    if (rowIndex > XmlConstants.ExcelMaxRows)
                    {
                        // закрываем текущий лист
                        writer.WriteEndElement(); // SheetData
                        writer.WriteEndElement(); // Worksheet
                        writer.Close();
                        writer.Dispose();

                        sheetId++;
                        rowIndex = 2; // снова после заголовка

                        // создаем новый лист
                        worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
                        writer = OpenXmlWriter.Create(worksheetPart);
                        writer.WriteStartElement(new Worksheet());
                        writer.WriteStartElement(new SheetData());
                        WriteHeader(writer, columns); // новый заголовок
                    }

                    // пишем строку
                    writer.WriteStartElement(new Row { RowIndex = rowIndex });
                    for (int i = 0; i < columns.Count; i++)
                    {
                        WriteCell(writer, accessors[i](row!));
                    }
                    writer.WriteEndElement(); // Row
                    rowIndex++;
                }
            }
            finally
            {
                // закрываем последний лист
                writer.WriteEndElement(); // SheetData
                writer.WriteEndElement(); // Worksheet
                writer.Close();
                writer.Dispose();
            }

            workbookPart.Workbook.Save();
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

        private static void WriteHeader(OpenXmlWriter writer, IReadOnlyList<ExcelColumn> columns)
        {
            writer.WriteStartElement(new Row { RowIndex = 1 });
            foreach (var col in columns)
            {
                writer.WriteElement(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(col.Header)
                });
            }
            writer.WriteEndElement(); // Row
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

        private static void WriteCell(OpenXmlWriter writer, object? value)
        {
            if (value is null)
            {
                writer.WriteElement(new Cell());
                return;
            }

            switch (value)
            {
                case string s:
                    if (s.Length > XmlConstants.MaxCellTextLength)
                    {
                        int allowedLength = XmlConstants.MaxCellTextLength - XmlConstants.JsonOverflowNotice.Length;
                        if (allowedLength < 0) allowedLength = XmlConstants.MaxCellTextLength;
                        string truncated = s.Substring(0, allowedLength) + XmlConstants.JsonOverflowNotice;
                        WriteString(writer, truncated);
                    }
                    else
                        WriteString(writer, s);
                    break;

                case Guid g:
                    WriteString(writer, g.ToString());
                    break;

                case bool b:
                    writer.WriteElement(new Cell
                    {
                        DataType = CellValues.Boolean,
                        CellValue = new CellValue(b ? "1" : "0")
                    });
                    break;

                case DateTime dt:
                    writer.WriteElement(new Cell
                    {
                        DataType = CellValues.Number,
                        CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture))
                    });
                    break;

                case sbyte or byte or short or ushort or int or uint or long or ulong
                   or float or double or decimal:
                    writer.WriteElement(new Cell
                    {
                        DataType = CellValues.Number,
                        CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture)!)
                    });
                    break;

                default:
                    WriteString(writer, value.ToString()!);
                    break;
            }
        }

        private static void WriteString(OpenXmlWriter writer, string value)
        {
            writer.WriteStartElement(new Cell { DataType = CellValues.InlineString });
            writer.WriteStartElement(new InlineString());
            writer.WriteElement(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
            writer.WriteEndElement(); // InlineString
            writer.WriteEndElement(); // Cell
        }
    }
}
