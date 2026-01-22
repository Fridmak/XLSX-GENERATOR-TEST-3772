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
        private static readonly ConcurrentDictionary<Type, Func<object, object?>[]> _propertyAccessors = new();

        public async Task<byte[]> GenerateAsync<T>(
            IAsyncEnumerable<T> rows,
            IReadOnlyList<ExcelColumn> columns,
            CancellationToken ct = default)
        {
            if (columns.Count == 0)
                throw new ArgumentException("Columns cannot be empty", nameof(columns));

            await using var stream = new MemoryStream(32 * 1024 * 1024);

            byte[] result;

            using (var document = SpreadsheetDocument.Create(
                       stream,
                       SpreadsheetDocumentType.Workbook,
                       autoSave: false))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    // Заголовки
                    WriteHeader(writer, columns);

                    // Кэш делегатов
                    var accessors = GetAccessors<T>(columns);

                    // Потоковая запись данных
                    await foreach (var row in rows.WithCancellation(ct))
                    {
                        writer.WriteStartElement(new Row());

                        for (int i = 0; i < columns.Count; i++)
                        {
                            var value = accessors[i](row!);
                            WriteCell(writer, value);
                        }

                        writer.WriteEndElement(); // Row
                    }

                    writer.WriteEndElement(); // SheetData
                    writer.WriteEndElement(); // Worksheet
                }

                workbookPart.Workbook.AppendChild(new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }));

                workbookPart.Workbook.Save();
            }

            result = stream.ToArray();
            return result;
        }

        private static void WriteHeader(OpenXmlWriter writer, IReadOnlyList<ExcelColumn> columns)
        {
            writer.WriteStartElement(new Row());

            foreach (var col in columns)
            {
                writer.WriteElement(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(col.Header)
                });
            }

            writer.WriteEndElement();
        }

        private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
        {
            return _propertyAccessors.GetOrAdd(typeof(T), _ =>
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
                        CellValue = new CellValue(
                            dt.ToOADate().ToString(CultureInfo.InvariantCulture))
                    });
                    break;

                case sbyte or byte or short or ushort or int or uint or long or ulong
                   or float or double or decimal:
                    writer.WriteElement(new Cell
                    {
                        DataType = CellValues.Number,
                        CellValue = new CellValue(
                            Convert.ToString(value, CultureInfo.InvariantCulture)!)
                    });
                    break;

                default:
                    WriteString(writer, value.ToString()!);
                    break;
            }
        }

        private static void WriteString(OpenXmlWriter writer, string value)
        {
            writer.WriteElement(new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(value)
            });
        }
    }
}
