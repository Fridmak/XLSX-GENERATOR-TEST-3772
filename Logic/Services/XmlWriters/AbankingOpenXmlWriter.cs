using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Constants;
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

        private sealed record CellPlan(object? Value, string? Text, bool IsLongText, int ChunkCount);

        public async Task GenerateAsync<T>(
            IAsyncEnumerable<T> rows,
            IReadOnlyList<ExcelColumn> columns,
            Stream output,
            CancellationToken ct = default)
        {
            if (columns.Count == 0)
                throw new ArgumentException("Columns cannot be empty", nameof(columns));

            using (var document = SpreadsheetDocument.Create(
                       output,
                       SpreadsheetDocumentType.Workbook,
                       autoSave: false))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                using var writer = OpenXmlWriter.Create(worksheetPart);

                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                WriteHeader(writer, columns);

                var accessors = GetAccessors<T>(columns);

                uint mainRowIndex = 2;

                await foreach (var row in rows.WithCancellation(ct))
                {
                    var plans = new CellPlan[columns.Count];
                    int rowSpan = 1;

                    for (int i = 0; i < columns.Count; i++)
                    {
                        var value = accessors[i](row!);

                        if (value is string s && s.Length > XmlConstants.MaxCellTextLength)
                        {
                            var chunkCount = (s.Length + XmlConstants.MaxCellTextLength - 1) / XmlConstants.MaxCellTextLength;
                            plans[i] = new CellPlan(value, s, IsLongText: true, ChunkCount: chunkCount);
                            if (chunkCount > rowSpan)
                                rowSpan = chunkCount;
                        }
                        else
                        {
                            plans[i] = new CellPlan(value, value as string, IsLongText: false, ChunkCount: 1);
                        }
                    }

                    for (int chunkRow = 0; chunkRow < rowSpan; chunkRow++)
                    {
                        writer.WriteStartElement(new Row { RowIndex = mainRowIndex });

                        for (int i = 0; i < columns.Count; i++)
                        {
                            var plan = plans[i];

                            if (!plan.IsLongText)
                            {
                                if (chunkRow == 0)
                                    WriteCell(writer, plan.Value);
                                else
                                    writer.WriteElement(new Cell());

                                continue;
                            }

                            if (plan.Text is null)
                            {
                                writer.WriteElement(new Cell());
                                continue;
                            }

                            int start = chunkRow * XmlConstants.MaxCellTextLength;
                            if (start >= plan.Text.Length)
                            {
                                writer.WriteElement(new Cell());
                                continue;
                            }

                            int len = Math.Min(XmlConstants.MaxCellTextLength, plan.Text.Length - start);
                            WriteString(writer, plan.Text.Substring(start, len));
                        }

                        writer.WriteEndElement();
                        mainRowIndex++;
                    }
                }

                writer.WriteEndElement(); // SheetData
                writer.WriteEndElement(); // Worksheet

                workbookPart.Workbook.AppendChild(new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    }));

                workbookPart.Workbook.Save();
            }
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
            writer.WriteStartElement(new Cell { DataType = CellValues.InlineString });
            writer.WriteStartElement(new InlineString());
            writer.WriteElement(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
            writer.WriteEndElement(); // InlineString
            writer.WriteEndElement(); // Cell
        }
    }
}
