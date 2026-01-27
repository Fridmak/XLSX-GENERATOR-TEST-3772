using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Models;
using ClosedXML.Excel;
using System.Collections.Concurrent;
using System.Reflection;

namespace Analitics6400.Logic.Services.XmlWriters;

public class AbankingClosedXmlWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        CancellationToken ct = default)
    {
        if (columns.Count == 0)
            throw new ArgumentException("Columns cannot be empty", nameof(columns));

        using var workbook = new XLWorkbook(XLEventTracking.Disabled);
        var ws = workbook.AddWorksheet("Sheet1");

        for (int c = 0; c < columns.Count; c++)
        {
            ws.Cell(1, c + 1).Value = columns[c].Header;
        }

        var accessors = GetAccessors<T>(columns);

        int rowIndex = 2;

        await foreach (var row in rows.WithCancellation(ct))
        {
            var values = new object?[columns.Count];
            int rowSpan = 1;

            for (int c = 0; c < columns.Count; c++)
            {
                var value = accessors[c](row!);
                values[c] = value;

                if (value is string s && s.Length > XmlConstants.MaxCellTextLength)
                {
                    var chunkCount = (s.Length + XmlConstants.MaxCellTextLength - 1) / XmlConstants.MaxCellTextLength;
                    if (chunkCount > rowSpan)
                        rowSpan = chunkCount;
                }
            }

            for (int chunkRow = 0; chunkRow < rowSpan; chunkRow++)
            {
                for (int c = 0; c < columns.Count; c++)
                {
                    var value = values[c];

                    if (value is null)
                    {
                        continue;
                    }

                    if (value is string s)
                    {
                        if (s.Length <= XmlConstants.MaxCellTextLength)
                        {
                            if (chunkRow == 0)
                                ws.Cell(rowIndex, c + 1).Value = s;
                            continue;
                        }

                        int start = chunkRow * XmlConstants.MaxCellTextLength;
                        if (start >= s.Length)
                            continue;

                        int len = Math.Min(XmlConstants.MaxCellTextLength, s.Length - start);
                        ws.Cell(rowIndex, c + 1).Value = s.Substring(start, len);
                        continue;
                    }

                    if (chunkRow == 0)
                    {
                        ws.Cell(rowIndex, c + 1).Value = value;
                    }
                }

                rowIndex++;
            }
        }

        workbook.SaveAs(output);
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
