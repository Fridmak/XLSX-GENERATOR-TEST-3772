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
        int rowIndex = 2; // строка после заголовка

        await foreach (var row in rows.WithCancellation(ct))
        {
            // Если достигнут максимум строк листа
            if (rowIndex > XmlConstants.ExcelMaxRows)
            {
                // Вставляем уведомление о переполнении в последнюю строку предыдущего листа
                for (int c = 0; c < columns.Count; c++)
                {
                    ws.Cell(XmlConstants.ExcelMaxRows, c + 1)
                        .SetValue(XmlConstants.JsonOverflowNotice);
                }

                // Создаем новый лист
                sheetId++;
                ws = workbook.AddWorksheet(GetSheetName(sheetId));
                WriteHeader(ws, columns);
                rowIndex = 2;
            }

            for (int c = 0; c < columns.Count; c++)
            {
                var value = accessors[c](row!);


                if (value is string s && s.Length > XmlConstants.MaxCellTextLength)
                {
                    int allowedLength = XmlConstants.MaxCellTextLength - XmlConstants.JsonOverflowNotice.Length;
                    if (allowedLength < 0) allowedLength = XmlConstants.MaxCellTextLength;

                    s = s.Substring(0, allowedLength) + XmlConstants.JsonOverflowNotice;
                    ws.Cell(rowIndex, c + 1).SetValue(s);
                }
                else
                {
                    SetCellValue(ws.Cell(rowIndex, c + 1), value);
                }
            }

            rowIndex++;

            if (rowIndex % 1000 == 0)
                await Task.Yield();
        }

        // Если последний лист не полный, ничего не делаем, т.к. переполнения нет
        workbook.SaveAs(output);
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

    private static void SetCellValue(IXLCell cell, object? value)
    {
        if (value is null)
        {
            cell.SetValue(string.Empty);
            return;
        }

        switch (value)
        {
            case string s:
                cell.SetValue(s);
                break;
            case Guid g:
                cell.SetValue(g.ToString());
                break;
            case bool b:
                cell.SetValue(b);
                break;
            case DateTime dt:
                cell.SetValue(dt);
                break;
            case DateTimeOffset dto:
                cell.SetValue(dto.DateTime);
                break;
            case sbyte or byte or short or ushort or int or uint or long or ulong:
                cell.SetValue(Convert.ToInt64(value));
                break;
            case float or double or decimal:
                cell.SetValue(Convert.ToDouble(value));
                break;
            default:
                cell.SetValue(value.ToString());
                break;
        }
    }
}
