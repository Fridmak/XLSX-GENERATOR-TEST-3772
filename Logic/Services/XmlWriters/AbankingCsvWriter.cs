using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingCsvWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type Type, string ColumnsKey), Func<object, object?>[]> _propertyAccessors = new();
    private const int _flushInterval = 10000;
    private readonly ILogger<AbankingCsvWriter> _logger;
    public string Extension => ".csv";

    public AbankingCsvWriter(ILogger<AbankingCsvWriter> logger)
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

        _logger.LogInformation("Начало генерации при помощи AbankingCsvWriter");

        using var writer = new StreamWriter(
            output,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            bufferSize: 1024 * 1024,
            leaveOpen: true);

        await WriteHeaderAsync(writer, columns, ct);

        var accessors = GetAccessors<T>(columns);
        var rowCount = 0;

        await foreach (var row in rows.WithCancellation(ct))
        {
            await WriteDataRowAsync(writer, row, columns, accessors, ct);

            rowCount++;
            if (rowCount % _flushInterval == 0)
            {
                await writer.FlushAsync(ct);
            }
        }

        await writer.FlushAsync(ct);
        _logger.LogInformation($"Сгенерировано строк: {rowCount}");
    }

    #region Helpers

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

    private static async Task WriteHeaderAsync(
        StreamWriter writer,
        IReadOnlyList<ExcelColumn> columns,
        CancellationToken ct)
    {
        var first = true;
        foreach (var column in columns)
        {
            if (!first)
            {
                await writer.WriteAsync(',');
            }

            await writer.WriteAsync(EscapeCsv(column.Header ?? string.Empty));
            first = false;
        }
        await writer.WriteLineAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static async Task WriteDataRowAsync<T>(
        StreamWriter writer,
        T row,
        IReadOnlyList<ExcelColumn> columns,
        Func<object, object?>[] accessors,
        CancellationToken ct)
    {
        var first = true;
        for (int i = 0; i < columns.Count; i++)
        {
            if (!first)
            {
                await writer.WriteAsync(',');
            }

            var value = accessors[i](row!);
            var formattedValue = FormatAndEscapeValue(value);
            await writer.WriteAsync(formattedValue);
            first = false;
        }
        await writer.WriteLineAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static string FormatAndEscapeValue(object? value)
    {
        if (value is null || value is DBNull)
        {
            return string.Empty;
        }

        string stringValue = value switch
        {
            string s => s,
            Guid g => g.ToString(),
            bool b => b ? "true" : "false",

            DateTime dt => dt.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture),
            DateTimeOffset dto => dto.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz", CultureInfo.InvariantCulture),

            sbyte or byte or short or ushort or int or uint
            or long or ulong or float or double or decimal
                => Convert.ToString(value, CultureInfo.InvariantCulture)!,

            _ => value.ToString() ?? string.Empty
        };

        return EscapeCsv(stringValue);
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var mustQuote = value.AsSpan().IndexOfAny("\",\r\n") >= 0;

        if (!mustQuote)
        {
            return value;
        }

        return "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }
    #endregion
}