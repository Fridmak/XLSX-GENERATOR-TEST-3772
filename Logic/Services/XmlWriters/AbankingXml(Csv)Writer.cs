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

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        CancellationToken ct = default)
    {
        if (columns.Count == 0)
            throw new ArgumentException("Columns cannot be empty", nameof(columns));

        // CSV почти всегда ожидают UTF8 без BOM
        using var writer = new StreamWriter(
            output,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            bufferSize: 1024 * 64,
            leaveOpen: true);

        // Header
        await WriteRowAsync(
            writer,
            columns.Select(c => c.Header),
            ct);

        var accessors = GetAccessors<T>(columns);

        await foreach (var row in rows.WithCancellation(ct))
        {
            var values = new string?[columns.Count];

            for (int i = 0; i < columns.Count; i++)
            {
                var value = accessors[i](row!);
                values[i] = FormatValue(value);
            }

            await WriteRowAsync(writer, values, ct);
        }

        await writer.FlushAsync();
    }

    // ---------------- helpers ----------------

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

    private static async Task WriteRowAsync(
        StreamWriter writer,
        IEnumerable<string?> values,
        CancellationToken ct)
    {
        bool first = true;

        foreach (var value in values)
        {
            if (!first)
                await writer.WriteAsync(',');

            await writer.WriteAsync(EscapeCsv(value));
            first = false;
        }

        await writer.WriteLineAsync();
        ct.ThrowIfCancellationRequested();
    }

    private static string EscapeCsv(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        bool mustQuote =
            value.Contains(',') ||
            value.Contains('"') ||
            value.Contains('\n') ||
            value.Contains('\r');

        if (!mustQuote)
            return value;

        // экранируем кавычки
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    private static string? FormatValue(object? value)
    {
        if (value is null)
            return null;

        return value switch
        {
            string s => s,
            Guid g => g.ToString(),
            bool b => b ? "true" : "false",

            DateTime dt => dt.ToString("O", CultureInfo.InvariantCulture),
            DateTimeOffset dto => dto.ToString("O", CultureInfo.InvariantCulture),

            sbyte or byte or short or ushort or int or uint or long or ulong
                or float or double or decimal
                => Convert.ToString(value, CultureInfo.InvariantCulture),

            _ => value.ToString()
        };
    }
}
