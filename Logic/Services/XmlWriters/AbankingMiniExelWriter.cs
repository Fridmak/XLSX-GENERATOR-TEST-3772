using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using System.Buffers;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingMiniExcelWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type, string), Func<object, object?>[]> _accessors = new();

    private const int SharedStringCacheBytes = 5 * 1024 * 1024; // 5MB
    private const int OptimalBufferSize = 512 * 1024;           // 512KB

    public string Extension => "xlsx";

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        string? baseSheetName = null,
        CancellationToken ct = default,
        bool useTemplate = false,
        Stream? templateStream = null)
    {
        if (output == null)
        {
            throw new ArgumentNullException(nameof(output));
        }
        if (!output.CanSeek)
        {
            throw new InvalidOperationException("Output stream must support seeking.");
        }
        if (columns == null || columns.Count == 0)
        {
            throw new ArgumentException("Columns are required.", nameof(columns));
        }

        if (useTemplate && templateStream != null)
        {
            await GenerateFromTemplateAsync(rows, columns, output, templateStream, baseSheetName, ct).ConfigureAwait(false);
            return;
        }

        var accessors = GetAccessors<T>(columns);
        var headers = columns.Select(c => c.Header ?? c.Name).ToArray();

        var config = new OpenXmlConfiguration
        {
            FastMode = false,
            TableStyles = TableStyles.None,
            AutoFilter = false,
            EnableSharedStringCache = true,
            SharedStringCacheSize = SharedStringCacheBytes,
            IgnoreEmptyRows = true,
            EnableWriteNullValueCell = true,
            WriteEmptyStringAsNull = false,
            EnableAutoWidth = false,
            BufferSize = OptimalBufferSize
        };

        var physicalRows = ConvertToStreamingRowsAsync(rows, headers, accessors, ct);

        await MiniExcel.SaveAsAsync(
            output,
            physicalRows,
            printHeader: true,
            sheetName: baseSheetName!,
            excelType: ExcelType.XLSX,
            configuration: config,
            cancellationToken: ct).ConfigureAwait(false);
    }

    /// <summary>
    /// ЖРЕТ ПАМЯТЬ! Использовать только для небольших наборов данных, которые гарантированно поместятся в память.
    /// Применяет данные к шаблону Excel (заполняет {{placeholders}} и коллекции {{collection.field}}).
    /// Шаблон должен содержать заголовки и формулы; данные дописываются в соответствии с шаблоном.
    /// </summary>
    private async Task GenerateFromTemplateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        Stream templateStream,
        string? baseSheetName,
        CancellationToken ct)
    {
        // MiniExcel.SaveAsByTemplate требует материализованную коллекцию
        var rowsList = new List<IDictionary<string, object?>>();
        var accessors = GetAccessors<T>(columns);
        var headers = columns.Select(c => c.Header ?? c.Name).ToArray();

        await foreach (var item in rows.WithCancellation(ct).ConfigureAwait(false))
        {
            var columnCount = headers.Length;
            var values = ArrayPool<object?>.Shared.Rent(columnCount);

            try
            {
                for (int i = 0; i < columnCount; i++)
                {
                    var rawValue = accessors[i](item!);
                    values[i] = NormalizeValue(rawValue);
                }

                foreach (var row in BuildChunkedRows(headers, values, ct))
                {
                    rowsList.Add(row);
                }
            }
            finally
            {
                ArrayPool<object?>.Shared.Return(values);
            }
        }

        var config = new OpenXmlConfiguration
        {
            IgnoreTemplateParameterMissing = true, // Игнорируем отсутствующие параметры в шаблоне
            EnableSharedStringCache = true,
            SharedStringCacheSize = SharedStringCacheBytes
        };

        templateStream.Seek(0, SeekOrigin.Begin);
        byte[] templateBytes = new byte[templateStream.Length];
        await templateStream.ReadAsync(templateBytes, 0, (int)templateStream.Length, ct).ConfigureAwait(false);

        // MiniExcel.SaveAsByTemplate заполнит {{placeholders}} и {{collection.field}} в шаблоне
        await MiniExcel.SaveAsByTemplateAsync(
            output,
            templateBytes,
            new { Data = rowsList }, // Обёртываем в объект с именем "Data" для использования {{Data.ColumnName}}
            configuration: config,
            cancellationToken: ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Конвертация записей в физические строки Excel с нарезкой длинных текстов.
    /// </summary>
    private static async IAsyncEnumerable<IDictionary<string, object?>> ConvertToStreamingRowsAsync<T>(
    IAsyncEnumerable<T> source,
    string[] headers,
    Func<object, object?>[] accessors,
    [EnumeratorCancellation] CancellationToken ct)
    {
        await foreach (var item in source.WithCancellation(ct).ConfigureAwait(false))
        {
            var columnCount = headers.Length;
            var values = ArrayPool<object?>.Shared.Rent(columnCount);

            try
            {
                for (int i = 0; i < columnCount; i++)
                {
                    var rawValue = accessors[i](item!);
                    values[i] = NormalizeValue(rawValue);
                }

                foreach (var row in BuildChunkedRows(headers, values, ct))
                {
                    yield return row;
                }
            }
            finally
            {
                ArrayPool<object?>.Shared.Return(values);
            }
        }
    }


    /// <summary>
    /// Строит физические строки Excel с нарезкой длинных текстов.
    /// </summary>
    private static IEnumerable<IDictionary<string, object?>> BuildChunkedRows(
        string[] headers,
        object?[] values,
        CancellationToken ct)
    {
        var columnCount = headers.Length;

        int maxChunks = 1;

        for (int i = 0; i < columnCount; i++)
        {
            if (values[i] is string str && str.Length > 0)
            {
                var chunks = (str.Length + XmlConstants.MaxCellTextLength - 1) / XmlConstants.MaxCellTextLength;
                if (chunks > maxChunks)
                {
                    maxChunks = chunks;
                }
            }
        }

        for (int chunk = 0; chunk < maxChunks; chunk++)
        {
            ct.ThrowIfCancellationRequested();

            var row = new Dictionary<string, object?>(capacity: columnCount);
            var isFirstChunk = chunk == 0;

            for (int i = 0; i < columnCount; i++)
            {
                var header = headers[i];
                var value = values[i];

                if (value is string s)
                {
                    if (isFirstChunk)
                    {
                        var take = s.Length > XmlConstants.MaxCellTextLength ? XmlConstants.MaxCellTextLength : s.Length;
                        var head = take == s.Length ? s : s.Substring(0, take);

                        if (s.Length > XmlConstants.MaxCellTextLength)
                        {
                            head = head + XmlConstants.JsonOverflowNotice;
                        }

                        row[header] = head;
                    }
                    else
                    {
                        var offset = chunk * XmlConstants.MaxCellTextLength;
                        if (offset < s.Length)
                        {
                            var len = System.Math.Min(XmlConstants.MaxCellTextLength, s.Length - offset);
                            row[header] = s.Substring(offset, len);
                        }
                        else
                        {
                            row[header] = string.Empty;
                        }
                    }
                }
                else
                {
                    row[header] = isFirstChunk ? value : string.Empty;
                }
            }

            yield return row;
        }
    }


    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static object? NormalizeValue(object? value)
    {
        return value switch
        {
            null => null,
            string s => s,
            bool b => b,
            byte _ => value,
            sbyte _ => value,
            short _ => value,
            ushort _ => value,
            int _ => value,
            uint _ => value,
            long _ => value,
            ulong _ => value,
            float _ => value,
            double _ => value,
            decimal _ => value,
            System.DateTime _ => value,
            System.DateTimeOffset _ => value,
            System.Guid g => g.ToString(),
            _ => JsonSerializer.Serialize(value, JsonSerializerOptions.Default)
        };
    }

    private static Func<object, object?>[] GetAccessors<T>(IReadOnlyList<ExcelColumn> columns)
    {
        var key = (typeof(T), string.Join('|', columns.Select(c => c.Name)));
        return _accessors.GetOrAdd(key, _ =>
        {
            var props = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

            return columns.Select(col =>
            {
                if (props.TryGetValue(col.Name, out var prop))
                {
                    var param = System.Linq.Expressions.Expression.Parameter(typeof(object), "obj");
                    var cast = System.Linq.Expressions.Expression.Convert(param, typeof(T));
                    var property = System.Linq.Expressions.Expression.Property(cast, prop);
                    var convert = System.Linq.Expressions.Expression.Convert(property, typeof(object));
                    var lambda = System.Linq.Expressions.Expression.Lambda<Func<object, object?>>(convert, param);
                    return lambda.Compile();
                }
                return static _ => null;
            }).ToArray();
        });
    }
}
