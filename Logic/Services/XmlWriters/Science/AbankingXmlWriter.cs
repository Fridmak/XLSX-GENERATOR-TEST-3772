using Analitics6400.Logic.Services.XmlWriters.Constants;
using Analitics6400.Logic.Services.XmlWriters.Interfaces;
using Analitics6400.Logic.Services.XmlWriters.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Buffers;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Analitics6400.Logic.Services.XmlWriters;

public sealed class AbankingXmlWriter : IXmlWriter
{
    private static readonly ConcurrentDictionary<(Type, string), Func<object, object?>[]> _accessors = new();
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public string Extension => "xlsx";

    public async Task GenerateAsync<T>(
        IAsyncEnumerable<T> rows,
        IReadOnlyList<ExcelColumn> columns,
        Stream output,
        CancellationToken ct = default)
    {
        using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var accessors = GetAccessors<T>(columns);
        uint sheetId = 1;
        uint rowIndex = 1;

        // Создаём первый лист
        WorksheetPart worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
        Stream stream = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
        XmlWriter writer = CreateXmlWriter(stream);

        try
        {
            WriteHeader(writer, columns);
            rowIndex = 2;

            await foreach (var row in rows.WithCancellation(ct))
            {
                ct.ThrowIfCancellationRequested();

                if (rowIndex > XmlConstants.ExcelMaxRows)
                {
                    // Закрываем текущий лист
                    CloseSheet(writer);
                    writer.Dispose();
                    stream.Dispose();

                    // Создаём новый лист
                    sheetId++;
                    rowIndex = 1;
                    worksheetPart = CreateWorksheetPart(workbookPart, sheets, sheetId);
                    stream = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
                    writer = CreateXmlWriter(stream);

                    WriteHeader(writer, columns);
                    rowIndex = 2;
                }

                int maxChunks = GetMaxChunks(row, accessors, columns);
                for (int chunk = 0; chunk < maxChunks; chunk++)
                {
                    WriteRowStart(writer, rowIndex);
                    for (int c = 0; c < columns.Count; c++)
                    {
                        WriteCellChunk(writer, accessors[c](row), chunk);
                    }
                    writer.WriteEndElement(); // </row>
                    rowIndex++;
                }

                // КРИТИЧЕСКИ ВАЖНО: сброс на диск после каждой строки
                writer.Flush();
                stream.Flush();

                if (rowIndex % 10000 == 0)
                    await Task.Yield();
            }

            CloseSheet(writer);
        }
        finally
        {
            writer?.Dispose();
            stream?.Dispose();
        }
    }

    private static int GetMaxChunks<T>(
        T row,
        Func<object, object?>[] accessors,
        IReadOnlyList<ExcelColumn> columns)
    {
        int maxChunks = 1;
        for (int c = 0; c < columns.Count; c++)
        {
            var value = accessors[c](row);
            if (value == null) continue;

            var text = value is string s ? s : value.ToString() ?? string.Empty;
            int chunks = (text.Length + XmlConstants.MaxCellTextLength - 1)
                         / XmlConstants.MaxCellTextLength;
            if (chunks > maxChunks) maxChunks = chunks;
        }
        return maxChunks;
    }

    private static void WriteCellChunk(XmlWriter writer, object? value, int chunkIndex)
    {
        if (value == null)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteEndElement();
            return;
        }

        ReadOnlySpan<char> textSpan = value switch
        {
            string s => s.AsSpan(),
            _ => (value.ToString() ?? string.Empty).AsSpan()
        };

        int max = XmlConstants.MaxCellTextLength;
        int offset = chunkIndex * max;
        int len = (offset < textSpan.Length)
            ? Math.Min(max, textSpan.Length - offset)
            : 0;

        if (len <= 0)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteEndElement();
            return;
        }

        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("t", "inlineStr");
        writer.WriteStartElement("is", SpreadsheetNamespace);
        writer.WriteStartElement("t", SpreadsheetNamespace);
        writer.WriteAttributeString("xml", "space", null, "preserve");

        WriteSpanToXml(writer, textSpan.Slice(offset, len));

        writer.WriteEndElement(); // </t>
        writer.WriteEndElement(); // </is>
        writer.WriteEndElement(); // </c>
    }

    private static void WriteSpanToXml(XmlWriter writer, ReadOnlySpan<char> span)
    {
        const int MaxPoolSize = 81920; // 80 KB — безопасный порог для избежания LOH

        if (span.Length <= MaxPoolSize)
        {
            var buffer = ArrayPool<char>.Shared.Rent(span.Length);
            try
            {
                span.CopyTo(buffer);
                writer.WriteChars(buffer, 0, span.Length);
            }
            finally
            {
                ArrayPool<char>.Shared.Return(buffer);
            }
        }
        else
        {
            // Разбиваем большие чанки (>80KB) на части
            int offset = 0;
            while (offset < span.Length)
            {
                int chunkSize = Math.Min(MaxPoolSize, span.Length - offset);
                var buffer = ArrayPool<char>.Shared.Rent(chunkSize);
                try
                {
                    span.Slice(offset, chunkSize).CopyTo(buffer);
                    writer.WriteChars(buffer, 0, chunkSize);
                }
                finally
                {
                    ArrayPool<char>.Shared.Return(buffer);
                }
                offset += chunkSize;
            }
        }
    }

    private static void WriteRowStart(XmlWriter writer, uint rowIndex)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", rowIndex.ToString());
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

    private static XmlWriter CreateXmlWriter(Stream stream)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = Encoding.UTF8,
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            Async = false,
            CloseOutput = false
            // ⚠️ BufferSize НЕ СУЩЕСТВУЕТ — убрано
        };

        var writer = XmlWriter.Create(stream, settings);

        // Пишем начальный XML напрямую
        writer.WriteStartDocument();
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);

        return writer;
    }

    private static void CloseSheet(XmlWriter writer)
    {
        writer.WriteEndElement(); // </sheetData>
        writer.WriteEndElement(); // </worksheet>
        writer.WriteEndDocument();
        writer.Flush();
    }

    private static void WriteHeader(XmlWriter writer, IReadOnlyList<ExcelColumn> columns)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", "1");

        foreach (var col in columns)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", SpreadsheetNamespace);
            writer.WriteStartElement("t", SpreadsheetNamespace);
            writer.WriteAttributeString("xml", "space", null, "preserve");
            writer.WriteString(col.Header ?? "");
            writer.WriteEndElement(); // </t>
            writer.WriteEndElement(); // </is>
            writer.WriteEndElement(); // </c>
        }

        writer.WriteEndElement(); // </row>
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