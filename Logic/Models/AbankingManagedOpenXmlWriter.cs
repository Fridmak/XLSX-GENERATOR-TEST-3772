using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Analitics6400.Logic.Models;

public sealed class AbankingManagedOpenXmlWriter : IDisposable
{
    private readonly WorksheetPart _worksheetPart;
    private readonly long _maxMemoryBytes;
    private OpenXmlWriter _writer;
    private long _estimatedMemoryUsage;
    private bool _isDisposed;
    private uint _rowsWritten;

    public AbankingManagedOpenXmlWriter(WorksheetPart worksheetPart, long maxMemoryBytes = 2L * 1024 * 1024 * 1024)
    {
        _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
        _maxMemoryBytes = maxMemoryBytes;

        CreateNewWriter();
    }

    private void CreateNewWriter()
    {
        // Закрываем предыдущий writer
        if (_writer != null)
        {
            _writer.WriteEndElement(); // SheetData
            _writer.WriteEndElement(); // Worksheet
            _writer.Close();
            _worksheetPart.Worksheet.Save();
        }

        // Создаем новый writer
        _writer = OpenXmlWriter.Create(_worksheetPart);

        // Начинаем структуру листа
        _writer.WriteStartElement(new Worksheet());
        _writer.WriteStartElement(new SheetData());

        _rowsWritten = 0;
        _estimatedMemoryUsage = 0;
    }

    public OpenXmlWriter Value()
    {
        CheckDisposed();
        return _writer;
    }

    public void WriteStartElement(OpenXmlElement element)
    {
        CheckDisposed();
        CheckAndRecreateIfNeeded(element);

        _writer.WriteStartElement(element);

        if (element is Row)
        {
            _rowsWritten++;
            _estimatedMemoryUsage += EstimateElementSize(element);
        }
    }

    public void WriteEndElement()
    {
        CheckDisposed();
        _writer.WriteEndElement();
    }

    public void WriteElement(OpenXmlElement element)
    {
        CheckDisposed();
        CheckAndRecreateIfNeeded(element);

        _writer.WriteElement(element);
        _estimatedMemoryUsage += EstimateElementSize(element);
    }

    public void FlushIfNeeded()
    {
        CheckDisposed();

        // Простая эвристика: если приближаемся к лимиту, сбрасываем
        if (_estimatedMemoryUsage > _maxMemoryBytes * 0.8) // 80% от лимита
        {
            FlushAndRecreate();
        }
    }

    public bool FlushAndRecreate()
    {
        CheckDisposed();

        if (_rowsWritten == 0)
            return false;

        CreateNewWriter();
        return true;
    }

    private void CheckAndRecreateIfNeeded(OpenXmlElement element)
    {
        // Если следующий элемент может превысить лимит, пересоздаем writer
        var estimatedSize = EstimateElementSize(element);
        if (_estimatedMemoryUsage + estimatedSize > _maxMemoryBytes)
        {
            CreateNewWriter();
        }
    }

    private long EstimateElementSize(OpenXmlElement element)
    {
        // Упрощенная оценка памяти
        if (element is Cell cell)
        {
            // Ячейка: базовый размер + текст
            var textSize = 0L;
            if (cell.InnerText != null)
            {
                // UTF-8 может быть 1-4 байта на символ, берем 2 как среднее
                textSize = cell.InnerText.Length * 2;
            }
            return 200 + textSize; // 200 байт на структуру XML
        }
        if (element is Row)
        {
            // Строка: базовый размер + ячейки (оценим позже)
            return 100;
        }
        // Консервативная оценка для других элементов
        return 256;
    }

    private void CheckDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(AbankingManagedOpenXmlWriter));
    }

    public void Dispose()
    {
        if (_isDisposed)
            return;

        try
        {
            if (_writer != null)
            {
                _writer.WriteEndElement(); // SheetData
                _writer.WriteEndElement(); // Worksheet
                _writer.Close();
                _worksheetPart.Worksheet.Save();
            }
        }
        finally
        {
            _writer = null!;
            _isDisposed = true;
        }
    }

    // Свойства для мониторинга
    public long EstimatedMemoryUsage => _estimatedMemoryUsage;
    public uint RowsWritten => _rowsWritten;
    public bool IsDisposed => _isDisposed;
}