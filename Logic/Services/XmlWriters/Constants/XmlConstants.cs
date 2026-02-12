namespace Analitics6400.Logic.Services.XmlWriters.Constants;

public static class XmlConstants
{
    public const string OpenXmlName = "OpenXml";

    public const string ClosedXmlName = "ClosedXml";

    public const string JsonOverflowNotice = " ... Используйте генерацию CSV чтобы увидеть полностью";

    public const int ExcelMaxRows = 1_048_576 - 1;

    public const int MaxCellTextLength = 32_767 - 1;

    /// <summary>
    /// Возьмем плохую ситуацию: 100% файлов по 1024кб и вся она в одном поле
    /// То есть 1 млн символов в одном файле, то есть 32 строки на одну запись
    /// Получается что 32767 записей - максимум на одном листе
    /// Возьмем 32 000 как ограничение
    /// </summary>
    public const int BatchSize = 1000;
}