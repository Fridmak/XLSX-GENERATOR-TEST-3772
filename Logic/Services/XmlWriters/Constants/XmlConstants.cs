namespace Analitics6400.Logic.Services.XmlWriters.Constants;

public static class XmlConstants
{
    public const string OpenXmlName = "OpenXml";

    public const string ClosedXmlName = "ClosedXml";

    public const string JsonOverflowNotice = " ... Используйте генерацию CSV чтобы увидеть полностью";

    public const int ExcelMaxRows = 1_048_576 - 100_000;

    public const int MaxCellTextLength = 32_767 - 100;

    public const long MaxSheetBytes = 999 * 1024 * 1024 ; // мягкое ограничение до 2ГБ (предел для формирования zip из xml)

    public const int TextChunkSize = 67 * 1024; // очень мягкий предел для чанка текста
}