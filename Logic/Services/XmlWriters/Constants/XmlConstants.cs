namespace Analitics6400.Logic.Services.XmlWriters.Constants;

public static class XmlConstants
{
    public const string OpenXmlName = "OpenXml";

    public const string ClosedXmlName = "ClosedXml";

    public const string JsonOverflowNotice = " ... Используйте генерацию CSV чтобы увидеть полностью";

    public const int ExcelMaxRows = 1_048_576 - 1;

    //public const int ExcelMaxRows = 340_576 - 1;
    //public const int MaxCellTextLength = 32_767 - 1;

    public const int MaxCellTextLength = 31_767 - 1;
}
