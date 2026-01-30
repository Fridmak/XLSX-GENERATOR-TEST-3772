namespace Analitics6400.Logic.Services.XmlWriters.Models;

public sealed class ExcelColumn
{
    public string Name { get; }
    public Type Type { get; }
    public string Header { get; }
    public bool Nullable { get; }

    public ExcelColumn(string name, Type type, string? header = null, bool nullable = true)
    {
        Name = name;
        Type = type;
        Header = header ?? name;
        Nullable = nullable;
    }
}
