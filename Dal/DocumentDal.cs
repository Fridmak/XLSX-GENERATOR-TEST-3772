using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Nodes;

namespace Analitics6400.Dal;

public class Document
{
    [Key]
    public Guid Id { get; set; }

    [Description("Идентификатор схемы документа")]
    public Guid? DocumentSchemaId { get; set; }

    [Description("Дата и время публикации документа")]
    public DateTime? Published { get; set; }

    [Description("Cтатус в архиве или нет")]
    public bool IsArchived { get; set; }

    [Description("Версия документа")]
    public double Version { get; set; }

    [Description("Должен ли этот документ участвовать в валидации на уникальность?")]
    public bool IsCanForValidate { get; set; }

    [Description("Данные документа")]
    public JsonObject? JsonData { get; set; }

    [Description("Дата изменения")]
    public DateTime? ChangedDateUtc { get; set; }
}
