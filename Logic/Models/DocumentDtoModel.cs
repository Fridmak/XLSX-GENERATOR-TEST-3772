using System.Text.Json.Nodes;

namespace Analitics6400.Logic.Models;

public record DocumentDtoModel
{
    public Guid Id { get; init; }
    public Guid? DocumentSchemaId { get; init; }
    public DateTime? Published { get; init; }
    public bool IsArchived { get; init; }
    public double Version { get; init; }
    public JsonObject? JsonData { get; init; }
    public bool IsCanForValidate { get; init; }
    public DateTime? ChangedDateUtc { get; init; }
}
