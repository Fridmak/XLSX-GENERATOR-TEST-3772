using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Nodes;

namespace Analitics6400.Dal.Services;

public class Document
{
    [Key]
    public Guid Id { get; set; }

    [Description("РРґРµРЅС‚РёС„РёРєР°С‚РѕСЂ СЃС…РµРјС‹ РґРѕРєСѓРјРµРЅС‚Р°")]
    public Guid? DocumentSchemaId { get; set; }

    [Description("Р”Р°С‚Р° Рё РІСЂРµРјСЏ РїСѓР±Р»РёРєР°С†РёРё РґРѕРєСѓРјРµРЅС‚Р°")]
    public DateTime? Published { get; set; }

    [Description("CС‚Р°С‚СѓСЃ РІ Р°СЂС…РёРІРµ РёР»Рё РЅРµС‚")]
    public bool IsArchived { get; set; }

    [Description("Р’РµСЂСЃРёСЏ РґРѕРєСѓРјРµРЅС‚Р°")]
    public double Version { get; set; }

    [Description("Р”РѕР»Р¶РµРЅ Р»Рё СЌС‚РѕС‚ РґРѕРєСѓРјРµРЅС‚ СѓС‡Р°СЃС‚РІРѕРІР°С‚СЊ РІ РІР°Р»РёРґР°С†РёРё РЅР° СѓРЅРёРєР°Р»СЊРЅРѕСЃС‚СЊ?")]
    public bool IsCanForValidate { get; set; }

    [Description("Р”Р°РЅРЅС‹Рµ РґРѕРєСѓРјРµРЅС‚Р°")]
    public JsonObject? JsonData { get; set; }

    [Description("Р”Р°С‚Р° РёР·РјРµРЅРµРЅРёСЏ")]
    public DateTime? ChangedDateUtc { get; set; }
}
