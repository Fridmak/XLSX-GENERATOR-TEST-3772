using Sylvan.Data.Excel;
using System.Data;
using System.Data.Common;

namespace Analitics6400.Logic.Services.XmlWriters;

//public class SylvanExcelWriter
//{
//    public async Task GenerateReportAsync(
//    DbCommand command,
//    string outputPath,
//    CancellationToken ct = default)
//    {
//        await using var fileStream = new FileStream(
//            outputPath,
//            FileMode.Create,
//            FileAccess.Write,
//            FileShare.None,
//            64 * 1024,
//            FileOptions.SequentialScan);

//        await using var reader = await command.ExecuteReaderAsync(CommandBehavior.SequentialAccess, ct);
//        await using var writer = ExcelDataWriter.Create(fileStream);

//        await writer.WriteAsync(reader, "Report", ct);
//    }
//}