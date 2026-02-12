using Analitics6400.Dal;
using Microsoft.EntityFrameworkCore;
using System.Text;
using System.Text.Json.Nodes;

namespace Analitics6400.Logic.Seed
{
    public class DocumentSeeder
    {
        private readonly DocumentDbContext _context;
        private readonly ILogger<DocumentSeeder> _logger;
        private readonly Random _random = new();

        public DocumentSeeder(DocumentDbContext context, ILogger<DocumentSeeder> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task SeedAsync(int totalDocuments = 100_000, int batchSize = 1000)
        {
            _logger.LogInformation($"Starting seeding of {totalDocuments} documents...");

            _context.ChangeTracker.AutoDetectChangesEnabled = false;
            _context.ChangeTracker.QueryTrackingBehavior = QueryTrackingBehavior.NoTracking;

            var batches = totalDocuments / batchSize;
            var jsonTemplate = GenerateJsonTemplate();

            for (int batch = 0; batch < batches; batch++)
            {
                await using var transaction = await _context.Database.BeginTransactionAsync();

                try
                {
                    await InsertBatchAsync(batchSize, batch, jsonTemplate);
                    await transaction.CommitAsync();

                    _logger.LogInformation($"Batch {batch + 1}/{batches} inserted ({batchSize * (batch + 1)} total)");

                    if ((batch + 1) % 10 == 0)
                    {
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    _logger.LogError(ex, $"Error in batch {batch + 1}");
                    throw;
                }
            }

            _context.ChangeTracker.AutoDetectChangesEnabled = true;
            _logger.LogInformation("Seeding completed!");
        }

        private async Task InsertBatchAsync(int batchSize, int batchNumber, JsonObject jsonTemplate)
        {
            var sqlBuilder = new StringBuilder();
            sqlBuilder.AppendLine("INSERT INTO \"documents\" (\"Id\", \"DocumentSchemaId\", \"Published\", \"IsArchived\", \"Version\", \"IsCanForValidate\", \"JsonData\", \"ChangedDateUtc\") VALUES ");

            var parameters = new List<object>();
            var paramIndex = 0;

            for (int i = 0; i < batchSize; i++)
            {
                if (i > 0)
                    sqlBuilder.Append(", ");

                var jsonSizeKb = _random.Next(512, 1024);
                var jsonData = GenerateJsonPayloadOptimized(jsonTemplate, jsonSizeKb * 1024, paramIndex);

                sqlBuilder.Append(
                  $"(@p{paramIndex++}, @p{paramIndex++}, @p{paramIndex++}, @p{paramIndex++}, " +
                  $"@p{paramIndex++}, @p{paramIndex++}, @p{paramIndex++}::jsonb, @p{paramIndex++})"
                );

                parameters.Add(Guid.NewGuid());
                parameters.Add(Guid.NewGuid());
                parameters.Add(DateTime.UtcNow);
                parameters.Add(false);
                parameters.Add(2.0);
                parameters.Add(false);
                parameters.Add(jsonData.ToJsonString());
                parameters.Add(DateTime.UtcNow);
            }

            var sql = sqlBuilder.ToString();
            await _context.Database.ExecuteSqlRawAsync(sql, parameters.ToArray());
        }

        private JsonObject GenerateJsonTemplate()
        {
            return new JsonObject
            {
                ["access_token"] = null,
                ["currentLanguage"] = "ru",
                ["data"] = string.Empty
            };
        }

        private JsonObject GenerateJsonPayloadOptimized(JsonObject template, int targetSizeBytes, int seed)
        {
            var random = new Random(seed + Environment.TickCount);
            var baseSize = template.ToJsonString().Length;
            var neededChars = targetSizeBytes - baseSize - 50;

            if (neededChars <= 0)
                return template;

            var dataString = GenerateLargeStringOptimized(neededChars);

            var jsonObj = new JsonObject
            {
                ["access_token"] = null,
                ["currentLanguage"] = "ru",
                ["data"] = dataString
            };

            return jsonObj;
        }

        private string GenerateLargeStringOptimized(int length)
        {
            var sb = new StringBuilder(length);
            const string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 _-.,:;";

            for (int i = 0; i < length; i++)
            {
                sb.Append(chars[_random.Next(chars.Length)]);
            }

            return sb.ToString();
        }
    }
}
