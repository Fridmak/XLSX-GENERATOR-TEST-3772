using Analitics6400.Dal;

namespace Analitics6400.Logic.Seed
{
    public class DocumentSeeder
    {
        private readonly DocumentDbContext _context;
        private readonly ILogger<DocumentSeeder> _logger;

        public DocumentSeeder(DocumentDbContext context, ILogger<DocumentSeeder> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task SeedAsync(int totalDocuments = 100_000, int batchSize = 2000)
        {
            var random = new Random();
            _context.ChangeTracker.AutoDetectChangesEnabled = false;

            for (int batch = 0; batch < totalDocuments / batchSize; batch++)
            {
                using var transaction = await _context.Database.BeginTransactionAsync();

                var documents = new Document[batchSize];
                for (int i = 0; i < batchSize; i++)
                {
                    int jsonSizeKb = random.Next(100, 1024);
                    string jsonPayload = GenerateJsonPayload(jsonSizeKb * 1024);

                    documents[i] = new Document
                    {
                        Id = Guid.NewGuid(),
                        DocumentSchemaId = Guid.NewGuid(),
                        Published = DateTime.UtcNow,
                        IsArchived = false,
                        Version = 2.0,
                        IsCanForValidate = false,
                        JsonData = jsonPayload,
                        ChangedDateUtc = DateTime.UtcNow
                    };
                }

                _context.Documents.AddRange(documents);
                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                _logger.LogInformation($"Batch {batch + 1}/{totalDocuments / batchSize} inserted");
            }
        }

        private string GenerateJsonPayload(int targetSizeBytes)
        {
            var obj = new
            {
                access_token = (string?)null,
                currentLanguage = "ru",
                data = new string('x', Math.Max(0, targetSizeBytes - 50))
            };

            return System.Text.Json.JsonSerializer.Serialize(obj);
        }
    }
}
