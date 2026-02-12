using Analitics6400.Dal.Services.Interfaces;
using Analitics6400.Logic.Models;
using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using System.Text.Json.Nodes;

namespace Analitics6400.Dal.Services;

public sealed class MockDocumentDalProvider : IDocumentProvider, IDisposable
{
    private readonly int _documentCount;
    private readonly Random _random;

    // Use a static cache to share memory across all instances of the provider
    // This ensures that even if you create multiple providers, they share the heavy string payloads.
    private static readonly ConcurrentDictionary<int, string> _paddingCache = new();

    private readonly List<string> _documentTypes = new()
    {
        "Invoice", "Contract", "Report", "Application", "Order",
        "Payment", "Receipt", "Statement", "Certificate", "License"
    };

    private readonly (int Min, int Max, double Probability)[] _sizeDistribution = new[]
    {
        (1_000, 40_000, 0.10),
        (40_001, 200_000, 0.10),
        (200_001, 1_000_000, 0.80)
    };

    public MockDocumentDalProvider(int documentCount = 1000, int randomSeed = 123456789)
    {
        _documentCount = documentCount;
        _random = new(Seed: randomSeed);
    }

    public async IAsyncEnumerable<DocumentDtoModel> GetDocumentsAsync(
        int? limit = null,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        int returned = 0;
        var documentCount = limit ?? _documentCount;

        for (int i = 0; i < documentCount; i++)
        {
            ct.ThrowIfCancellationRequested();

            var targetSize = GetTargetSize();
            var jsonObject = GenerateJsonOfExactSize(targetSize, i);

            yield return new DocumentDtoModel
            {
                Id = Guid.NewGuid(),
                DocumentSchemaId = Guid.NewGuid(),
                Published = DateTime.UtcNow.AddDays(-_random.Next(0, 365)),
                Version = Math.Round(1.0 + _random.NextDouble() * 9, 1),
                IsArchived = _random.NextDouble() < 0.1,
                JsonData = jsonObject,
                IsCanForValidate = _random.NextDouble() < 0.8,
                ChangedDateUtc = DateTime.UtcNow.AddHours(-_random.Next(0, 720))
            };

            returned++;
            if (limit.HasValue && returned >= limit.Value)
                yield break;

            // Yield occasionally to keep UI/Thread responsive, but not too often
            if (i % 100 == 0)
                await Task.Yield();
        }
    }

    private int GetTargetSize()
    {
        var r = _random.NextDouble();
        double cumulative = 0;

        foreach (var (min, max, probability) in _sizeDistribution)
        {
            cumulative += probability;
            if (r <= cumulative)
            {
                // QUANTIZATION FIX:
                // Round the size to the nearest 1024 bytes (1KB).
                // This drastically increases the hit rate of our _paddingCache.
                // Instead of 1,000,000 unique string sizes, we have ~1000 unique sizes.
                var size = _random.Next(min, max + 1);
                return (int)Math.Ceiling(size / 1024.0) * 1024;
            }
        }

        return 1_000_000; // Fallback max size
    }

    private JsonObject GenerateJsonOfExactSize(int targetSizeBytes, int index)
    {
        var docType = _documentTypes[_random.Next(_documentTypes.Count)];
        var obj = CreateBaseJsonObject(docType, index);

        var paddingKey = "payload"; // Keeping key constant saves a tiny bit more memory, but optional.

        // Estimate base size (approx 200 bytes)
        var estimatedCurrentSize = 200;
        var bytesNeeded = targetSizeBytes - estimatedCurrentSize;

        if (bytesNeeded > 50)
        {
            // MEMORY FIX:
            // Instead of 'new string(...)', we get a reference to an existing string from cache.
            // 1000 documents will now point to the SAME string instance in memory.
            // This reduces RAM usage for the payload by factor of ~N (where N is reuse count).

            // Adjust bytesNeeded to match our quantization step to ensure cache hit
            var quantizedSize = (int)Math.Ceiling(bytesNeeded / 1024.0) * 1024;

            var padding = _paddingCache.GetOrAdd(quantizedSize, size => new string('x', size));

            obj[paddingKey] = padding;
        }
        else
        {
            obj["data"] = "x";
        }

        return obj;
    }

    private JsonObject CreateBaseJsonObject(string docType, int index)
    {
        return new JsonObject
        {
            ["documentId"] = Guid.NewGuid().ToString(),
            ["type"] = docType,
            ["number"] = $"DOC-{DateTime.UtcNow:yyyyMMdd}-{index:00000}",
            ["createdAt"] = DateTime.UtcNow.ToString("O"),
            ["status"] = _random.NextDouble() < 0.8 ? "Active" : "Archived"
        };
    }

    public void Dispose()
    {
        // Cache is static, so we don't clear it here to benefit other instances.
        // If you strictly want to clear it, remove 'static' from the dictionary definition.
    }
}