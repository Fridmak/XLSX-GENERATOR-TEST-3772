using Analitics6400.Dal.Services;
using Microsoft.EntityFrameworkCore;

namespace Analitics6400.Dal;

public class DocumentDbContext : DbContext
{
    public DbSet<Document> Documents { get; set; }

    public DocumentDbContext(DbContextOptions<DocumentDbContext> options)
        : base(options)
    {
    }

    public DocumentDbContext() { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Document>()
            .ToTable("documents")
            .Property(d => d.JsonData)
            .HasColumnType("jsonb");
    }
}
