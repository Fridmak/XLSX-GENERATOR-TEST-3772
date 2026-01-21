using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Analitics6400.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "documents",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uuid", nullable: false),
                    DocumentSchemaId = table.Column<Guid>(type: "uuid", nullable: true),
                    Published = table.Column<DateTime>(type: "timestamp with time zone", nullable: true),
                    IsArchived = table.Column<bool>(type: "boolean", nullable: false),
                    Version = table.Column<double>(type: "double precision", nullable: false),
                    IsCanForValidate = table.Column<bool>(type: "boolean", nullable: false),
                    JsonData = table.Column<string>(type: "jsonb", nullable: false),
                    ChangedDateUtc = table.Column<DateTime>(type: "timestamp with time zone", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_documents", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "documents");
        }
    }
}
