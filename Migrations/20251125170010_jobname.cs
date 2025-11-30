using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class jobname : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "WorkJobNames",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NamaPekerjaan = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    CreatedAt = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_WorkJobNames", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_WorkTasks_NamaPIC_Status",
                table: "WorkTasks",
                columns: new[] { "NamaPIC", "Status" });

            migrationBuilder.CreateIndex(
                name: "IX_WorkTasks_Prioritas",
                table: "WorkTasks",
                column: "Prioritas");

            migrationBuilder.CreateIndex(
                name: "IX_WorkJobNames_NamaPekerjaan",
                table: "WorkJobNames",
                column: "NamaPekerjaan",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "WorkJobNames");

            migrationBuilder.DropIndex(
                name: "IX_WorkTasks_NamaPIC_Status",
                table: "WorkTasks");

            migrationBuilder.DropIndex(
                name: "IX_WorkTasks_Prioritas",
                table: "WorkTasks");
        }
    }
}
