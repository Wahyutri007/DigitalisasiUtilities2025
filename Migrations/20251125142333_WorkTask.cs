using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class WorkTask : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "WorkTasks",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NamaPekerjaan = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    StartDate = table.Column<DateTime>(type: "datetime2", nullable: true),
                    DueDate = table.Column<DateTime>(type: "datetime2", nullable: true),
                    NamaRequest = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    ProgressPercent = table.Column<int>(type: "int", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false),
                    KeteranganTindakLanjut = table.Column<string>(type: "nvarchar(1000)", maxLength: 1000, nullable: true),
                    Prioritas = table.Column<int>(type: "int", nullable: false),
                    NamaPIC = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_WorkTasks", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "WorkTasks");
        }
    }
}
