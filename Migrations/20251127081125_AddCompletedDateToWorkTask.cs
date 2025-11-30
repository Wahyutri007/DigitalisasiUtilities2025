using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class AddCompletedDateToWorkTask : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "CompletedDate",
                table: "WorkTasks",
                type: "datetime2",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CompletedDate",
                table: "WorkTasks");
        }
    }
}
