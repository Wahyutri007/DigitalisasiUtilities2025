using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class pekerjalookupus : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_Employees_NopegPersero",
                table: "Employees");

            migrationBuilder.AlterColumn<string>(
                name: "EmailPertamina",
                table: "Employees",
                type: "nvarchar(200)",
                maxLength: 200,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(150)",
                oldMaxLength: 150,
                oldNullable: true);

            migrationBuilder.CreateTable(
                name: "EmployeeLookups",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Category = table.Column<string>(type: "nvarchar(40)", maxLength: 40, nullable: false),
                    Value = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_EmployeeLookups", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Employees_NopegPersero",
                table: "Employees",
                column: "NopegPersero");

            migrationBuilder.CreateIndex(
                name: "IX_EmployeeLookups_Category_Value",
                table: "EmployeeLookups",
                columns: new[] { "Category", "Value" },
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "EmployeeLookups");

            migrationBuilder.DropIndex(
                name: "IX_Employees_NopegPersero",
                table: "Employees");

            migrationBuilder.AlterColumn<string>(
                name: "EmailPertamina",
                table: "Employees",
                type: "nvarchar(150)",
                maxLength: 150,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(200)",
                oldMaxLength: 200,
                oldNullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_Employees_NopegPersero",
                table: "Employees",
                column: "NopegPersero",
                unique: true);
        }
    }
}
