using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class Jenisdinas : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "PenanggungJawab",
                table: "MaterialTxns",
                type: "nvarchar(150)",
                maxLength: 150,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(120)",
                oldMaxLength: 120,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "Keterangan",
                table: "MaterialTxns",
                type: "nvarchar(500)",
                maxLength: 500,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(300)",
                oldMaxLength: 300,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "Satuan",
                table: "Materials",
                type: "nvarchar(30)",
                maxLength: 30,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(30)",
                oldMaxLength: 30);

            migrationBuilder.AlterColumn<string>(
                name: "PosisiBarang",
                table: "Materials",
                type: "nvarchar(150)",
                maxLength: 150,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(120)",
                oldMaxLength: 120,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "NamaBarang",
                table: "Materials",
                type: "nvarchar(200)",
                maxLength: 200,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "nvarchar(150)",
                oldMaxLength: 150);

            migrationBuilder.AlterColumn<string>(
                name: "Nama",
                table: "JenisBarangs",
                type: "nvarchar(150)",
                maxLength: 150,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "nvarchar(100)",
                oldMaxLength: 100);

            migrationBuilder.AlterColumn<string>(
                name: "Keterangan",
                table: "JenisBarangs",
                type: "nvarchar(300)",
                maxLength: 300,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(200)",
                oldMaxLength: 200,
                oldNullable: true);

            migrationBuilder.AddColumn<int>(
                name: "JenisDinasId",
                table: "Dinas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateTable(
                name: "JenisDinas",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Nama = table.Column<string>(type: "nvarchar(120)", maxLength: 120, nullable: false),
                    Keterangan = table.Column<string>(type: "nvarchar(400)", maxLength: 400, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_JenisDinas", x => x.Id);
                });

            // SEED HANYA UNTUK JENISDINAS (JenisBarangs sudah di-seed di migration lain)
            migrationBuilder.InsertData(
                table: "JenisDinas",
                columns: new[] { "Id", "Keterangan", "Nama" },
                values: new object[,]
                {
                    { 1, "Pelatihan/Workshop", "Pelatihan" },
                    { 2, "Kunjungan Kerja", "Kunjungan" }
                });

            migrationBuilder.CreateIndex(
                name: "IX_Dinas_JenisDinasId",
                table: "Dinas",
                column: "JenisDinasId");

            migrationBuilder.AddForeignKey(
                name: "FK_Dinas_JenisDinas_JenisDinasId",
                table: "Dinas",
                column: "JenisDinasId",
                principalTable: "JenisDinas",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Dinas_JenisDinas_JenisDinasId",
                table: "Dinas");

            migrationBuilder.DropTable(
                name: "JenisDinas");

            migrationBuilder.DropIndex(
                name: "IX_Dinas_JenisDinasId",
                table: "Dinas");

            migrationBuilder.DropColumn(
                name: "JenisDinasId",
                table: "Dinas");

            migrationBuilder.AlterColumn<string>(
                name: "PenanggungJawab",
                table: "MaterialTxns",
                type: "nvarchar(120)",
                maxLength: 120,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(150)",
                oldMaxLength: 150,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "Keterangan",
                table: "MaterialTxns",
                type: "nvarchar(300)",
                maxLength: 300,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(500)",
                oldMaxLength: 500,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "Satuan",
                table: "Materials",
                type: "nvarchar(30)",
                maxLength: 30,
                nullable: false,
                defaultValue: "",
                oldClrType: typeof(string),
                oldType: "nvarchar(30)",
                oldMaxLength: 30,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "PosisiBarang",
                table: "Materials",
                type: "nvarchar(120)",
                maxLength: 120,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(150)",
                oldMaxLength: 150,
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "NamaBarang",
                table: "Materials",
                type: "nvarchar(150)",
                maxLength: 150,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "nvarchar(200)",
                oldMaxLength: 200);

            migrationBuilder.AlterColumn<string>(
                name: "Nama",
                table: "JenisBarangs",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: false,
                oldClrType: typeof(string),
                oldType: "nvarchar(150)",
                oldMaxLength: 150);

            migrationBuilder.AlterColumn<string>(
                name: "Keterangan",
                table: "JenisBarangs",
                type: "nvarchar(200)",
                maxLength: 200,
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(300)",
                oldMaxLength: 300,
                oldNullable: true);
        }
    }
}
