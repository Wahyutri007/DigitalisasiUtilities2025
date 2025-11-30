using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class dinaslokasi : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DeleteData(
                table: "JenisDinas",
                keyColumn: "Id",
                keyValue: 1);

            migrationBuilder.DeleteData(
                table: "JenisDinas",
                keyColumn: "Id",
                keyValue: 2);

            migrationBuilder.DeleteData(
                table: "MasterSertifikasis",
                keyColumn: "Id",
                keyValue: 3);

            migrationBuilder.DeleteData(
                table: "MasterSertifikasis",
                keyColumn: "Id",
                keyValue: 4);

            migrationBuilder.DeleteData(
                table: "MasterSertifikasis",
                keyColumn: "Id",
                keyValue: 5);

            migrationBuilder.AlterColumn<int>(
                name: "Sifat",
                table: "Dinas",
                type: "int",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AddColumn<string>(
                name: "Lokasi",
                table: "Dinas",
                type: "nvarchar(200)",
                maxLength: 200,
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_JenisDinas_Nama",
                table: "JenisDinas",
                column: "Nama");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_JenisDinas_Nama",
                table: "JenisDinas");

            migrationBuilder.DropColumn(
                name: "Lokasi",
                table: "Dinas");

            migrationBuilder.AlterColumn<int>(
                name: "Sifat",
                table: "Dinas",
                type: "int",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "int",
                oldNullable: true);

            migrationBuilder.InsertData(
                table: "JenisDinas",
                columns: new[] { "Id", "Keterangan", "Nama" },
                values: new object[,]
                {
                    { 1, "Pelatihan/Workshop", "Pelatihan" },
                    { 2, "Kunjungan Kerja", "Kunjungan" }
                });

            migrationBuilder.InsertData(
                table: "MasterSertifikasis",
                columns: new[] { "Id", "Kategori", "Keterangan", "Nama" },
                values: new object[,]
                {
                    { 3, null, null, "First Aider" },
                    { 4, null, null, "Fire Fighting" },
                    { 5, null, null, "Confined Space" }
                });
        }
    }
}
