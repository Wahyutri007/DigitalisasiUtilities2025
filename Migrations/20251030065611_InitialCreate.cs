using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Employees",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NamaLengkap = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    NopegPersero = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false),
                    NopegKPI = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    Jabatan = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true),
                    TempatLahir = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true),
                    TanggalLahir = table.Column<DateTime>(type: "datetime2", nullable: true),
                    JenisKelamin = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    AlamatDomisili = table.Column<string>(type: "nvarchar(250)", maxLength: 250, nullable: true),
                    StatusPernikahan = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    NamaSuamiIstri = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    JumlahAnak = table.Column<int>(type: "int", nullable: true),
                    TanggalMPPK = table.Column<DateTime>(type: "datetime2", nullable: true),
                    TanggalPensiun = table.Column<DateTime>(type: "datetime2", nullable: true),
                    IntakePertama = table.Column<DateTime>(type: "datetime2", nullable: true),
                    IntakeAktif = table.Column<DateTime>(type: "datetime2", nullable: true),
                    ClusteringIntake = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    PendidikanTerakhir = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Jurusan = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Institusi = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true),
                    UkWearpack = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    UkKemejaPutih = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    UkKemejaProduk = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    UkSepatuSafety = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    EmailPertamina = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: true),
                    NomorHP = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    NomorTeleponRumah = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Employees", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "MasterSertifikasis",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Nama = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    Kategori = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Keterangan = table.Column<string>(type: "nvarchar(300)", maxLength: 300, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_MasterSertifikasis", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Cutis",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    EmployeeId = table.Column<int>(type: "int", nullable: false),
                    TanggalMulai = table.Column<DateTime>(type: "datetime2", nullable: false),
                    TanggalSelesai = table.Column<DateTime>(type: "datetime2", nullable: false),
                    JumlahHari = table.Column<int>(type: "int", nullable: false),
                    Tujuan = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    Keterangan = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Cutis", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Cutis_Employees_EmployeeId",
                        column: x => x.EmployeeId,
                        principalTable: "Employees",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "Dinas",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    EmployeeId = table.Column<int>(type: "int", nullable: false),
                    Sifat = table.Column<int>(type: "int", nullable: false),
                    TanggalBerangkat = table.Column<DateTime>(type: "datetime2", nullable: false),
                    TanggalPulang = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Kegiatan = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Dinas", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Dinas_Employees_EmployeeId",
                        column: x => x.EmployeeId,
                        principalTable: "Employees",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "Sertifikasis",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    EmployeeId = table.Column<int>(type: "int", nullable: false),
                    MasterSertifikasiId = table.Column<int>(type: "int", nullable: true),
                    NamaSertifikasi = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: false),
                    TanggalMulai = table.Column<DateTime>(type: "datetime2", nullable: true),
                    TanggalSelesai = table.Column<DateTime>(type: "datetime2", nullable: true),
                    BerlakuDari = table.Column<DateTime>(type: "datetime2", nullable: true),
                    BerlakuSampai = table.Column<DateTime>(type: "datetime2", nullable: true),
                    Lokasi = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    Keterangan = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Sertifikasis", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Sertifikasis_Employees_EmployeeId",
                        column: x => x.EmployeeId,
                        principalTable: "Employees",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                    table.ForeignKey(
                        name: "FK_Sertifikasis_MasterSertifikasis_MasterSertifikasiId",
                        column: x => x.MasterSertifikasiId,
                        principalTable: "MasterSertifikasis",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                });

            migrationBuilder.InsertData(
                table: "MasterSertifikasis",
                columns: new[] { "Id", "Kategori", "Keterangan", "Nama" },
                values: new object[,]
                {
                    { 1, null, null, "Operator Boiler" },
                    { 2, null, null, "Ahli K3 Umum" },
                    { 3, null, null, "First Aider" },
                    { 4, null, null, "Fire Fighting" },
                    { 5, null, null, "Confined Space" }
                });

            migrationBuilder.CreateIndex(
                name: "IX_Cutis_EmployeeId",
                table: "Cutis",
                column: "EmployeeId");

            migrationBuilder.CreateIndex(
                name: "IX_Dinas_EmployeeId",
                table: "Dinas",
                column: "EmployeeId");

            migrationBuilder.CreateIndex(
                name: "IX_Employees_NopegKPI",
                table: "Employees",
                column: "NopegKPI");

            migrationBuilder.CreateIndex(
                name: "IX_Employees_NopegPersero",
                table: "Employees",
                column: "NopegPersero",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_Sertifikasis_EmployeeId",
                table: "Sertifikasis",
                column: "EmployeeId");

            migrationBuilder.CreateIndex(
                name: "IX_Sertifikasis_MasterSertifikasiId",
                table: "Sertifikasis",
                column: "MasterSertifikasiId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Cutis");

            migrationBuilder.DropTable(
                name: "Dinas");

            migrationBuilder.DropTable(
                name: "Sertifikasis");

            migrationBuilder.DropTable(
                name: "Employees");

            migrationBuilder.DropTable(
                name: "MasterSertifikasis");
        }
    }
}
