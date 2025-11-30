using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class AddCutiLimitToEmployee : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalSelesai",
                table: "Sertifikasis",
                type: "date",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "datetime2",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalMulai",
                table: "Sertifikasis",
                type: "date",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "datetime2",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "BerlakuSampai",
                table: "Sertifikasis",
                type: "date",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "datetime2",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "BerlakuDari",
                table: "Sertifikasis",
                type: "date",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "datetime2",
                oldNullable: true);

            migrationBuilder.AddColumn<int>(
                name: "CutiLimit",
                table: "Employees",
                type: "int",
                nullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalPulang",
                table: "Dinas",
                type: "date",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "datetime2");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalBerangkat",
                table: "Dinas",
                type: "date",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "datetime2");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalSelesai",
                table: "Cutis",
                type: "date",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "datetime2");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalMulai",
                table: "Cutis",
                type: "date",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "datetime2");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CutiLimit",
                table: "Employees");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalSelesai",
                table: "Sertifikasis",
                type: "datetime2",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "date",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalMulai",
                table: "Sertifikasis",
                type: "datetime2",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "date",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "BerlakuSampai",
                table: "Sertifikasis",
                type: "datetime2",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "date",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "BerlakuDari",
                table: "Sertifikasis",
                type: "datetime2",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "date",
                oldNullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalPulang",
                table: "Dinas",
                type: "datetime2",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "date");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalBerangkat",
                table: "Dinas",
                type: "datetime2",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "date");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalSelesai",
                table: "Cutis",
                type: "datetime2",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "date");

            migrationBuilder.AlterColumn<DateTime>(
                name: "TanggalMulai",
                table: "Cutis",
                type: "datetime2",
                nullable: false,
                oldClrType: typeof(DateTime),
                oldType: "date");
        }
    }
}
