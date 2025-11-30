using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace UtilitiesHR.Migrations
{
    /// <inheritdoc />
    public partial class Fix_Material_Baseline_Safe : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            // ===== 0) NETRALISIR BLOK AUTOGENERATE YG DROP INDEX / REBIND DEFAULT =====
            // Jika ada index IX_Materials_JenisBarangId, drop-lah dgn guard; kalau tidak ada, skip.
            migrationBuilder.Sql(@"
IF EXISTS (
    SELECT 1 FROM sys.indexes 
    WHERE name = 'IX_Materials_JenisBarangId' AND object_id = OBJECT_ID('dbo.Materials')
)
    DROP INDEX [IX_Materials_JenisBarangId] ON [dbo].[Materials];
");

            // ===== 1) Master JenisBarangs — buat kalau belum ada + seed minimal =====
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.JenisBarangs','U') IS NULL
BEGIN
    CREATE TABLE dbo.JenisBarangs
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_JenisBarangs PRIMARY KEY,
        Nama NVARCHAR(100) NOT NULL,
        Keterangan NVARCHAR(200) NULL
    );

    INSERT INTO dbo.JenisBarangs(Nama, Keterangan)
    VALUES (N'Sparepart', N'Sparepart umum'),
           (N'Consumable', N'Bahan habis pakai');
END
");

            // ===== 2) Materials — buat kalau belum ada =====
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.Materials','U') IS NULL
BEGIN
    CREATE TABLE dbo.Materials
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Materials PRIMARY KEY,
        KodeBarang    NVARCHAR(50)  NOT NULL,
        NamaBarang    NVARCHAR(150) NOT NULL,
        PosisiBarang  NVARCHAR(120) NULL,
        JumlahBarang  INT           NOT NULL CONSTRAINT DF_Materials_JumlahBarang DEFAULT(0),
        Satuan        NVARCHAR(30)  NOT NULL CONSTRAINT DF_Materials_Satuan DEFAULT(N'unit'),
        JenisBarangId INT           NOT NULL
    );
END
");

            // ===== 2a) Pastikan kolom/tipe/index/constraint pada Materials =====
            migrationBuilder.Sql(@"
-- KodeBarang + unique index
IF COL_LENGTH('dbo.Materials','KodeBarang') IS NULL
    EXEC('ALTER TABLE dbo.Materials ADD KodeBarang NVARCHAR(50) NOT NULL CONSTRAINT DF_TMP_Kode DEFAULT(N'''')');
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='UX_Materials_KodeBarang' AND object_id=OBJECT_ID('dbo.Materials'))
    CREATE UNIQUE INDEX UX_Materials_KodeBarang ON dbo.Materials(KodeBarang);

-- NamaBarang
IF COL_LENGTH('dbo.Materials','NamaBarang') IS NULL
    EXEC('ALTER TABLE dbo.Materials ADD NamaBarang NVARCHAR(150) NOT NULL CONSTRAINT DF_TMP_Nama DEFAULT(N'''')');
ELSE
    EXEC('ALTER TABLE dbo.Materials ALTER COLUMN NamaBarang NVARCHAR(150) NOT NULL');

-- PosisiBarang
IF COL_LENGTH('dbo.Materials','PosisiBarang') IS NULL
    EXEC('ALTER TABLE dbo.Materials ADD PosisiBarang NVARCHAR(120) NULL');

-- JumlahBarang + default
IF COL_LENGTH('dbo.Materials','JumlahBarang') IS NULL
    EXEC('ALTER TABLE dbo.Materials ADD JumlahBarang INT NOT NULL CONSTRAINT DF_Materials_JumlahBarang DEFAULT(0)');
ELSE
BEGIN
    DECLARE @dc_jumlah sysname;
    SELECT @dc_jumlah = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.default_object_id = dc.object_id
    WHERE dc.parent_object_id = OBJECT_ID('dbo.Materials') AND c.name = 'JumlahBarang';
    IF @dc_jumlah IS NULL
        EXEC('ALTER TABLE dbo.Materials ADD CONSTRAINT DF_Materials_JumlahBarang DEFAULT(0) FOR JumlahBarang');
END

-- Satuan + default
IF COL_LENGTH('dbo.Materials','Satuan') IS NULL
BEGIN
    EXEC('ALTER TABLE dbo.Materials ADD Satuan NVARCHAR(30) NULL');
    EXEC('UPDATE dbo.Materials SET Satuan = N''unit'' WHERE Satuan IS NULL');
    EXEC('ALTER TABLE dbo.Materials ALTER COLUMN Satuan NVARCHAR(30) NOT NULL');
    IF NOT EXISTS (
        SELECT 1 FROM sys.default_constraints dc
        JOIN sys.columns c ON c.default_object_id = dc.object_id
        WHERE dc.parent_object_id = OBJECT_ID('dbo.Materials') AND c.name = 'Satuan'
    )
        EXEC('ALTER TABLE dbo.Materials ADD CONSTRAINT DF_Materials_Satuan DEFAULT(N''unit'') FOR Satuan');
END
ELSE
BEGIN
    DECLARE @dc_satuan sysname;
    SELECT @dc_satuan = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.default_object_id = dc.object_id
    WHERE dc.parent_object_id = OBJECT_ID('dbo.Materials') AND c.name = 'Satuan';
    IF @dc_satuan IS NULL
        EXEC('ALTER TABLE dbo.Materials ADD CONSTRAINT DF_Materials_Satuan DEFAULT(N''unit'') FOR Satuan');
    EXEC('ALTER TABLE dbo.Materials ALTER COLUMN Satuan NVARCHAR(30) NOT NULL');
END

-- JenisBarangId (NOT NULL + default 1) — JANGAN re-bind kalau default sudah ada
IF COL_LENGTH('dbo.Materials','JenisBarangId') IS NULL
BEGIN
    EXEC('ALTER TABLE dbo.Materials ADD JenisBarangId INT NOT NULL DEFAULT(1)');
END
ELSE
BEGIN
    EXEC('UPDATE dbo.Materials SET JenisBarangId = ISNULL(JenisBarangId,1)');
    DECLARE @dc_jenis sysname;
    SELECT @dc_jenis = dc.name
    FROM sys.default_constraints dc
    JOIN sys.columns c ON c.default_object_id = dc.object_id
    WHERE dc.parent_object_id = OBJECT_ID('dbo.Materials') AND c.name = 'JenisBarangId';
    IF @dc_jenis IS NULL
        EXEC('ALTER TABLE dbo.Materials ADD CONSTRAINT DF_Materials_JenisBarangId DEFAULT(1) FOR JenisBarangId');
END
");

            // Index untuk JenisBarangId (optional)
            migrationBuilder.Sql(@"
IF NOT EXISTS (
    SELECT 1 FROM sys.indexes WHERE name='IX_Materials_JenisBarangId' AND object_id=OBJECT_ID('dbo.Materials')
)
    CREATE INDEX IX_Materials_JenisBarangId ON dbo.Materials(JenisBarangId);
");

            // ===== 2b) FK Materials → JenisBarangs (buat jika belum ada) =====
            migrationBuilder.Sql(@"
IF NOT EXISTS (
    SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_Materials_JenisBarangs_JenisBarangId'
)
BEGIN
    ALTER TABLE dbo.Materials WITH NOCHECK
        ADD CONSTRAINT FK_Materials_JenisBarangs_JenisBarangId
        FOREIGN KEY (JenisBarangId) REFERENCES dbo.JenisBarangs(Id)
        ON DELETE NO ACTION;
END
");

            // ===== 3) MaterialTxns — buat kalau belum ada =====
            migrationBuilder.Sql(@"
IF OBJECT_ID('dbo.MaterialTxns','U') IS NULL
BEGIN
    CREATE TABLE dbo.MaterialTxns
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_MaterialTxns PRIMARY KEY,
        MaterialId INT NOT NULL,
        Jenis INT NOT NULL,           -- 1=Masuk, 2=Keluar
        Tanggal DATE NOT NULL,
        Jumlah INT NOT NULL,
        PenanggungJawab NVARCHAR(120) NULL,
        Keterangan NVARCHAR(300) NULL
    );

    ALTER TABLE dbo.MaterialTxns
        ADD CONSTRAINT FK_MaterialTxns_Materials_MaterialId
        FOREIGN KEY (MaterialId) REFERENCES dbo.Materials(Id)
        ON DELETE CASCADE;

    CREATE INDEX IX_MaterialTxns_MaterialId ON dbo.MaterialTxns(MaterialId);
    CREATE INDEX IX_MaterialTxns_Tanggal   ON dbo.MaterialTxns(Tanggal);
END
");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {

        }
    }
}
