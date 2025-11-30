using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Models;

namespace UtilitiesHR.Data
{
    // IMPORTANT: sekarang pakai IdentityDbContext<ApplicationUser>
    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        // ======= HR Entities =======
        public DbSet<Employee> Employees => Set<Employee>();
        public DbSet<Dinas> Dinas => Set<Dinas>();
        public DbSet<JenisDinas> JenisDinas => Set<JenisDinas>();
        public DbSet<Cuti> Cutis => Set<Cuti>();
        public DbSet<Sertifikasi> Sertifikasis => Set<Sertifikasi>();
        public DbSet<MasterSertifikasi> MasterSertifikasis => Set<MasterSertifikasi>();

        // ======= Material Utilities Entities =======
        public DbSet<JenisBarang> JenisBarangs => Set<JenisBarang>();
        public DbSet<Material> Materials => Set<Material>();
        public DbSet<MaterialTxn> MaterialTxns => Set<MaterialTxn>();

        // ======= Lookup untuk Employee (Jabatan / Pendidikan / dst) =======
        public DbSet<EmployeeLookup> EmployeeLookups => Set<EmployeeLookup>();

        // ======= Outstanding & Master Nama Pekerjaan =======
        public DbSet<WorkTask> WorkTasks => Set<WorkTask>();
        public DbSet<WorkJobName> WorkJobNames => Set<WorkJobName>();

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // PENTING: panggil base agar tabel Identity (AspNetUsers, dsb) tetap terkonfigurasi
            base.OnModelCreating(modelBuilder);

            // ===== Employee =====
            modelBuilder.Entity<Employee>(e =>
            {
                e.HasIndex(x => x.NopegPersero).IsUnique(false); // ada kemungkinan duplikat di template
                e.HasIndex(x => x.NopegKPI);

                e.Property(x => x.NamaLengkap)
                    .HasMaxLength(200)
                    .IsRequired();

                e.Property(x => x.EmailPertamina)
                    .HasMaxLength(200);
            });

            // ===== EmployeeLookup =====
            modelBuilder.Entity<EmployeeLookup>(l =>
            {
                l.Property(x => x.Category)
                    .HasMaxLength(40)
                    .IsRequired();

                l.Property(x => x.Value)
                    .HasMaxLength(200)
                    .IsRequired();

                l.HasIndex(x => new { x.Category, x.Value }).IsUnique();
            });

            // ===== Dinas =====
            modelBuilder.Entity<Dinas>(d =>
            {
                d.Property(x => x.Sifat)
                 .HasConversion<int?>()
                 .IsRequired(false);

                d.HasOne(x => x.Employee)
                 .WithMany(x => x.DinasList)
                 .HasForeignKey(x => x.EmployeeId)
                 .OnDelete(DeleteBehavior.Cascade);

                d.HasOne(x => x.JenisDinas)
                 .WithMany(x => x.DinasList)
                 .HasForeignKey(x => x.JenisDinasId)
                 .OnDelete(DeleteBehavior.Restrict);
            });

            // ===== JenisDinas =====
            modelBuilder.Entity<JenisDinas>(j =>
            {
                j.Property(x => x.Nama)
                 .HasMaxLength(120)
                 .IsRequired();

                j.Property(x => x.Keterangan)
                 .HasMaxLength(400);

                j.HasIndex(x => x.Nama);
            });

            // ===== Cuti =====
            modelBuilder.Entity<Cuti>(c =>
            {
                c.HasOne(x => x.Employee)
                 .WithMany(x => x.CutiList)
                 .HasForeignKey(x => x.EmployeeId)
                 .OnDelete(DeleteBehavior.Cascade);
            });

            // ===== Sertifikasi =====
            modelBuilder.Entity<Sertifikasi>(s =>
            {
                s.HasOne(x => x.Employee)
                 .WithMany(x => x.SertifikasiList)
                 .HasForeignKey(x => x.EmployeeId)
                 .OnDelete(DeleteBehavior.Cascade);

                s.HasOne(x => x.MasterSertifikasi)
                 .WithMany()
                 .HasForeignKey(x => x.MasterSertifikasiId)
                 .OnDelete(DeleteBehavior.SetNull);
            });

            // ===== Seed contoh Master Sertifikasi =====
            modelBuilder.Entity<MasterSertifikasi>().HasData(
                new MasterSertifikasi { Id = 1, Nama = "Operator Boiler" },
                new MasterSertifikasi { Id = 2, Nama = "Ahli K3 Umum" }
            );

            // ===== Material =====
            modelBuilder.Entity<Material>(m =>
            {
                m.HasIndex(x => x.KodeBarang).IsUnique();

                m.Property(x => x.JumlahBarang)
                 .HasDefaultValue(0);

                m.HasOne(x => x.JenisBarang)
                 .WithMany(j => j.Materials)
                 .HasForeignKey(x => x.JenisBarangId)
                 .OnDelete(DeleteBehavior.Restrict);
            });

            // ===== MaterialTxn =====
            modelBuilder.Entity<MaterialTxn>(t =>
            {
                t.HasOne(x => x.Material)
                 .WithMany(m => m.Txns)
                 .HasForeignKey(x => x.MaterialId)
                 .OnDelete(DeleteBehavior.Cascade);
            });

            // ===== Seed JenisBarang =====
            modelBuilder.Entity<JenisBarang>().HasData(
                new JenisBarang { Id = 1, Nama = "Sparepart", Keterangan = "Sparepart umum" },
                new JenisBarang { Id = 2, Nama = "Consumable", Keterangan = "Bahan habis pakai" }
            );

            // ===== WorkJobName (Master Nama Pekerjaan) =====
            modelBuilder.Entity<WorkJobName>(w =>
            {
                w.Property(x => x.NamaPekerjaan)
                 .HasMaxLength(200)
                 .IsRequired();

                // Biar tidak dobel-dobel nama sama
                w.HasIndex(x => x.NamaPekerjaan)
                 .IsUnique();
            });

            // ===== WorkTask (Outstanding Pekerjaan) =====
            modelBuilder.Entity<WorkTask>(w =>
            {
                w.Property(x => x.NamaPekerjaan)
                 .HasMaxLength(200)
                 .IsRequired();

                w.Property(x => x.NamaRequest)
                 .HasMaxLength(200);

                w.Property(x => x.NamaPIC)
                 .HasMaxLength(150);

                w.Property(x => x.KeteranganTindakLanjut)
                 .HasMaxLength(1000);

                // Index untuk dashboard by PIC & status
                w.HasIndex(x => new { x.NamaPIC, x.Status });
                w.HasIndex(x => x.Prioritas);
            });
        }
    }
}
