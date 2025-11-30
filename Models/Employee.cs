using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    public class Employee
    {
        public int Id { get; set; }

        // ===== A. Identitas =====
        [Required, StringLength(200)]
        public string NamaLengkap { get; set; } = string.Empty;

        [Required, StringLength(50)]
        public string NopegPersero { get; set; } = string.Empty;

        [StringLength(50)]
        public string? NopegKPI { get; set; }

        [StringLength(150)]
        public string? Jabatan { get; set; }

        [StringLength(150)]
        public string? TempatLahir { get; set; }

        [DataType(DataType.Date)]
        public DateTime? TanggalLahir { get; set; }

        [StringLength(20)]
        public string? JenisKelamin { get; set; }

        [StringLength(250)]
        public string? AlamatDomisili { get; set; }

        [StringLength(50)]
        public string? StatusPernikahan { get; set; }

        [StringLength(200)]
        public string? NamaSuamiIstri { get; set; }

        public int? JumlahAnak { get; set; }

        // ===== B. Kepegawaian =====
        [DataType(DataType.Date)]
        public DateTime? TanggalMPPK { get; set; }

        [DataType(DataType.Date)]
        public DateTime? TanggalPensiun { get; set; }

        /// <summary>
        /// Intake pertama: DISIMPAN SEBAGAI STRING / VARCHAR
        /// (contoh: "2019-01-01", "Jan 2018", "Oktober 2020", dsb).
        /// </summary>
        [StringLength(50)]
        public string? IntakePertama { get; set; }

        /// <summary>
        /// Intake / perpanjangan kontrak aktif sekarang (DateTime).
        /// </summary>
        [DataType(DataType.Date)]
        public DateTime? IntakeAktif { get; set; }

        [StringLength(100)]
        public string? ClusteringIntake { get; set; }

        // ===== C. Pendidikan & Ukuran =====
        [StringLength(100)]
        public string? PendidikanTerakhir { get; set; }

        [StringLength(100)]
        public string? Jurusan { get; set; }

        [StringLength(150)]
        public string? Institusi { get; set; }

        [StringLength(20)]
        public string? UkWearpack { get; set; }

        [StringLength(20)]
        public string? UkKemejaPutih { get; set; }

        [StringLength(20)]
        public string? UkKemejaProduk { get; set; }

        [StringLength(20)]
        public string? UkSepatuSafety { get; set; }

        // ===== D. Kontak =====
        [EmailAddress, StringLength(150)]
        public string? EmailPertamina { get; set; }

        [Phone, StringLength(50)]
        public string? NomorHP { get; set; }

        [StringLength(50)]
        public string? NomorTeleponRumah { get; set; }

        [Range(0, 365)]
        public int? CutiLimit { get; set; }

        // ===== Navigations =====
        public List<Dinas> DinasList { get; set; } = new();
        public List<Cuti> CutiList { get; set; } = new();
        public List<Sertifikasi> SertifikasiList { get; set; } = new();

        // ===== Properti Turunan (tidak disimpan di DB) =====

        [NotMapped]
        public int Usia
        {
            get
            {
                if (TanggalLahir is null) return 0;

                var days = (DateTime.Today - TanggalLahir.Value.Date).TotalDays;
                if (days <= 0) return 0;

                return (int)(days / 365.25);
            }
        }

        /// <summary>
        /// Masa kerja dalam tahun, dibulatkan ke bawah.
        /// Patokan utama = IntakePertama (string, diparse ke DateTime kalau bisa),
        /// fallback = IntakeAktif.
        /// </summary>
        [NotMapped]
        public int MasaKerjaTahun
        {
            get
            {
                DateTime? start = null;

                // 1. Coba parse IntakePertama (string / varchar)
                if (!string.IsNullOrWhiteSpace(IntakePertama))
                {
                    var s = IntakePertama.Trim();

                    if (DateTime.TryParse(s, out var d1))
                    {
                        start = d1;
                    }
                    else if (double.TryParse(s, out var oa)
                             && oa > 20000 && oa < 60000)
                    {
                        // Antisipasi kalau valuenya numeric OADate dari Excel
                        start = DateTime.FromOADate(oa);
                    }
                }

                // 2. Kalau tidak berhasil, fallback ke IntakeAktif
                if (start is null)
                    start = IntakeAktif;

                if (start is null)
                    return 0;

                var diffDays = (DateTime.Today - start.Value.Date).TotalDays;
                if (diffDays <= 0)
                    return 0;

                return (int)(diffDays / 365.25);
            }
        }
    }
}
