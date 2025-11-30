using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    public class Sertifikasi
    {
        public int Id { get; set; }
        [Required] public int EmployeeId { get; set; }
        public Employee? Employee { get; set; }

        public int? MasterSertifikasiId { get; set; }
        public MasterSertifikasi? MasterSertifikasi { get; set; }

        [Required, StringLength(200)] public string NamaSertifikasi { get; set; } = string.Empty;

        // =================================================================
        // PERBAIKAN: Mengganti semua DateTime? menjadi DateOnly?
        // =================================================================

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateOnly? TanggalMulai { get; set; }

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateOnly? TanggalSelesai { get; set; }

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateOnly? BerlakuDari { get; set; }

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        public DateOnly? BerlakuSampai { get; set; }

        [StringLength(200)] public string? Lokasi { get; set; }
        [StringLength(500)] public string? Keterangan { get; set; }

        // PERBAIKAN: Menyesuaikan perbandingan untuk DateOnly
        public bool IsExpired => BerlakuSampai is not null && BerlakuSampai.Value < DateOnly.FromDateTime(DateTime.Today);
    }
}