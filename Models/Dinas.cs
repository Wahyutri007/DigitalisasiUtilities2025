using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    public enum SifatDinas { PCU = 1, Utilities = 2 }

    public class Dinas
    {
        public int Id { get; set; }

        [Required]
        public int EmployeeId { get; set; }
        public Employee? Employee { get; set; }

        [Required]
        public int JenisDinasId { get; set; }
        public JenisDinas? JenisDinas { get; set; }

        // Sekarang nullable (opsional)
        public SifatDinas? Sifat { get; set; }

        // Lokasi opsional, ikut template import
        [StringLength(200)]
        public string? Lokasi { get; set; }

        [DataType(DataType.Date), Column(TypeName = "date")]
        [Required]
        public DateTime TanggalBerangkat { get; set; }

        [DataType(DataType.Date), Column(TypeName = "date")]
        [Required]
        public DateTime TanggalPulang { get; set; }

        [Required, StringLength(500)]
        public string Kegiatan { get; set; } = string.Empty;

        [NotMapped]
        public int LamaHari => (int)Math.Max(1, (TanggalPulang.Date - TanggalBerangkat.Date).TotalDays + 1);
    }

    public class JenisDinas
    {
        public int Id { get; set; }

        [Required, StringLength(120)]
        public string Nama { get; set; } = string.Empty;

        [StringLength(400)]
        public string? Keterangan { get; set; }

        public ICollection<Dinas> DinasList { get; set; } = new List<Dinas>();
    }

    // Ringkasan (Index ringkas)
    public sealed record GroupRow(int EmployeeId, string Nama, string Nopeg, int Jumlah)
    {
        public static GroupRow Create(int employeeId, string nama, string nopeg, int jumlah)
            => new(employeeId, nama, nopeg, jumlah);
    }
}
