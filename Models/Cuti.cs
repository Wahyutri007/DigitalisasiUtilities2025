using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    public class Cuti
    {
        public int Id { get; set; }
        [Required] public int EmployeeId { get; set; }
        public Employee? Employee { get; set; }

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        [Required] public DateTime TanggalMulai { get; set; }

        [DataType(DataType.Date)]
        [Column(TypeName = "date")]
        [Required] public DateTime TanggalSelesai { get; set; }

        [Range(1, 365)] public int JumlahHari { get; set; }
        [StringLength(200)] public string? Tujuan { get; set; }
        [StringLength(500)] public string? Keterangan { get; set; }
    }
}
