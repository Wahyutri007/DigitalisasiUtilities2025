using System.ComponentModel.DataAnnotations;

namespace UtilitiesHR.Models
{
    public class MasterSertifikasi
    {
        public int Id { get; set; }
        [Required, StringLength(200)] public string Nama { get; set; } = string.Empty;
        [StringLength(100)] public string? Kategori { get; set; }
        [StringLength(300)] public string? Keterangan { get; set; }
    }
}
