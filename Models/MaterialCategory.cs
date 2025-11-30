using System.ComponentModel.DataAnnotations;

namespace UtilitiesHR.Models
{
    public class MaterialCategory
    {
        public int Id { get; set; }
        [Required, StringLength(120)] public string Nama { get; set; } = string.Empty;
        [StringLength(300)] public string? Keterangan { get; set; }

        public List<Material> Materials { get; set; } = new();
    }
}
