using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    // ======= Material Models =======
    public enum JenisTransaksi { Masuk = 1, Keluar = 2 }

    public class JenisBarang
    {
        public int Id { get; set; }
        
        [Required, StringLength(150)] 
        public string Nama { get; set; } = string.Empty;
        
        [StringLength(300)] 
        public string? Keterangan { get; set; }

        // Navigation property untuk relationship dengan Material
        public ICollection<Material> Materials { get; set; } = new List<Material>();
    }

    public class Material
    {
        public int Id { get; set; }

        [Required, StringLength(50)]
        public string KodeBarang { get; set; } = string.Empty;

        [Required, StringLength(200)]
        public string NamaBarang { get; set; } = string.Empty;

        [Required]
        public int JenisBarangId { get; set; }
        
        [ForeignKey("JenisBarangId")]
        public JenisBarang? JenisBarang { get; set; }

        [StringLength(150)]
        public string? PosisiBarang { get; set; }

        // stok berjalan
        public int JumlahBarang { get; set; } = 0;

        [StringLength(30)]
        public string? Satuan { get; set; } = "unit";

        // Navigation property untuk relationship dengan MaterialTxn
        public ICollection<MaterialTxn> Txns { get; set; } = new List<MaterialTxn>();
    }

    public class MaterialTxn
    {
        public int Id { get; set; }

        [Required]
        public int MaterialId { get; set; }
        
        [ForeignKey("MaterialId")]
        public Material? Material { get; set; }

        [Required]
        public JenisTransaksi Jenis { get; set; } = JenisTransaksi.Masuk;

        [DataType(DataType.Date)]
        public DateTime Tanggal { get; set; } = DateTime.Today;

        [Range(1, int.MaxValue)]
        public int Jumlah { get; set; }

        [StringLength(150)]
        public string? PenanggungJawab { get; set; }

        [StringLength(500)]
        public string? Keterangan { get; set; }
    }
}