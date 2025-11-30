namespace UtilitiesHR.Models
{
    // FIX: DTO/VM dipindahkan ke namespace Model

    public class TxnDto
    {
        public int Id { get; set; }
        public int MaterialId { get; set; }
        public JenisTransaksi Jenis { get; set; }
        public DateTime Tanggal { get; set; } = DateTime.Today;
        public int Jumlah { get; set; }
        public string? PenanggungJawab { get; set; }
        public string? Keterangan { get; set; }
    }

    public class MaterialDetailsVM
    {
        public int Id { get; set; }
        public string KodeBarang { get; set; } = "";
        public string NamaBarang { get; set; } = "";
        public string JenisBarang { get; set; } = "";
        public string? PosisiBarang { get; set; }
        public int Stok { get; set; }
        public string? Satuan { get; set; }
        public int MasukBulanIni { get; set; }
        public int KeluarBulanIni { get; set; }
        public int TotalMasuk { get; set; }
        public int TotalKeluar { get; set; }
        public List<MaterialTxnRow> Txns { get; set; } = []; // FIX: Inisialisasi sederhana
    }

    public class MaterialTxnRow
    {
        public int Id { get; set; }
        public DateTime Tanggal { get; set; }
        public JenisTransaksi Jenis { get; set; }
        public int Jumlah { get; set; }
        public string? PenanggungJawab { get; set; }
        public string? Keterangan { get; set; }
        public int SisaStok { get; set; }
    }
}