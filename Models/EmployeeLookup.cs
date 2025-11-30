namespace UtilitiesHR.Models
{
    public class EmployeeLookup
    {
        public int Id { get; set; }

        // Contoh category: "Jabatan", "Pendidikan", "Jurusan", "Institusi", "StatusNikah"
        public string Category { get; set; } = string.Empty;

        public string Value { get; set; } = string.Empty;
    }
}
