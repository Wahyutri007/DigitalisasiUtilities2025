using UtilitiesHR.Models;

namespace UtilitiesHR.ViewModels
{
    public class EmployeeDetailVm
    {
        public Employee Employee { get; set; } = null!;

        // hanya dipakai buat info ringkasan
        public int TotalDinas { get; set; }
        public int TotalCuti { get; set; }
        public int TotalSertifikasi { get; set; }
    }
}
