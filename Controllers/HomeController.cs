using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using System.Text.RegularExpressions;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    // Item label + value sederhana untuk grafik
    public class SimpleCountItem
    {
        public string Label { get; set; } = "";
        public int Value { get; set; }
    }

    // ViewModel Dashboard
    public class HomeDashboardVM
    {
        // ==== SUMMARY CARD ====
        public int TotalEmployee { get; set; }
        public int TotalDinas { get; set; }
        public int TotalDinasAktif { get; set; }
        public int TotalDinasNonAktif { get; set; }
        public int TotalCuti { get; set; }
        public int TotalCutiAktif { get; set; }
        public int TotalCutiSelesai { get; set; }
        public int TotalMaterial { get; set; }
        public int TotalSertifikasi { get; set; }
        public int TotalSertifikasiAktif { get; set; }
        public int TotalSertifikasiExpired { get; set; }
        public int TotalUsers { get; set; }

        // Untuk user biasa (data pribadi)
        public int MyTotalSertifikasi { get; set; }
        public int MyTotalSertifikasiAktif { get; set; }
        public int MyTotalSertifikasiExpired { get; set; }
        public int MyTotalDinas { get; set; }
        public int MyTotalCuti { get; set; }

        // ==== GRAFIK PEGAWAI ====
        public List<SimpleCountItem> EmployeesByPendidikan { get; set; } = new();
        public List<SimpleCountItem> EmployeesByJobType { get; set; } = new();

        // GRAFIK TREN REKRUTMEN (lama – gabungan)
        public List<SimpleCountItem> EmployeesByIntakeYear { get; set; } = new();

        // Baru: Intake pertama & Intake aktif per tahun
        public List<SimpleCountItem> EmployeesByIntakePertamaYear { get; set; } = new();
        public List<SimpleCountItem> EmployeesByIntakeAktifYear { get; set; } = new();

        // ==== GRAFIK DINAS / CUTI ====
        public List<SimpleCountItem> DinasByJenis { get; set; } = new();
        public List<SimpleCountItem> CutiByYear { get; set; } = new();

        // ==== GRAFIK MATERIAL & SERTIFIKASI ====
        public List<SimpleCountItem> MaterialsTopStock { get; set; } = new();
        public List<SimpleCountItem> SertifikasiByNama { get; set; } = new();

        // ==== OUTSTANDING WORKTASKS ====
        public int TotalWorkTasks { get; set; }
        public int TotalOutstandingWorkTasks { get; set; }
        public List<SimpleCountItem> WorkTasksByStatus { get; set; } = new();
        public List<SimpleCountItem> WorkTasksByPriority { get; set; } = new();

        // BARU: NAMA PEKERJAAN TERBANYAK (TOP 10)
        public List<SimpleCountItem> WorkTasksByJobName { get; set; } = new();
    }

    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;

        public HomeController(
            ILogger<HomeController> logger,
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager)
        {
            _logger = logger;
            _db = db;
            _userManager = userManager;
        }

        private bool IsAdminLike =>
            User.IsInRole("Admin") ||
            User.IsInRole("SuperAdmin") ||
            User.IsInRole("Supervisor");

        private bool IsUserRole => User.IsInRole("User");

        // Helper: ambil tahun 4 digit dari string IntakePertama (contoh: "BK JT 2015" -> 2015)
        private static int? ExtractYearFromIntake(string? intakeText)
        {
            if (string.IsNullOrWhiteSpace(intakeText))
                return null;

            var match = Regex.Match(intakeText, @"(19|20)\d{2}");
            if (match.Success && int.TryParse(match.Value, out var year))
                return year;

            return null;
        }

        public async Task<IActionResult> Index()
        {
            var vm = new HomeDashboardVM();
            var today = DateTime.Today;

            var employeesQ = _db.Employees.AsQueryable();
            var dinasQ = _db.Dinas.AsQueryable();
            var cutiQ = _db.Cutis.AsQueryable();
            var tasksQ = _db.WorkTasks.AsQueryable(); // Outstanding pekerjaan

            var allSert = await _db.Sertifikasis
                .Include(s => s.MasterSertifikasi)
                .AsNoTracking()
                .ToListAsync();

            ApplicationUser? appUser = null;
            int? empId = null;

            // ===== FILTER UNTUK ROLE USER BIASA (hanya data sendiri) =====
            if (IsUserRole)
            {
                appUser = await _userManager.GetUserAsync(User);
                empId = appUser?.EmployeeId;

                if (empId.HasValue)
                {
                    employeesQ = employeesQ.Where(e => e.Id == empId.Value);
                    dinasQ = dinasQ.Where(d => d.EmployeeId == empId.Value);
                    cutiQ = cutiQ.Where(c => c.EmployeeId == empId.Value);
                    allSert = allSert.Where(s => s.EmployeeId == empId.Value).ToList();
                    // WorkTasks tetap global
                }
            }

            var employeesList = await employeesQ
                .AsNoTracking()
                .ToListAsync();

            // ===== SUMMARY PEGAWAI & USER =====
            if (IsAdminLike)
            {
                vm.TotalEmployee = await _db.Employees.CountAsync();
                vm.TotalUsers = await _db.Users.CountAsync();
            }
            else
            {
                vm.TotalEmployee = employeesList.Count;
                vm.TotalUsers = 1;
            }

            // ===== SUMMARY SERTIFIKASI =====
            vm.TotalSertifikasi = allSert.Count;
            vm.TotalSertifikasiExpired = allSert.Count(s => s.IsExpired);
            vm.TotalSertifikasiAktif = vm.TotalSertifikasi - vm.TotalSertifikasiExpired;

            if (empId.HasValue)
            {
                vm.MyTotalSertifikasi = vm.TotalSertifikasi;
                vm.MyTotalSertifikasiAktif = vm.TotalSertifikasiAktif;
                vm.MyTotalSertifikasiExpired = vm.TotalSertifikasiExpired;
            }

            // ===== SUMMARY MATERIAL =====
            vm.TotalMaterial = await _db.Materials.CountAsync();

            // ===== SUMMARY DINAS =====
            vm.TotalDinas = await dinasQ.CountAsync();
            vm.MyTotalDinas = vm.TotalDinas;

            vm.TotalDinasAktif = await dinasQ.CountAsync(d =>
                d.TanggalBerangkat <= today &&
                d.TanggalPulang >= today
            );
            vm.TotalDinasNonAktif = vm.TotalDinas - vm.TotalDinasAktif;

            // ===== SUMMARY CUTI =====
            vm.TotalCuti = await cutiQ.CountAsync();
            vm.MyTotalCuti = vm.TotalCuti;

            vm.TotalCutiAktif = await cutiQ.CountAsync(c =>
                c.TanggalMulai <= today &&
                c.TanggalSelesai >= today
            );
            vm.TotalCutiSelesai = vm.TotalCuti - vm.TotalCutiAktif;

            // ===== GRAFIK PEGAWAI: Pendidikan =====
            vm.EmployeesByPendidikan = employeesList
                .GroupBy(e => string.IsNullOrWhiteSpace(e.PendidikanTerakhir)
                    ? "Lainnya"
                    : e.PendidikanTerakhir!)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .ToList();

            // ===== GRAFIK PEGAWAI: Jabatan / Jenis Pekerjaan =====
            vm.EmployeesByJobType = employeesList
                .GroupBy(e => string.IsNullOrWhiteSpace(e.Jabatan)
                    ? "Lainnya"
                    : e.Jabatan!)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .ToList();

            // ===== BARU: GRAFIK PEGAWAI – Intake Pertama per Tahun =====
            vm.EmployeesByIntakePertamaYear = employeesList
                .Select(e => ExtractYearFromIntake(e.IntakePertama))
                .Where(y => y.HasValue)
                .GroupBy(y => y!.Value)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key.ToString(),
                    Value = g.Count()
                })
                .OrderBy(x => x.Label)
                .ToList();

            // ===== BARU: GRAFIK PEGAWAI – Intake Aktif per Tahun =====
            vm.EmployeesByIntakeAktifYear = employeesList
                .Where(e => e.IntakeAktif.HasValue)
                .GroupBy(e => e.IntakeAktif!.Value.Year)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key.ToString(),
                    Value = g.Count()
                })
                .OrderBy(x => x.Label)
                .ToList();

            // OPSIONAL: gabungan lama
            vm.EmployeesByIntakeYear = vm.EmployeesByIntakePertamaYear
                .Concat(vm.EmployeesByIntakeAktifYear)
                .GroupBy(x => x.Label)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Sum(i => i.Value)
                })
                .OrderBy(x => x.Label)
                .ToList();

            // ===== GRAFIK DINAS: per Jenis Dinas =====
            vm.DinasByJenis = await dinasQ
                .Include(d => d.JenisDinas)
                .GroupBy(d => d.JenisDinas != null
                    ? (d.JenisDinas.Nama ?? "Lainnya")
                    : "Lainnya")
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .Take(10)
                .ToListAsync();

            // ===== GRAFIK CUTI: per Tahun =====
            vm.CutiByYear = await cutiQ
                .GroupBy(c => c.TanggalMulai.Year)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key.ToString(),
                    Value = g.Count()
                })
                .OrderBy(x => x.Label)
                .ToListAsync();

            // ===== GRAFIK MATERIAL: Top stok =====
            vm.MaterialsTopStock = await _db.Materials
                .OrderByDescending(m => m.JumlahBarang)
                .Select(m => new SimpleCountItem
                {
                    Label = m.NamaBarang,
                    Value = m.JumlahBarang
                })
                .Take(10)
                .ToListAsync();

            // ===== GRAFIK SERTIFIKASI: Nama sertifikasi vs jumlah orang =====
            vm.SertifikasiByNama = allSert
                .GroupBy(s => s.NamaSertifikasi ?? s.MasterSertifikasi?.Nama ?? "(tanpa nama)")
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .ToList();

            // ===== OUTSTANDING WORKTASKS (STATUS & PRIORITAS) =====
            vm.TotalWorkTasks = await tasksQ.CountAsync();
            vm.TotalOutstandingWorkTasks = await tasksQ.CountAsync(t => t.Status != WorkTaskStatus.Done);

            vm.WorkTasksByStatus = await tasksQ
                .GroupBy(t => t.Status)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key.ToString(),
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .ToListAsync();

            vm.WorkTasksByPriority = await tasksQ
                .GroupBy(t => t.Prioritas)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key.ToString(),
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .ToListAsync();

            // ===== BARU: OUTSTANDING WORKTASKS – NAMA PEKERJAAN TERBANYAK (TOP 10) =====
            vm.WorkTasksByJobName = await tasksQ
                .GroupBy(t => string.IsNullOrWhiteSpace(t.NamaPekerjaan)
                    ? "(Tanpa nama)"
                    : t.NamaPekerjaan!)
                .Select(g => new SimpleCountItem
                {
                    Label = g.Key,
                    Value = g.Count()
                })
                .OrderByDescending(x => x.Value)
                .Take(10)
                .ToListAsync();

            return View(vm);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel
            {
                RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
            });
        }
    }
}
