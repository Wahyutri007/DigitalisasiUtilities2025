using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using UtilitiesHR.ViewModels;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class EmployeesController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly ILogger<EmployeesController> _logger;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly RoleManager<IdentityRole> _roleManager;

        private const string DefaultPassword = "Pertamina@2025";

        static EmployeesController()
        {
            ExcelPackage.License.SetNonCommercialPersonal("UtilitiesHR Dev");
        }

        public EmployeesController(
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager,
            RoleManager<IdentityRole> roleManager,
            ILogger<EmployeesController> logger)
        {
            _db = db;
            _userManager = userManager;
            _roleManager = roleManager;
            _logger = logger;
        }

        // ================== HELPER ROLE ==================
        private bool IsAdminLike()
        {
            return User.IsInRole("Admin")
                   || User.IsInRole("SuperAdmin")
                   || User.IsInRole("Supervisor");
        }

        // ================== HELPER: SYNC & AUTO USER ==================

        private async Task CreateUserFromEmployeeIfNeededAsync(Employee emp)
        {
            if (emp == null) return;

            if (string.IsNullOrWhiteSpace(emp.EmailPertamina))
            {
                _logger.LogInformation("Employee {EmpId} tidak punya EmailPertamina, skip auto user.", emp.Id);
                return;
            }

            var email = emp.EmailPertamina.Trim();
            var normalized = _userManager.NormalizeEmail(email);

            var existing = await _userManager.Users
                .FirstOrDefaultAsync(u => u.EmployeeId == emp.Id ||
                                          u.NormalizedEmail == normalized);

            if (existing != null)
            {
                _logger.LogInformation("AutoUser: user sudah ada untuk EmpId={EmpId} / Email={Email}, skip.", emp.Id, email);
                return;
            }

            var user = new ApplicationUser
            {
                Email = email,
                UserName = email,
                FullName = emp.NamaLengkap ?? email,
                EmployeeId = emp.Id,
                EmailConfirmed = true,
                MustChangePassword = true,          // 🔐 wajib ganti password pada login pertama
                LastPasswordChangedAt = null,       // belum pernah ganti
                LockoutEnabled = true
            };

            var createResult = await _userManager.CreateAsync(user, DefaultPassword);
            if (!createResult.Succeeded)
            {
                _logger.LogWarning("AutoUser: gagal create user untuk EmpId={EmpId}: {Errors}",
                    emp.Id,
                    string.Join("; ", createResult.Errors.Select(e => e.Description)));
                return;
            }

            if (!await _roleManager.RoleExistsAsync("User"))
            {
                await _roleManager.CreateAsync(new IdentityRole("User"));
            }

            await _userManager.AddToRoleAsync(user, "User");

            _logger.LogInformation("AutoUser: berhasil buat akun untuk EmpId={EmpId}, Email={Email}.", emp.Id, email);
        }

        private async Task SyncUserFromEmployeeAsync(Employee emp)
        {
            if (emp == null) return;

            var user = await _userManager.Users
                .FirstOrDefaultAsync(u => u.EmployeeId == emp.Id);

            if (user == null)
            {
                return;
            }

            user.FullName = emp.NamaLengkap;

            if (!string.IsNullOrWhiteSpace(emp.EmailPertamina))
            {
                var email = emp.EmailPertamina.Trim();
                var normalized = _userManager.NormalizeEmail(email);

                var other = await _userManager.Users
                    .FirstOrDefaultAsync(u =>
                        u.Id != user.Id &&
                        u.NormalizedEmail == normalized);

                if (other == null)
                {
                    user.Email = email;
                    user.UserName = email;
                }
                else
                {
                    _logger.LogWarning(
                        "SyncUser: Email {Email} sudah dipakai UserId={OtherId}, user EmployeeId={EmpId} tidak diubah emailnya.",
                        email, other.Id, emp.Id);
                }
            }

            var result = await _userManager.UpdateAsync(user);
            if (!result.Succeeded)
            {
                _logger.LogWarning(
                    "SyncUser: gagal update user untuk EmployeeId={EmpId}: {Errors}",
                    emp.Id,
                    string.Join("; ", result.Errors.Select(e => e.Description)));
            }
        }

        // ============= API BASIC INFO (untuk Create User manual) =============
        [HttpGet]
        public async Task<IActionResult> GetBasicInfo(int id)
        {
            var emp = await _db.Employees
                .AsNoTracking()
                .Where(e => e.Id == id)
                .Select(e => new
                {
                    e.Id,
                    Nama = e.NamaLengkap,
                    Nopeg = e.NopegPersero,
                    Email = e.EmailPertamina
                })
                .FirstOrDefaultAsync();

            if (emp == null) return NotFound();

            return Json(emp);
        }

        // ================== INDEX ==================
        public async Task<IActionResult> Index(string? q, int? firstEditId)
        {
            IQueryable<Employee> query = _db.Employees.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(e =>
                    (e.NamaLengkap ?? "").Contains(q) ||
                    (e.NopegPersero ?? "").Contains(q) ||
                    (e.NopegKPI ?? "").Contains(q) ||
                    (e.EmailPertamina ?? "").Contains(q)
                );
            }

            bool isAdminLike = IsAdminLike();
            bool isUserRole = User.IsInRole("User");
            int? currentEmpId = null;

            // cari EmployeeId dari user yang sedang login (kalau ada)
            var appUser = await _userManager.GetUserAsync(User);
            if (appUser?.EmployeeId is int empId && empId > 0)
            {
                currentEmpId = empId;
            }

            // ⛔️ SEBELUMNYA: user biasa hanya boleh lihat dirinya sendiri (query = query.Where(e => e.Id == empId))
            // ✅ SEKARANG: user biasa boleh lihat semua pekerja, jadi query tidak dibatasi
            // (kontrol edit dilakukan di view + di action Edit)

            ViewBag.Search = q;
            ViewBag.TotalEmployee = await _db.Employees.CountAsync();
            ViewBag.TotalEmail = await _db.Employees.CountAsync(e => !string.IsNullOrEmpty(e.EmailPertamina));
            ViewBag.TotalPhone = await _db.Employees.CountAsync(e => !string.IsNullOrEmpty(e.NomorHP));

            ViewBag.CurrentEmployeeId = currentEmpId ?? 0;
            ViewBag.IsAdminLike = isAdminLike;
            ViewBag.IsUserRole = isUserRole;

            var items = await query
                .OrderBy(e => e.NamaLengkap)
                .ToListAsync();

            // urutkan supaya baris dirinya sendiri muncul di baris pertama
            if (currentEmpId.HasValue)
            {
                items = items
                    .OrderByDescending(e => e.Id == currentEmpId.Value) // true dulu, lalu yg lain
                    .ThenBy(e => e.NamaLengkap)
                    .ToList();
            }

            // kirim id ke view → dipakai JS untuk auto-buka popup edit (kalau perlu)
            ViewBag.FirstEditId = firstEditId ?? 0;

            return View(items);
        }

        // ============= DETAIL HALAMAN PENUH =============
        [HttpGet]
        public async Task<IActionResult> Details(int id)
        {
            var emp = await _db.Employees
                .Include(e => e.DinasList)
                .Include(e => e.CutiList)
                .Include(e => e.SertifikasiList)
                .FirstOrDefaultAsync(e => e.Id == id);

            if (emp == null) return NotFound();

            var vm = new EmployeeDetailVm
            {
                Employee = emp,
                TotalDinas = emp.DinasList?.Count ?? 0,
                TotalCuti = emp.CutiList?.Count ?? 0,
                TotalSertifikasi = emp.SertifikasiList?.Count ?? 0
            };

            return View(vm);
        }

        // HALAMAN KHUSUS: LENGKAPI DATA SETELAH PERTAMA KALI LOGIN
        [Authorize]
        [HttpGet]
        public async Task<IActionResult> FirstLoginEdit()
        {
            var appUser = await _userManager.GetUserAsync(User);
            if (appUser == null)
            {
                return RedirectToAction("Login", "Account");
            }

            if (!(appUser.EmployeeId is int empId) || empId <= 0)
            {
                TempData["ProfileMessage"] = "Akun Anda belum terhubung dengan data pekerja.";
                return RedirectToAction("Index");
            }

            var emp = await _db.Employees.FindAsync(empId);
            if (emp == null)
            {
                TempData["ProfileMessage"] = "Data pekerja Anda tidak ditemukan.";
                return RedirectToAction("Index");
            }

            await LoadLookupToViewBagAsync();
            return View("FirstLoginEdit", emp);
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> FirstLoginEdit(Employee model)
        {
            var appUser = await _userManager.GetUserAsync(User);
            if (appUser == null)
            {
                return RedirectToAction("Login", "Account");
            }

            if (!(appUser.EmployeeId is int empId) || empId <= 0 || empId != model.Id)
            {
                // User mencoba edit data pekerja lain
                return Forbid();
            }

            if (!ModelState.IsValid)
            {
                await LoadLookupToViewBagAsync();
                return View("FirstLoginEdit", model);
            }

            var emp = await _db.Employees.FindAsync(empId);
            if (emp == null)
            {
                TempData["ProfileMessage"] = "Data pekerja Anda tidak ditemukan.";
                return RedirectToAction("Index");
            }

            // === UPDATE FIELD SAMA SEPERTI ACTION Edit ===
            emp.NamaLengkap = model.NamaLengkap;
            emp.NopegPersero = model.NopegPersero;
            emp.NopegKPI = model.NopegKPI;
            emp.Jabatan = model.Jabatan;
            emp.TempatLahir = model.TempatLahir;
            emp.TanggalLahir = model.TanggalLahir;
            emp.JenisKelamin = model.JenisKelamin;
            emp.AlamatDomisili = model.AlamatDomisili;

            emp.StatusPernikahan = model.StatusPernikahan;
            emp.NamaSuamiIstri = model.NamaSuamiIstri;
            emp.JumlahAnak = model.JumlahAnak;

            emp.TanggalMPPK = model.TanggalMPPK;
            emp.TanggalPensiun = model.TanggalPensiun;
            emp.IntakePertama = model.IntakePertama;
            emp.IntakeAktif = model.IntakeAktif;
            emp.ClusteringIntake = model.ClusteringIntake;

            emp.PendidikanTerakhir = model.PendidikanTerakhir;
            emp.Jurusan = model.Jurusan;
            emp.Institusi = model.Institusi;

            emp.UkWearpack = model.UkWearpack;
            emp.UkKemejaPutih = model.UkKemejaPutih;
            emp.UkKemejaProduk = model.UkKemejaProduk;
            emp.UkSepatuSafety = model.UkSepatuSafety;

            emp.EmailPertamina = model.EmailPertamina;
            emp.NomorHP = model.NomorHP;
            emp.NomorTeleponRumah = model.NomorTeleponRumah;
            emp.CutiLimit = model.CutiLimit;

            await _db.SaveChangesAsync();
            await UpsertLookupsFromEmployee(emp);
            await CreateUserFromEmployeeIfNeededAsync(emp);
            await SyncUserFromEmployeeAsync(emp);

            TempData["ProfileMessage"] = "Data pekerja Anda berhasil diperbarui.";
            return RedirectToAction("Index", "Home"); // setelah isi data, masuk dashboard
        }

        // ========== LOOKUP HELPER ==========
        private async Task LoadLookupToViewBagAsync()
        {
            ViewBag.JabatanList = new List<string>();
            ViewBag.PendidikanList = new List<string>();
            ViewBag.JurusanList = new List<string>();
            ViewBag.InstitusiList = new List<string>();
            ViewBag.StatusNikahList = new List<string>();

            ViewBag.UkWearpackList = new List<string>();
            ViewBag.UkKemejaPutihList = new List<string>();
            ViewBag.UkKemejaProdukList = new List<string>();
            ViewBag.UkSepatuSafetyList = new List<string>();

            try
            {
                var lookups = await _db.EmployeeLookups.AsNoTracking().ToListAsync();

                ViewBag.JabatanList = lookups
                    .Where(x => x.Category == "Jabatan")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.PendidikanList = lookups
                    .Where(x => x.Category == "Pendidikan")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.JurusanList = lookups
                    .Where(x => x.Category == "Jurusan")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.InstitusiList = lookups
                    .Where(x => x.Category == "Institusi")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.StatusNikahList = lookups
                    .Where(x => x.Category == "StatusNikah")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.UkWearpackList = lookups
                    .Where(x => x.Category == "UkWearpack")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.UkKemejaPutihList = lookups
                    .Where(x => x.Category == "UkKemejaPutih")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.UkKemejaProdukList = lookups
                    .Where(x => x.Category == "UkKemejaProduk")
                    .Select(x => x.Value).OrderBy(x => x).ToList();

                ViewBag.UkSepatuSafetyList = lookups
                    .Where(x => x.Category == "UkSepatuSafety")
                    .Select(x => x.Value).OrderBy(x => x).ToList();
            }
            catch (SqlException ex)
            {
                _logger.LogWarning(ex, "Tabel EmployeeLookups belum tersedia, lookup di-skip.");
            }

            if (!(ViewBag.UkWearpackList as List<string>)!.Any())
                ViewBag.UkWearpackList = new List<string> { "S", "M", "L", "XL", "XXL" };

            if (!(ViewBag.UkKemejaPutihList as List<string>)!.Any())
                ViewBag.UkKemejaPutihList = new List<string> { "S", "M", "L", "XL", "XXL" };

            if (!(ViewBag.UkKemejaProdukList as List<string>)!.Any())
                ViewBag.UkKemejaProdukList = new List<string> { "S", "M", "L", "XL", "XXL" };

            if (!(ViewBag.UkSepatuSafetyList as List<string>)!.Any())
                ViewBag.UkSepatuSafetyList = new List<string> { "38", "39", "40", "41", "42", "43", "44" };
        }

        // ========== FORM CREATE / EDIT ==========

        [HttpGet]
        public async Task<IActionResult> Create()
        {
            await LoadLookupToViewBagAsync();
            return PartialView("_Form", new Employee());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Employee model)
        {
            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Validasi gagal. Periksa kembali input.",
                    errors
                });
            }

            try
            {
                _db.Employees.Add(model);
                await _db.SaveChangesAsync();
                await UpsertLookupsFromEmployee(model);

                await CreateUserFromEmployeeIfNeededAsync(model);

                return Json(new { success = true, message = "Data pekerja berhasil disimpan dan akun login (jika email tersedia) sudah dibuat." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal menyimpan pekerja (Create)");
                return Json(new { success = false, message = "Gagal menyimpan data: " + ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            var emp = await _db.Employees.FindAsync(id);
            if (emp == null) return NotFound();

            // batas edit: user biasa hanya boleh edit dirinya sendiri
            if (!IsAdminLike())
            {
                var appUser = await _userManager.GetUserAsync(User);
                if (appUser?.EmployeeId != id)
                {
                    return Forbid();
                }
            }

            await LoadLookupToViewBagAsync();
            return PartialView("_Form", emp);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Employee model)
        {
            if (id != model.Id)
                return Json(new { success = false, message = "ID tidak cocok." });

            // batas edit di sisi server
            if (!IsAdminLike())
            {
                var appUser = await _userManager.GetUserAsync(User);
                if (appUser?.EmployeeId != id)
                {
                    return Json(new { success = false, message = "Anda tidak berhak mengubah data pekerja ini." });
                }
            }

            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Validasi gagal. Periksa kembali input.",
                    errors
                });
            }

            var emp = await _db.Employees.FindAsync(id);
            if (emp == null)
                return Json(new { success = false, message = "Data pekerja tidak ditemukan." });

            try
            {
                emp.NamaLengkap = model.NamaLengkap;
                emp.NopegPersero = model.NopegPersero;
                emp.NopegKPI = model.NopegKPI;
                emp.Jabatan = model.Jabatan;
                emp.TempatLahir = model.TempatLahir;
                emp.TanggalLahir = model.TanggalLahir;
                emp.JenisKelamin = model.JenisKelamin;
                emp.AlamatDomisili = model.AlamatDomisili;
                emp.StatusPernikahan = model.StatusPernikahan;
                emp.NamaSuamiIstri = model.NamaSuamiIstri;
                emp.JumlahAnak = model.JumlahAnak;

                emp.TanggalMPPK = model.TanggalMPPK;
                emp.TanggalPensiun = model.TanggalPensiun;
                emp.IntakePertama = model.IntakePertama;
                emp.IntakeAktif = model.IntakeAktif;
                emp.ClusteringIntake = model.ClusteringIntake;

                emp.PendidikanTerakhir = model.PendidikanTerakhir;
                emp.Jurusan = model.Jurusan;
                emp.Institusi = model.Institusi;

                emp.UkWearpack = model.UkWearpack;
                emp.UkKemejaPutih = model.UkKemejaPutih;
                emp.UkKemejaProduk = model.UkKemejaProduk;
                emp.UkSepatuSafety = model.UkSepatuSafety;

                emp.EmailPertamina = model.EmailPertamina;
                emp.NomorHP = model.NomorHP;
                emp.NomorTeleponRumah = model.NomorTeleponRumah;
                emp.CutiLimit = model.CutiLimit;

                await _db.SaveChangesAsync();
                await UpsertLookupsFromEmployee(emp);

                await CreateUserFromEmployeeIfNeededAsync(emp);
                await SyncUserFromEmployeeAsync(emp);

                return Json(new { success = true, message = "Data pekerja & akun login berhasil diperbarui." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal update pekerja (Edit)");
                return Json(new { success = false, message = "Gagal menyimpan data: " + ex.Message });
            }
        }

        // ========== DELETE ==========
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var emp = await _db.Employees.FindAsync(id);
            if (emp == null)
            {
                return Json(new
                {
                    success = false,
                    message = "Data pekerja tidak ditemukan atau sudah dihapus."
                });
            }

            var nama = emp.NamaLengkap ?? "(tanpa nama)";
            var nopeg = emp.NopegPersero ?? "-";

            var user = await _userManager.Users
                .FirstOrDefaultAsync(u => u.EmployeeId == id);

            using var trx = await _db.Database.BeginTransactionAsync();
            try
            {
                if (user != null)
                {
                    var resultUser = await _userManager.DeleteAsync(user);
                    if (!resultUser.Succeeded)
                    {
                        var msg = string.Join("; ", resultUser.Errors.Select(e => e.Description));
                        await trx.RollbackAsync();
                        return Json(new
                        {
                            success = false,
                            message = "Gagal menghapus akun login: " + msg
                        });
                    }
                }

                _db.Employees.Remove(emp);
                await _db.SaveChangesAsync();

                await trx.CommitAsync();

                var extra = user != null
                    ? " Akun login yang terhubung juga ikut dihapus."
                    : string.Empty;

                return Json(new
                {
                    success = true,
                    message = $"Data pekerja \"{nama}\" ({nopeg}) berhasil dihapus.{extra}"
                });
            }
            catch (Exception ex)
            {
                await trx.RollbackAsync();
                _logger.LogError(ex, "Gagal menghapus pekerja Id={Id}", id);

                return Json(new
                {
                    success = false,
                    message = "Gagal menghapus data pekerja. Silakan coba lagi atau hubungi administrator."
                });
            }
        }

        [HttpGet]
        public IActionResult ImportDialog()
        {
            return PartialView("ImportDialog");
        }

        // ========== IMPORT EXCEL ==========

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile? file)
        {
            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            int imported = 0;
            int row = 3;

            var newEmployees = new List<Employee>();
            var updatedEmployeeIds = new HashSet<int>();

            try
            {
                using var ms = new MemoryStream();
                await file.CopyToAsync(ms);
                ms.Position = 0;

                using var package = new ExcelPackage(ms);
                var ws = package.Workbook.Worksheets[0];
                if (ws == null)
                    return Json(new { success = false, message = "Worksheet tidak ditemukan." });

                var title = ws.Cells[1, 1].GetValue<string>()?.Trim() ?? string.Empty;
                var hNama = ws.Cells[2, 2].GetValue<string>()?.Trim() ?? string.Empty;
                var hNopeg = ws.Cells[2, 3].GetValue<string>()?.Trim() ?? string.Empty;

                if (!title.Contains("TEMPLATE DATA KARYAWAN", StringComparison.OrdinalIgnoreCase) ||
                    !hNama.Contains("NamaLengkap", StringComparison.OrdinalIgnoreCase) ||
                    !hNopeg.Contains("NopegPersero", StringComparison.OrdinalIgnoreCase))
                {
                    return Json(new
                    {
                        success = false,
                        message = "Template tidak dikenali sebagai template karyawan. " +
                                  "Pastikan Anda menggunakan file template karyawan UtilitiesHR, " +
                                  "bukan template file lain."
                    });
                }

                DateTime? ParseDate(object? v)
                {
                    if (v == null) return null;

                    if (v is DateTime dt) return dt;

                    if (v is double d)
                    {
                        return DateTime.FromOADate(d);
                    }

                    if (v is string s)
                    {
                        s = s.Trim();
                        if (string.IsNullOrWhiteSpace(s)) return null;

                        if (DateTime.TryParse(s, out var d1))
                            return d1;

                        if (double.TryParse(s, out var d2) &&
                            d2 > 20000 && d2 < 60000)
                        {
                            return DateTime.FromOADate(d2);
                        }
                    }

                    return null;
                }

                int? ParseInt(string? s)
                {
                    if (string.IsNullOrWhiteSpace(s)) return null;
                    if (int.TryParse(s, out var v)) return v;
                    return null;
                }

                while (true)
                {
                    var nama = ws.Cells[row, 2].GetValue<string>()?.Trim();
                    var nopeg = ws.Cells[row, 3].GetValue<string>()?.Trim();

                    if (string.IsNullOrWhiteSpace(nama) && string.IsNullOrWhiteSpace(nopeg))
                        break;

                    if (string.IsNullOrWhiteSpace(nama) || string.IsNullOrWhiteSpace(nopeg))
                        throw new Exception($"Baris {row}: Nama dan Nopeg wajib diisi.");

                    string? nopegKpi = ws.Cells[row, 4].GetValue<string>()?.Trim();
                    string? jabatan = ws.Cells[row, 5].GetValue<string>()?.Trim();
                    string? jenisKelamin = ws.Cells[row, 6].GetValue<string>()?.Trim();
                    string? tempatLahir = ws.Cells[row, 7].GetValue<string>()?.Trim();
                    object? tglLahirVal = ws.Cells[row, 8].Value;
                    string? statusNikah = ws.Cells[row, 9].GetValue<string>()?.Trim();
                    string? namaSuamiIstri = ws.Cells[row, 10].GetValue<string>()?.Trim();
                    string? jmlAnakStr = ws.Cells[row, 11].GetValue<string>()?.Trim();
                    object? tglMppkVal = ws.Cells[row, 12].Value;
                    object? tglPensiunVal = ws.Cells[row, 13].Value;
                    string? intakePertamaStr = ws.Cells[row, 14].Text?.Trim();
                    object? intakeAktifVal = ws.Cells[row, 15].Value;
                    string? clustering = ws.Cells[row, 16].GetValue<string>()?.Trim();
                    string? pendidikan = ws.Cells[row, 17].GetValue<string>()?.Trim();
                    string? jurusan = ws.Cells[row, 18].GetValue<string>()?.Trim();
                    string? institusi = ws.Cells[row, 19].GetValue<string>()?.Trim();
                    string? ukWearpack = ws.Cells[row, 20].GetValue<string>()?.Trim();
                    string? ukKemejaPutih = ws.Cells[row, 21].GetValue<string>()?.Trim();
                    string? ukKemejaProduk = ws.Cells[row, 22].GetValue<string>()?.Trim();
                    string? ukSepatuSafety = ws.Cells[row, 23].GetValue<string>()?.Trim();
                    string? email = ws.Cells[row, 24].GetValue<string>()?.Trim();
                    string? phone = ws.Cells[row, 25].GetValue<string>()?.Trim();
                    string? telRumah = ws.Cells[row, 26].GetValue<string>()?.Trim();
                    string? cutiLimitStr = ws.Cells[row, 27].GetValue<string>()?.Trim();
                    string? alamatDomisili = ws.Cells[row, 28].GetValue<string>()?.Trim();

                    var existing = await _db.Employees
                        .FirstOrDefaultAsync(e => e.NopegPersero == nopeg);

                    if (existing == null)
                    {
                        var emp = new Employee
                        {
                            NamaLengkap = nama,
                            NopegPersero = nopeg,
                            NopegKPI = nopegKpi,
                            Jabatan = jabatan,
                            JenisKelamin = jenisKelamin,
                            TempatLahir = tempatLahir,
                            TanggalLahir = ParseDate(tglLahirVal),
                            StatusPernikahan = statusNikah,
                            NamaSuamiIstri = namaSuamiIstri,
                            JumlahAnak = ParseInt(jmlAnakStr),
                            TanggalMPPK = ParseDate(tglMppkVal),
                            TanggalPensiun = ParseDate(tglPensiunVal),
                            IntakePertama = intakePertamaStr,
                            IntakeAktif = ParseDate(intakeAktifVal),
                            ClusteringIntake = clustering,
                            PendidikanTerakhir = pendidikan,
                            Jurusan = jurusan,
                            Institusi = institusi,
                            UkWearpack = ukWearpack,
                            UkKemejaPutih = ukKemejaPutih,
                            UkKemejaProduk = ukKemejaProduk,
                            UkSepatuSafety = ukSepatuSafety,
                            EmailPertamina = email,
                            NomorHP = phone,
                            NomorTeleponRumah = telRumah,
                            CutiLimit = ParseInt(cutiLimitStr),
                            AlamatDomisili = alamatDomisili
                        };

                        _db.Employees.Add(emp);
                        newEmployees.Add(emp);
                    }
                    else
                    {
                        existing.NamaLengkap = nama;
                        existing.NopegKPI = nopegKpi ?? existing.NopegKPI;
                        existing.Jabatan = jabatan ?? existing.Jabatan;
                        existing.JenisKelamin = jenisKelamin ?? existing.JenisKelamin;
                        existing.TempatLahir = tempatLahir ?? existing.TempatLahir;
                        existing.TanggalLahir = ParseDate(tglLahirVal) ?? existing.TanggalLahir;
                        existing.StatusPernikahan = statusNikah ?? existing.StatusPernikahan;
                        existing.NamaSuamiIstri = namaSuamiIstri ?? existing.NamaSuamiIstri;
                        existing.JumlahAnak = ParseInt(jmlAnakStr) ?? existing.JumlahAnak;
                        existing.TanggalMPPK = ParseDate(tglMppkVal) ?? existing.TanggalMPPK;
                        existing.TanggalPensiun = ParseDate(tglPensiunVal) ?? existing.TanggalPensiun;
                        if (!string.IsNullOrWhiteSpace(intakePertamaStr))
                            existing.IntakePertama = intakePertamaStr;
                        existing.IntakeAktif = ParseDate(intakeAktifVal) ?? existing.IntakeAktif;
                        existing.ClusteringIntake = clustering ?? existing.ClusteringIntake;
                        existing.PendidikanTerakhir = pendidikan ?? existing.PendidikanTerakhir;
                        existing.Jurusan = jurusan ?? existing.Jurusan;
                        existing.Institusi = institusi ?? existing.Institusi;
                        existing.UkWearpack = ukWearpack ?? existing.UkWearpack;
                        existing.UkKemejaPutih = ukKemejaPutih ?? existing.UkKemejaPutih;
                        existing.UkKemejaProduk = ukKemejaProduk ?? existing.UkKemejaProduk;
                        existing.UkSepatuSafety = ukSepatuSafety ?? existing.UkSepatuSafety;
                        existing.EmailPertamina = email ?? existing.EmailPertamina;
                        existing.NomorHP = phone ?? existing.NomorHP;
                        existing.NomorTeleponRumah = telRumah ?? existing.NomorTeleponRumah;
                        existing.CutiLimit = ParseInt(cutiLimitStr) ?? existing.CutiLimit;
                        existing.AlamatDomisili = alamatDomisili ?? existing.AlamatDomisili;

                        if (existing.Id != 0)
                            updatedEmployeeIds.Add(existing.Id);
                    }

                    await UpsertSingleLookup("Jabatan", jabatan);
                    await UpsertSingleLookup("Pendidikan", pendidikan);
                    await UpsertSingleLookup("Jurusan", jurusan);
                    await UpsertSingleLookup("Institusi", institusi);
                    await UpsertSingleLookup("StatusNikah", statusNikah);
                    await UpsertSingleLookup("UkWearpack", ukWearpack);
                    await UpsertSingleLookup("UkKemejaPutih", ukKemejaPutih);
                    await UpsertSingleLookup("UkKemejaProduk", ukKemejaProduk);
                    await UpsertSingleLookup("UkSepatuSafety", ukSepatuSafety);

                    imported++;
                    row++;
                }

                await _db.SaveChangesAsync();

                foreach (var emp in newEmployees)
                {
                    await CreateUserFromEmployeeIfNeededAsync(emp);
                    await SyncUserFromEmployeeAsync(emp);
                }

                foreach (var empId in updatedEmployeeIds)
                {
                    var emp = await _db.Employees.FindAsync(empId);
                    if (emp != null)
                    {
                        await CreateUserFromEmployeeIfNeededAsync(emp);
                        await SyncUserFromEmployeeAsync(emp);
                    }
                }

                return Json(new
                {
                    success = true,
                    imported,
                    message = $"Berhasil mengimpor {imported} baris."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Import karyawan gagal pada baris {Row}", row);
                return Json(new { success = false, message = $"Import gagal pada baris {row}: {ex.Message}" });
            }
        }

        // ========== TEMPLATE IMPORT EXCEL ==========

        [HttpGet]
        public IActionResult TemplateImport()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Template");

            ws.Cells[1, 1].Value = "TEMPLATE DATA KARYAWAN UTILITIES";
            ws.Cells[1, 1, 1, 28].Merge = true;
            ws.Cells[1, 1, 1, 28].Style.Font.Bold = true;

            int c = 1;
            ws.Cells[2, c++].Value = "No";
            ws.Cells[2, c++].Value = "NamaLengkap*";
            ws.Cells[2, c++].Value = "NopegPersero*";
            ws.Cells[2, c++].Value = "NopegKPI";
            ws.Cells[2, c++].Value = "Jabatan";
            ws.Cells[2, c++].Value = "JenisKelamin";
            ws.Cells[2, c++].Value = "TempatLahir";
            ws.Cells[2, c++].Value = "TanggalLahir";
            ws.Cells[2, c++].Value = "StatusPernikahan";
            ws.Cells[2, c++].Value = "NamaSuamiIstri";
            ws.Cells[2, c++].Value = "JumlahAnak";
            ws.Cells[2, c++].Value = "TanggalMPPK";
            ws.Cells[2, c++].Value = "TanggalPensiun";
            ws.Cells[2, c++].Value = "IntakePertama";
            ws.Cells[2, c++].Value = "IntakeAktif";
            ws.Cells[2, c++].Value = "ClusteringIntake";
            ws.Cells[2, c++].Value = "PendidikanTerakhir";
            ws.Cells[2, c++].Value = "Jurusan";
            ws.Cells[2, c++].Value = "Institusi";
            ws.Cells[2, c++].Value = "UkWearpack";
            ws.Cells[2, c++].Value = "UkKemejaPutih";
            ws.Cells[2, c++].Value = "UkKemejaProduk";
            ws.Cells[2, c++].Value = "UkSepatuSafety";
            ws.Cells[2, c++].Value = "EmailPertamina";
            ws.Cells[2, c++].Value = "NomorHP";
            ws.Cells[2, c++].Value = "NomorTeleponRumah";
            ws.Cells[2, c++].Value = "CutiLimit";
            ws.Cells[2, c++].Value = "AlamatDomisili";

            ws.Cells[2, 1, 2, 28].Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            var fileName = $"template-data-karyawan-{DateTime.Now:yyyyMMdd}.xlsx";
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ========== EXPORT CSV ==========

        [HttpGet]
        public async Task<IActionResult> ExportCsv()
        {
            var data = await _db.Employees.AsNoTracking()
                .OrderBy(e => e.NamaLengkap)
                .ToListAsync();

            var sb = new StringBuilder();
            sb.AppendLine("No;NamaLengkap;NopegPersero;NopegKPI;Jabatan;JenisKelamin;TempatLahir;TanggalLahir;StatusPernikahan;NamaSuamiIstri;JumlahAnak;TanggalMPPK;TanggalPensiun;IntakePertama;IntakeAktif;ClusteringIntake;PendidikanTerakhir;Jurusan;Institusi;UkWearpack;UkKemejaPutih;UkKemejaProduk;UkSepatuSafety;EmailPertamina;NomorHP;NomorTeleponRumah;CutiLimit;AlamatDomisili");

            int no = 1;
            foreach (var e in data)
            {
                string[] cols =
                {
                    no.ToString(),
                    e.NamaLengkap,
                    e.NopegPersero,
                    e.NopegKPI ?? "",
                    e.Jabatan ?? "",
                    e.JenisKelamin ?? "",
                    e.TempatLahir ?? "",
                    e.TanggalLahir?.ToString("yyyy-MM-dd") ?? "",
                    e.StatusPernikahan ?? "",
                    e.NamaSuamiIstri ?? "",
                    e.JumlahAnak?.ToString() ?? "",
                    e.TanggalMPPK?.ToString("yyyy-MM-dd") ?? "",
                    e.TanggalPensiun?.ToString("yyyy-MM-dd") ?? "",
                    e.IntakePertama ?? "",
                    e.IntakeAktif?.ToString("yyyy-MM-dd") ?? "",
                    e.ClusteringIntake ?? "",
                    e.PendidikanTerakhir ?? "",
                    e.Jurusan ?? "",
                    e.Institusi ?? "",
                    e.UkWearpack ?? "",
                    e.UkKemejaPutih ?? "",
                    e.UkKemejaProduk ?? "",
                    e.UkSepatuSafety ?? "",
                    e.EmailPertamina ?? "",
                    e.NomorHP ?? "",
                    e.NomorTeleponRumah ?? "",
                    e.CutiLimit?.ToString() ?? "",
                    e.AlamatDomisili ?? ""
                };

                sb.AppendLine(string.Join(";", cols));
                no++;
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var fileName = $"data-karyawan-{DateTime.Now:yyyyMMdd}.csv";
            return File(bytes, "text/csv", fileName);
        }

        // ========== EXPORT EXCEL ==========

        [HttpGet]
        public async Task<IActionResult> ExportExcel()
        {
            var data = await _db.Employees.AsNoTracking()
                .OrderBy(e => e.NamaLengkap)
                .ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Karyawan");

            ws.Cells[1, 1].Value = "DATA KARYAWAN UTILITIES";
            ws.Cells[1, 1, 1, 28].Merge = true;
            ws.Cells[1, 1, 1, 28].Style.Font.Bold = true;

            int c = 1;
            ws.Cells[2, c++].Value = "No";
            ws.Cells[2, c++].Value = "NamaLengkap";
            ws.Cells[2, c++].Value = "NopegPersero";
            ws.Cells[2, c++].Value = "NopegKPI";
            ws.Cells[2, c++].Value = "Jabatan";
            ws.Cells[2, c++].Value = "JenisKelamin";
            ws.Cells[2, c++].Value = "TempatLahir";
            ws.Cells[2, c++].Value = "TanggalLahir";
            ws.Cells[2, c++].Value = "StatusPernikahan";
            ws.Cells[2, c++].Value = "NamaSuamiIstri";
            ws.Cells[2, c++].Value = "JumlahAnak";
            ws.Cells[2, c++].Value = "TanggalMPPK";
            ws.Cells[2, c++].Value = "TanggalPensiun";
            ws.Cells[2, c++].Value = "IntakePertama";
            ws.Cells[2, c++].Value = "IntakeAktif";
            ws.Cells[2, c++].Value = "ClusteringIntake";
            ws.Cells[2, c++].Value = "PendidikanTerakhir";
            ws.Cells[2, c++].Value = "Jurusan";
            ws.Cells[2, c++].Value = "Institusi";
            ws.Cells[2, c++].Value = "UkWearpack";
            ws.Cells[2, c++].Value = "UkKemejaPutih";
            ws.Cells[2, c++].Value = "UkKemejaProduk";
            ws.Cells[2, c++].Value = "UkSepatuSafety";
            ws.Cells[2, c++].Value = "EmailPertamina";
            ws.Cells[2, c++].Value = "NomorHP";
            ws.Cells[2, c++].Value = "NomorTeleponRumah";
            ws.Cells[2, c++].Value = "CutiLimit";
            ws.Cells[2, c++].Value = "AlamatDomisili";

            int row = 3;
            int no = 1;

            foreach (var e in data)
            {
                c = 1;
                ws.Cells[row, c++].Value = no;
                ws.Cells[row, c++].Value = e.NamaLengkap;
                ws.Cells[row, c++].Value = e.NopegPersero;
                ws.Cells[row, c++].Value = e.NopegKPI;
                ws.Cells[row, c++].Value = e.Jabatan;
                ws.Cells[row, c++].Value = e.JenisKelamin;
                ws.Cells[row, c++].Value = e.TempatLahir;
                ws.Cells[row, c++].Value = e.TanggalLahir?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = e.StatusPernikahan;
                ws.Cells[row, c++].Value = e.NamaSuamiIstri;
                ws.Cells[row, c++].Value = e.JumlahAnak;
                ws.Cells[row, c++].Value = e.TanggalMPPK?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = e.TanggalPensiun?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = e.IntakePertama;
                ws.Cells[row, c++].Value = e.IntakeAktif?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = e.ClusteringIntake;
                ws.Cells[row, c++].Value = e.PendidikanTerakhir;
                ws.Cells[row, c++].Value = e.Jurusan;
                ws.Cells[row, c++].Value = e.Institusi;
                ws.Cells[row, c++].Value = e.UkWearpack;
                ws.Cells[row, c++].Value = e.UkKemejaPutih;
                ws.Cells[row, c++].Value = e.UkKemejaProduk;
                ws.Cells[row, c++].Value = e.UkSepatuSafety;
                ws.Cells[row, c++].Value = e.EmailPertamina;
                ws.Cells[row, c++].Value = e.NomorHP;
                ws.Cells[row, c++].Value = e.NomorTeleponRumah;
                ws.Cells[row, c++].Value = e.CutiLimit;
                ws.Cells[row, c++].Value = e.AlamatDomisili;

                row++;
                no++;
            }

            ws.Cells[2, 1, 2, 28].Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            var fileName = $"data-karyawan-{DateTime.Now:yyyyMMdd}.xlsx";
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ========== HELPER LOOKUP ==========

        private async Task UpsertLookupsFromEmployee(Employee e)
        {
            await UpsertSingleLookup("Jabatan", e.Jabatan);
            await UpsertSingleLookup("Pendidikan", e.PendidikanTerakhir);
            await UpsertSingleLookup("Jurusan", e.Jurusan);
            await UpsertSingleLookup("Institusi", e.Institusi);
            await UpsertSingleLookup("StatusNikah", e.StatusPernikahan);
            await UpsertSingleLookup("UkWearpack", e.UkWearpack);
            await UpsertSingleLookup("UkKemejaPutih", e.UkKemejaPutih);
            await UpsertSingleLookup("UkKemejaProduk", e.UkKemejaProduk);
            await UpsertSingleLookup("UkSepatuSafety", e.UkSepatuSafety);
        }

        private async Task UpsertSingleLookup(string category, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            value = value.Trim();

            try
            {
                var exists = await _db.EmployeeLookups
                    .AnyAsync(x => x.Category == category && x.Value == value);

                if (!exists)
                {
                    _db.EmployeeLookups.Add(new EmployeeLookup
                    {
                        Category = category,
                        Value = value
                    });
                    await _db.SaveChangesAsync();
                }
            }
            catch (SqlException ex)
            {
                _logger.LogWarning(ex, "Tabel EmployeeLookups belum tersedia, upsert di-skip.");
            }
        }
    }
}
