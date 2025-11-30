using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class SertifikasisController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;

        static SertifikasisController()
        {
            ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR - PT KPI RU II");
        }

        public SertifikasisController(ApplicationDbContext db, UserManager<ApplicationUser> userManager)
        {
            _db = db;
            _userManager = userManager;
        }

        private bool IsAdminLike =>
            User.IsInRole("Admin") || User.IsInRole("SuperAdmin") || User.IsInRole("Supervisor");

        private bool IsUserRole => User.IsInRole("User");

        // ================== Helper: set current employee untuk user biasa ==================
        private async Task<(int? employeeId, string employeeName)> GetCurrentEmployeeInfo()
        {
            if (!IsAdminLike && IsUserRole)
            {
                var appUser = await _userManager.GetUserAsync(User);
                if (appUser?.EmployeeId is int empId)
                {
                    var emp = await _db.Employees
                        .AsNoTracking()
                        .FirstOrDefaultAsync(e => e.Id == empId);

                    return (empId, emp?.NamaLengkap ?? "");
                }
            }
            return (null, "");
        }

        private async Task SetCurrentEmployeeViewBags()
        {
            var (employeeId, employeeName) = await GetCurrentEmployeeInfo();
            if (employeeId.HasValue)
            {
                ViewBag.CurrentEmployeeId = employeeId.Value;
                ViewBag.CurrentEmployeeName = employeeName;
            }
        }

        // =========================== INDEX ===========================
        public async Task<IActionResult> Index(int? employeeId, string? q)
        {
            await SetCurrentEmployeeViewBags();

            var (currentEmpId, currentEmpName) = await GetCurrentEmployeeInfo();

            // Kalau role User biasa -> cuma boleh lihat data sendiri
            if (!IsAdminLike && IsUserRole)
            {
                if (currentEmpId.HasValue)
                {
                    employeeId = currentEmpId.Value; // paksa filter
                }
                else
                {
                    ViewBag.Employees = new SelectList(Enumerable.Empty<object>());
                    return View(new List<Sertifikasi>());
                }
            }

            // Dropdown pegawai
            if (IsAdminLike)
            {
                ViewBag.Employees = new SelectList(
                    await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                    "Id", "NamaLengkap", employeeId
                );
            }
            else if (IsUserRole && currentEmpId.HasValue)
            {
                var listSelf = await _db.Employees
                    .Where(e => e.Id == currentEmpId.Value)
                    .OrderBy(e => e.NamaLengkap)
                    .ToListAsync();

                ViewBag.Employees = new SelectList(listSelf, "Id", "NamaLengkap", employeeId);
            }
            else
            {
                ViewBag.Employees = new SelectList(Enumerable.Empty<object>());
            }

            var query = _db.Sertifikasis
                .Include(s => s.Employee)
                .Include(s => s.MasterSertifikasi)
                .AsQueryable();

            if (employeeId is not null)
                query = query.Where(s => s.EmployeeId == employeeId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(s =>
                    (s.Employee!.NamaLengkap ?? "").Contains(q) ||
                    (s.Employee!.NopegPersero ?? "").Contains(q) ||
                    (s.NamaSertifikasi ?? "").Contains(q));
            }

            var data = await query
                .OrderByDescending(s => s.BerlakuSampai)
                .ThenByDescending(s => s.TanggalMulai)
                .ToListAsync();

            return View(data);
        }

        // ======================= DETAILS POPUP ==========================
        public async Task<IActionResult> Details(int id)
        {
            await SetCurrentEmployeeViewBags();

            var s = await _db.Sertifikasis
                .Include(x => x.Employee)
                .Include(x => x.MasterSertifikasi)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (s == null) return NotFound();

            // User biasa hanya boleh lihat miliknya sendiri
            if (!IsAdminLike && IsUserRole)
            {
                var (currentEmpId, _) = await GetCurrentEmployeeInfo();
                if (!currentEmpId.HasValue || currentEmpId.Value != s.EmployeeId)
                    return Forbid();
            }

            return PartialView("Details", s);
        }

        // ======================== DROPDOWNS ==========================
        private async Task LoadDropdowns(object? selectedEmployeeId = null, string? selectedNamaSert = null)
        {
            // Pegawai (untuk admin)
            if (IsAdminLike)
            {
                ViewBag.EmployeeId = new SelectList(
                    await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                    "Id", "NamaLengkap", selectedEmployeeId
                );
            }

            // Master sertifikasi → dipakai sebagai opsi nama (value = Nama, text = Nama)
            var masters = await _db.MasterSertifikasis
                .OrderBy(x => x.Nama)
                .ToListAsync();

            var items = masters
                .Select(m => new SelectListItem
                {
                    Value = m.Nama ?? "",
                    Text = m.Nama ?? "",
                    Selected = (!string.IsNullOrWhiteSpace(selectedNamaSert)
                                && string.Equals(m.Nama?.Trim(), selectedNamaSert.Trim(), StringComparison.OrdinalIgnoreCase))
                })
                .ToList();

            // === PENTING: kalau nilai lama (NamaSertifikasi) tidak ada di master,
            // tetap masukkan ke dropdown supaya muncul di Edit ===
            if (!string.IsNullOrWhiteSpace(selectedNamaSert)
                && !items.Any(i => string.Equals(i.Value.Trim(), selectedNamaSert.Trim(), StringComparison.OrdinalIgnoreCase)))
            {
                items.Add(new SelectListItem
                {
                    Value = selectedNamaSert.Trim(),
                    Text = selectedNamaSert.Trim(),
                    Selected = true
                });
            }

            ViewBag.SertifikasiOptions = items;
        }

        // ========================== CREATE ===========================
        public async Task<IActionResult> Create(int? employeeId)
        {
            await SetCurrentEmployeeViewBags();

            var (currentEmpId, currentEmpName) = await GetCurrentEmployeeInfo();

            // Untuk user biasa: paksa employeeId = dirinya sendiri
            if (!IsAdminLike && IsUserRole)
            {
                if (currentEmpId.HasValue)
                {
                    employeeId = currentEmpId.Value;
                }
                else
                {
                    // Jika user biasa tapi tidak punya employeeId, tampilkan error
                    return Json(new { success = false, message = "Akun Anda tidak terdaftar sebagai karyawan. Silakan hubungi administrator." });
                }
            }

            await LoadDropdowns(employeeId, null);

            var model = new Sertifikasi
            {
                EmployeeId = employeeId ?? 0
            };

            return PartialView("Create", model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Sertifikasi model)
        {
            // Dapatkan info employee current user
            var (currentEmpId, currentEmpName) = await GetCurrentEmployeeInfo();

            // Untuk user biasa: override EmployeeId dengan ID mereka sendiri
            if (!IsAdminLike && IsUserRole)
            {
                if (!currentEmpId.HasValue)
                {
                    return Json(new { success = false, message = "Akun Anda tidak terdaftar sebagai karyawan. Silakan hubungi administrator." });
                }

                // Override EmployeeId dengan ID user yang login
                model.EmployeeId = currentEmpId.Value;
            }

            // Validasi EmployeeId untuk user biasa
            if (!IsAdminLike && IsUserRole && model.EmployeeId != currentEmpId)
            {
                return Json(new { success = false, message = "Anda hanya boleh menginput sertifikasi untuk diri sendiri." });
            }

            // Validasi dulu
            if (!ModelState.IsValid)
            {
                await LoadDropdowns(model.EmployeeId, model.NamaSertifikasi);
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
                return Json(new { success = false, message = "Data tidak valid.", errors });
            }

            // ====== MAP NamaSertifikasi → MasterSertifikasi (AUTO CREATE JIKA BARU) ======
            if (!string.IsNullOrWhiteSpace(model.NamaSertifikasi))
            {
                var namaTrim = model.NamaSertifikasi.Trim();

                var master = await _db.MasterSertifikasis
                    .FirstOrDefaultAsync(m => m.Nama == namaTrim);

                if (master == null)
                {
                    master = new MasterSertifikasi
                    {
                        Nama = namaTrim
                    };
                    _db.MasterSertifikasis.Add(master);
                    await _db.SaveChangesAsync();
                }

                model.MasterSertifikasiId = master.Id;
            }
            else
            {
                model.MasterSertifikasiId = null;
            }

            try
            {
                _db.Add(model);
                await _db.SaveChangesAsync();
                return Json(new { success = true, message = "Data sertifikasi berhasil disimpan.", employeeId = model.EmployeeId });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Terjadi kesalahan saat menyimpan data: " + ex.Message });
            }
        }

        // =========================== EDIT ============================
        public async Task<IActionResult> Edit(int id)
        {
            await SetCurrentEmployeeViewBags();

            var s = await _db.Sertifikasis
                .Include(x => x.Employee)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (s == null) return NotFound();

            // User biasa hanya boleh edit miliknya sendiri
            if (!IsAdminLike && IsUserRole)
            {
                var (currentEmpId, _) = await GetCurrentEmployeeInfo();
                if (!currentEmpId.HasValue || currentEmpId.Value != s.EmployeeId)
                    return Forbid();
            }

            // Load dropdown + sertakan nama lama kalau tidak ada di master
            await LoadDropdowns(s.EmployeeId, s.NamaSertifikasi);
            return PartialView("Edit", s);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Sertifikasi model)
        {
            if (id != model.Id) return BadRequest();

            // Dapatkan info employee current user
            var (currentEmpId, currentEmpName) = await GetCurrentEmployeeInfo();

            // Untuk user biasa: pastikan mereka hanya edit data milik sendiri
            if (!IsAdminLike && IsUserRole)
            {
                if (!currentEmpId.HasValue)
                {
                    return Json(new { success = false, message = "Akun Anda tidak terdaftar sebagai karyawan. Silakan hubungi administrator." });
                }

                // Cek apakah data yang akan diupdate memang milik user ini
                var existing = await _db.Sertifikasis
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.Id == id);

                if (existing == null || existing.EmployeeId != currentEmpId.Value)
                {
                    return Json(new { success = false, message = "Anda tidak berhak mengubah data ini." });
                }

                // Override EmployeeId dengan ID user yang login
                model.EmployeeId = currentEmpId.Value;
            }

            // Validasi dulu
            if (!ModelState.IsValid)
            {
                await LoadDropdowns(model.EmployeeId, model.NamaSertifikasi);
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
                return Json(new { success = false, message = "Data tidak valid.", errors });
            }

            // ====== MAP NamaSertifikasi → MasterSertifikasi (AUTO CREATE JIKA BARU) ======
            if (!string.IsNullOrWhiteSpace(model.NamaSertifikasi))
            {
                var namaTrim = model.NamaSertifikasi.Trim();

                var master = await _db.MasterSertifikasis
                    .FirstOrDefaultAsync(m => m.Nama == namaTrim);

                if (master == null)
                {
                    master = new MasterSertifikasi
                    {
                        Nama = namaTrim
                    };
                    _db.MasterSertifikasis.Add(master);
                    await _db.SaveChangesAsync();
                }

                model.MasterSertifikasiId = master.Id;
            }
            else
            {
                model.MasterSertifikasiId = null;
            }

            try
            {
                _db.Update(model);
                await _db.SaveChangesAsync();
                return Json(new { success = true, message = "Data sertifikasi berhasil diupdate.", employeeId = model.EmployeeId });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Terjadi kesalahan saat mengupdate data: " + ex.Message });
            }
        }

        // ========================== DELETE ===========================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var s = await _db.Sertifikasis.FindAsync(id);
            if (s == null)
            {
                return Json(new { success = false, message = "Data tidak ditemukan." });
            }

            // User biasa hanya boleh hapus miliknya sendiri
            if (!IsAdminLike && IsUserRole)
            {
                var (currentEmpId, _) = await GetCurrentEmployeeInfo();
                if (!currentEmpId.HasValue || currentEmpId.Value != s.EmployeeId)
                {
                    return Json(new { success = false, message = "Anda tidak berhak menghapus data ini." });
                }
            }

            try
            {
                _db.Sertifikasis.Remove(s);
                await _db.SaveChangesAsync();
                return Json(new { success = true, message = "Data sertifikasi berhasil dihapus." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Terjadi kesalahan saat menghapus data: " + ex.Message });
            }
        }

        // ================= TEMPLATE & EXPORT =========================
        [HttpGet]
        public IActionResult DownloadTemplateExcel()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("TemplateSertifikasi");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Nopeg";
            ws.Cells[1, 3].Value = "NamaSertifikasi";
            ws.Cells[1, 4].Value = "Lokasi";
            ws.Cells[1, 5].Value = "TanggalMulai (yyyy-MM-dd)";
            ws.Cells[1, 6].Value = "TanggalSelesai (yyyy-MM-dd)";
            ws.Cells[1, 7].Value = "BerlakuDari (yyyy-MM-dd)";
            ws.Cells[1, 8].Value = "BerlakuSampai (yyyy-MM-dd)";
            ws.Cells[1, 9].Value = "Keterangan";

            using (var range = ws.Cells[1, 1, 1, 9])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                    .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            const string fileName = "Template_Import_Sertifikasi.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(bytes, contentType, fileName);
        }

        [HttpGet]
        public IActionResult DownloadTemplateCsv()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Nama,Nopeg,NamaSertifikasi,Lokasi,TanggalMulai (yyyy-MM-dd),TanggalSelesai (yyyy-MM-dd),BerlakuDari (yyyy-MM-dd),BerlakuSampai (yyyy-MM-dd),Keterangan");
            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            const string fileName = "Template_Import_Sertifikasi.csv";
            return File(bytes, "text/csv", fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportExcel(int? employeeId, string? q)
        {
            if (!IsAdminLike && IsUserRole)
            {
                var (currentEmpId, _) = await GetCurrentEmployeeInfo();
                if (currentEmpId.HasValue)
                {
                    employeeId = currentEmpId.Value;
                }
                else
                {
                    return Forbid();
                }
            }

            var list = await BuildFilteredQuery(employeeId, q)
                .OrderBy(s => s.Employee!.NamaLengkap)
                .ThenByDescending(s => s.BerlakuSampai)
                .ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sertifikasi");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Nopeg";
            ws.Cells[1, 3].Value = "NamaSertifikasi";
            ws.Cells[1, 4].Value = "Lokasi";
            ws.Cells[1, 5].Value = "TanggalMulai";
            ws.Cells[1, 6].Value = "TanggalSelesai";
            ws.Cells[1, 7].Value = "BerlakuDari";
            ws.Cells[1, 8].Value = "BerlakuSampai";
            ws.Cells[1, 9].Value = "Keterangan";
            ws.Cells[1, 10].Value = "Status";

            using (var range = ws.Cells[1, 1, 1, 10])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                    .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            int row = 2;
            foreach (var s in list)
            {
                ws.Cells[row, 1].Value = s.Employee?.NamaLengkap;
                ws.Cells[row, 2].Value = s.Employee?.NopegPersero;
                ws.Cells[row, 3].Value = s.NamaSertifikasi;
                ws.Cells[row, 4].Value = s.Lokasi;
                ws.Cells[row, 5].Value = s.TanggalMulai?.ToString("yyyy-MM-dd");
                ws.Cells[row, 6].Value = s.TanggalSelesai?.ToString("yyyy-MM-dd");
                ws.Cells[row, 7].Value = s.BerlakuDari?.ToString("yyyy-MM-dd");
                ws.Cells[row, 8].Value = s.BerlakuSampai?.ToString("yyyy-MM-dd");
                ws.Cells[row, 9].Value = s.Keterangan;
                ws.Cells[row, 10].Value = s.IsExpired ? "Expired" : "Aktif";
                row++;
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            const string fileName = "Data_Sertifikasi.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(bytes, contentType, fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportCsv(int? employeeId, string? q)
        {
            if (!IsAdminLike && IsUserRole)
            {
                var (currentEmpId, _) = await GetCurrentEmployeeInfo();
                if (currentEmpId.HasValue)
                {
                    employeeId = currentEmpId.Value;
                }
                else
                {
                    return Forbid();
                }
            }

            var list = await BuildFilteredQuery(employeeId, q)
                .OrderBy(s => s.Employee!.NamaLengkap)
                .ThenByDescending(s => s.BerlakuSampai)
                .ToListAsync();

            var sb = new StringBuilder();
            sb.AppendLine("Nama,Nopeg,NamaSertifikasi,Lokasi,TanggalMulai,TanggalSelesai,BerlakuDari,BerlakuSampai,Keterangan,Status");

            foreach (var s in list)
            {
                string status = s.IsExpired ? "Expired" : "Aktif";

                string line = string.Join(",",
                    Quote(s.Employee?.NamaLengkap),
                    Quote(s.Employee?.NopegPersero),
                    Quote(s.NamaSertifikasi),
                    Quote(s.Lokasi),
                    Quote(s.TanggalMulai?.ToString("yyyy-MM-dd")),
                    Quote(s.TanggalSelesai?.ToString("yyyy-MM-dd")),
                    Quote(s.BerlakuDari?.ToString("yyyy-MM-dd")),
                    Quote(s.BerlakuSampai?.ToString("yyyy-MM-dd")),
                    Quote(s.Keterangan),
                    Quote(status)
                );

                sb.AppendLine(line);
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            const string fileName = "Data_Sertifikasi.csv";
            return File(bytes, "text/csv", fileName);

            static string Quote(string? value)
            {
                value ??= "";
                if (value.Contains('"') || value.Contains(','))
                    value = "\"" + value.Replace("\"", "\"\"") + "\"";
                return value;
            }
        }

        private IQueryable<Sertifikasi> BuildFilteredQuery(int? employeeId, string? q)
        {
            var query = _db.Sertifikasis
                .Include(s => s.Employee)
                .Include(s => s.MasterSertifikasi)
                .AsQueryable();

            if (employeeId is not null)
                query = query.Where(s => s.EmployeeId == employeeId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(s =>
                    (s.Employee!.NamaLengkap ?? "").Contains(q) ||
                    (s.Employee!.NopegPersero ?? "").Contains(q) ||
                    (s.NamaSertifikasi ?? "").Contains(q));
            }

            return query;
        }

        // ============================ IMPORT ==========================
        [HttpGet]
        public IActionResult Import()
        {
            return PartialView("Import");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (!IsAdminLike)
                return Json(new { success = false, message = "Hanya admin/supervisor yang boleh import data sertifikasi." });

            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            var employees = await _db.Employees
                .Select(e => new { e.Id, e.NamaLengkap, e.NopegPersero })
                .ToListAsync();

            var empDict = employees.ToDictionary(
                e => $"{(e.NamaLengkap ?? "").Trim()}|{(e.NopegPersero ?? "").Trim()}",
                e => e.Id,
                StringComparer.OrdinalIgnoreCase
            );

            // Ambil master sertifikasi yg sudah ada
            var masters = await _db.MasterSertifikasis.ToListAsync();
            // Dictionary nama -> entity master
            var masterDict = masters.ToDictionary(
                m => (m.Nama ?? "").Trim(),
                m => m,
                StringComparer.OrdinalIgnoreCase
            );

            var validList = new List<Sertifikasi>();
            var errorList = new List<string>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using var package = new ExcelPackage(stream);
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null)
                    return Json(new { success = false, message = "Worksheet tidak ditemukan di file." });

                int row = 2;
                while (true)
                {
                    var nama = ws.Cells[row, 1].Text?.Trim();
                    var nopeg = ws.Cells[row, 2].Text?.Trim();
                    var namaSert = ws.Cells[row, 3].Text?.Trim();
                    var lokasi = ws.Cells[row, 4].Text?.Trim();
                    var tMulaiStr = ws.Cells[row, 5].Text?.Trim();
                    var tSelesaiStr = ws.Cells[row, 6].Text?.Trim();
                    var berlakuDariStr = ws.Cells[row, 7].Text?.Trim();
                    var berlakuSampaiStr = ws.Cells[row, 8].Text?.Trim();
                    var ket = ws.Cells[row, 9].Text?.Trim();

                    if (string.IsNullOrEmpty(nama) &&
                        string.IsNullOrEmpty(nopeg) &&
                        string.IsNullOrEmpty(namaSert) &&
                        string.IsNullOrEmpty(tMulaiStr) &&
                        string.IsNullOrEmpty(tSelesaiStr) &&
                        string.IsNullOrEmpty(berlakuDariStr) &&
                        string.IsNullOrEmpty(berlakuSampaiStr) &&
                        string.IsNullOrEmpty(lokasi) &&
                        string.IsNullOrEmpty(ket))
                    {
                        break;
                    }

                    if (string.IsNullOrEmpty(nama) || string.IsNullOrEmpty(nopeg))
                    {
                        errorList.Add($"Baris {row}: Nama/Nopeg kosong.");
                        row++;
                        continue;
                    }

                    var keyEmp = $"{nama}|{nopeg}";
                    if (!empDict.TryGetValue(keyEmp, out int empId))
                    {
                        errorList.Add($"Baris {row}: Kombinasi Nama '{nama}' dan Nopeg '{nopeg}' tidak ditemukan di master Employee.");
                        row++;
                        continue;
                    }

                    DateOnly? tMulai = null, tSelesai = null, berlakuDari = null, berlakuSampai = null;

                    if (!string.IsNullOrEmpty(tMulaiStr))
                    {
                        if (!DateTime.TryParse(tMulaiStr, out var dtMulai))
                        {
                            errorList.Add($"Baris {row}: TanggalMulai tidak valid (format harus yyyy-MM-dd).");
                            row++;
                            continue;
                        }
                        tMulai = DateOnly.FromDateTime(dtMulai);
                    }

                    if (!string.IsNullOrEmpty(tSelesaiStr))
                    {
                        if (!DateTime.TryParse(tSelesaiStr, out var dtSelesai))
                        {
                            errorList.Add($"Baris {row}: TanggalSelesai tidak valid (format harus yyyy-MM-dd).");
                            row++;
                            continue;
                        }
                        tSelesai = DateOnly.FromDateTime(dtSelesai);
                    }

                    if (!string.IsNullOrEmpty(berlakuDariStr))
                    {
                        if (!DateTime.TryParse(berlakuDariStr, out var dtBDari))
                        {
                            errorList.Add($"Baris {row}: BerlakuDari tidak valid (format harus yyyy-MM-dd).");
                            row++;
                            continue;
                        }
                        berlakuDari = DateOnly.FromDateTime(dtBDari);
                    }

                    if (!string.IsNullOrEmpty(berlakuSampaiStr))
                    {
                        if (!DateTime.TryParse(berlakuSampaiStr, out var dtBSampai))
                        {
                            errorList.Add($"Baris {row}: BerlakuSampai tidak valid (format harus yyyy-MM-dd).");
                            row++;
                            continue;
                        }
                        berlakuSampai = DateOnly.FromDateTime(dtBSampai);
                    }

                    // ====== AUTO CREATE MASTER SERTIFIKASI JIKA NAMA BARU ======
                    MasterSertifikasi? masterEntity = null;
                    if (!string.IsNullOrEmpty(namaSert))
                    {
                        var keyNama = namaSert.Trim();

                        if (!masterDict.TryGetValue(keyNama, out masterEntity))
                        {
                            masterEntity = new MasterSertifikasi
                            {
                                Nama = keyNama
                            };
                            _db.MasterSertifikasis.Add(masterEntity);
                            masterDict[keyNama] = masterEntity;
                        }
                    }

                    validList.Add(new Sertifikasi
                    {
                        EmployeeId = empId,
                        MasterSertifikasi = masterEntity, // EF akan set MasterSertifikasiId otomatis
                        NamaSertifikasi = namaSert,
                        Lokasi = lokasi,
                        TanggalMulai = tMulai,
                        TanggalSelesai = tSelesai,
                        BerlakuDari = berlakuDari,
                        BerlakuSampai = berlakuSampai,
                        Keterangan = ket
                    });

                    row++;
                }
            }

            if (validList.Count == 0)
            {
                var msg = "Tidak ada data valid yang dapat diimport.";
                if (errorList.Count > 0)
                    msg += " Contoh error: " + string.Join(" | ", errorList.Take(5));

                return Json(new { success = false, message = msg, errors = errorList });
            }

            await _db.Sertifikasis.AddRangeAsync(validList);
            await _db.SaveChangesAsync();

            var summary = $"Berhasil import {validList.Count} baris sertifikasi.";
            if (errorList.Count > 0)
                summary += $" {errorList.Count} baris gagal (lihat detail error).";

            return Json(new { success = true, message = summary, errors = errorList });
        }
    }
}
