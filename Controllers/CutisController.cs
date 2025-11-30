using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text;
using System.Text.Json;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class CutisController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;
        private const int DEFAULT_QUOTA = 12;

        // EPPlus (once per type)
        static CutisController()
        {
            ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR - PT KPI RU II");
        }

        // Constructor tunggal
        public CutisController(
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager)
        {
            _db = db;
            _userManager = userManager;
        }

        // =========================================================
        // ==================== HELPER CUTI =========================
        // =========================================================

        private async Task<int> UsedDaysThisYearAsync(int employeeId, int year, int? excludeCutiId = null)
        {
            var q = _db.Cutis
                .Where(c => c.EmployeeId == employeeId && c.TanggalMulai.Year == year);

            if (excludeCutiId is not null)
                q = q.Where(c => c.Id != excludeCutiId);

            var list = await q.AsNoTracking().ToListAsync();

            return list.Sum(c =>
            {
                if (c.JumlahHari <= 0)
                {
                    var d = (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1);
                    return d;
                }
                return c.JumlahHari;
            });
        }

        private async Task<int> GetQuotaAsync(int employeeId)
        {
            var emp = await _db.Employees
                .AsNoTracking()
                .FirstOrDefaultAsync(e => e.Id == employeeId);

            return emp?.CutiLimit ?? DEFAULT_QUOTA;
        }

        private async Task<bool> HasOverlapAsync(
            int employeeId,
            DateTime mulai,
            DateTime selesai,
            int? excludeCutiId = null)
        {
            mulai = mulai.Date;
            selesai = selesai.Date;

            var query = _db.Cutis
                .Where(c => c.EmployeeId == employeeId);

            if (excludeCutiId is not null)
                query = query.Where(c => c.Id != excludeCutiId);

            return await query.AnyAsync(c =>
                c.TanggalMulai <= selesai &&
                c.TanggalSelesai >= mulai
            );
        }

        // =========================================================
        // ======================= INDEX ============================
        // =========================================================
        public async Task<IActionResult> Index(string? q)
        {
            int year = DateTime.Today.Year;

            // Base query
            IQueryable<Employee> empQuery = _db.Employees
                .Include(e => e.CutiList)
                .AsNoTracking();

            // Filter search text
            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                empQuery = empQuery.Where(e =>
                    (e.NamaLengkap ?? "").Contains(q) ||
                    (e.NopegPersero ?? "").Contains(q));
            }

            // --- ROLE FILTER --------------------------------------
            bool isAdminLike = User.IsInRole("Admin") || User.IsInRole("SuperAdmin");

            if (!isAdminLike && User.IsInRole("User"))
            {
                // User biasa → hanya lihat cuti dirinya sendiri
                var appUser = await _userManager.GetUserAsync(User);

                if (appUser?.EmployeeId is int empId)
                {
                    empQuery = empQuery.Where(e => e.Id == empId);
                }
                else
                {
                    // kalau user tidak punya relasi Employee, kosongkan saja
                    empQuery = empQuery.Where(e => false);
                }
            }

            var employees = await empQuery
                .OrderBy(e => e.NamaLengkap)
                .ToListAsync();

            var model = employees
                .Select(e =>
                {
                    int kuota = e.CutiLimit ?? DEFAULT_QUOTA;
                    int ambilTh = e.CutiList
                        .Where(c => c.TanggalMulai.Year == year)
                        .Sum(c => c.JumlahHari > 0
                            ? c.JumlahHari
                            : (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1));

                    return new CutiSummaryVM
                    {
                        EmployeeId = e.Id,
                        Nama = e.NamaLengkap ?? "(tanpa nama)",
                        Nopeg = e.NopegPersero ?? "-",
                        Kuota = kuota,
                        Diambil = ambilTh
                    };
                })
                .OrderBy(x => x.Nama)
                .ToList();

            return View(model);
        }

        // =========================================================
        // ====================== DETAILS ===========================
        // =========================================================
        public async Task<IActionResult> Details(int employeeId)
        {
            // Batasi user biasa hanya boleh lihat details dirinya sendiri
            bool isAdminLike = User.IsInRole("Admin") || User.IsInRole("SuperAdmin");
            if (!isAdminLike && User.IsInRole("User"))
            {
                var appUser = await _userManager.GetUserAsync(User);
                if (appUser?.EmployeeId != employeeId)
                {
                    // kalau coba akses employee lain → 403
                    return Forbid();
                }
            }

            var emp = await _db.Employees
                .Include(e => e.CutiList)
                .FirstOrDefaultAsync(e => e.Id == employeeId);

            if (emp == null) return NotFound();

            int year = DateTime.Today.Year;
            int kuota = emp.CutiLimit ?? DEFAULT_QUOTA;

            int terpakai = emp.CutiList
                .Where(c => c.TanggalMulai.Year == year)
                .Sum(c => c.JumlahHari > 0
                    ? c.JumlahHari
                    : (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1));

            var vm = new CutiDetailsVM
            {
                Emp = emp,
                Kuota = kuota,
                Terpakai = terpakai,
                Sisa = kuota - terpakai,
                Items = emp.CutiList.OrderByDescending(c => c.TanggalMulai).ToList()
            };

            return View(vm);
        }

        // =========================================================
        // ====================== CREATE (GET) ======================
        // =========================================================
        public async Task<IActionResult> Create(int? employeeId)
        {
            ViewBag.EmployeeId = new SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", employeeId
            );

            int year = DateTime.Today.Year;
            var employees = await _db.Employees
                .Include(e => e.CutiList)
                .ToListAsync();

            var usageDict = employees.ToDictionary(
                e => e.Id,
                e => e.CutiList
                    .Where(c => c.TanggalMulai.Year == year)
                    .Sum(c => c.JumlahHari > 0
                        ? c.JumlahHari
                        : (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1))
            );

            var quotaDict = employees.ToDictionary(
                e => e.Id,
                e => e.CutiLimit ?? DEFAULT_QUOTA
            );

            ViewBag.UsageJson = JsonSerializer.Serialize(usageDict);
            ViewBag.QuotaJson = JsonSerializer.Serialize(quotaDict);

            return PartialView("Create", new Cuti
            {
                EmployeeId = employeeId ?? 0,
                TanggalMulai = DateTime.Today,
                TanggalSelesai = DateTime.Today,
                JumlahHari = 1
            });
        }

        // ====================== CREATE (POST) =====================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(
            [Bind("Id,EmployeeId,TanggalMulai,TanggalSelesai,JumlahHari,Tujuan,Keterangan")]
            Cuti model)
        {
            model.TanggalMulai = model.TanggalMulai.Date;
            model.TanggalSelesai = model.TanggalSelesai.Date;

            if (model.TanggalSelesai < model.TanggalMulai)
                ModelState.AddModelError("TanggalSelesai", "Tanggal selesai tidak boleh sebelum mulai.");

            int defaultHari = (int)Math.Max(
                1,
                (model.TanggalSelesai - model.TanggalMulai).TotalDays + 1
            );

            if (model.JumlahHari <= 0)
                model.JumlahHari = defaultHari;

            if (model.JumlahHari <= 0)
                ModelState.AddModelError("JumlahHari", "Jumlah hari cuti harus minimal 1 hari.");

            if (await HasOverlapAsync(model.EmployeeId, model.TanggalMulai, model.TanggalSelesai))
            {
                ModelState.AddModelError(string.Empty,
                    "Tanggal cuti yang diajukan bertumpukan dengan cuti yang sudah ada.");
            }

            int year = model.TanggalMulai.Year;
            int kuota = await GetQuotaAsync(model.EmployeeId);
            int used = await UsedDaysThisYearAsync(model.EmployeeId, year);

            if (used + model.JumlahHari > kuota)
            {
                ModelState.AddModelError(string.Empty,
                    $"Kuota cuti tahun {year} sudah habis / terlampaui. " +
                    $"Terpakai {used} dari {kuota} hari. " +
                    $"Pengajuan {model.JumlahHari} hari membuat total {used + model.JumlahHari} hari, melebihi kuota.");
            }

            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Data tidak valid.",
                    errors,
                    employeeId = model.EmployeeId
                });
            }

            _db.Add(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data cuti berhasil disimpan.",
                employeeId = model.EmployeeId
            });
        }

        // =========================================================
        // ======================= EDIT (GET) =======================
        // =========================================================
        public async Task<IActionResult> Edit(int id)
        {
            var c = await _db.Cutis.FindAsync(id);
            if (c == null) return NotFound();

            ViewBag.EmployeeId = new SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", c.EmployeeId
            );

            int year = DateTime.Today.Year;
            var employees = await _db.Employees.Include(e => e.CutiList).ToListAsync();

            var usageDict = employees.ToDictionary(
                e => e.Id,
                e => e.CutiList
                    .Where(cuti => cuti.Id != id && cuti.TanggalMulai.Year == year)
                    .Sum(cuti => cuti.JumlahHari > 0
                        ? cuti.JumlahHari
                        : (int)Math.Max(1, (cuti.TanggalSelesai.Date - cuti.TanggalMulai.Date).TotalDays + 1))
            );

            var quotaDict = employees.ToDictionary(
                e => e.Id,
                e => e.CutiLimit ?? DEFAULT_QUOTA
            );

            ViewBag.UsageJson = JsonSerializer.Serialize(usageDict);
            ViewBag.QuotaJson = JsonSerializer.Serialize(quotaDict);

            if (c.JumlahHari <= 0)
            {
                c.JumlahHari = (int)Math.Max(
                    1,
                    (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1
                );
            }

            return PartialView("Edit", c);
        }

        // ======================= EDIT (POST) ======================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(
            int id,
            [Bind("Id,EmployeeId,TanggalMulai,TanggalSelesai,JumlahHari,Tujuan,Keterangan")]
            Cuti model)
        {
            if (id != model.Id) return BadRequest();

            model.TanggalMulai = model.TanggalMulai.Date;
            model.TanggalSelesai = model.TanggalSelesai.Date;

            if (model.TanggalSelesai < model.TanggalMulai)
                ModelState.AddModelError("TanggalSelesai", "Tanggal selesai tidak boleh sebelum mulai.");

            int defaultHari = (int)Math.Max(
                1,
                (model.TanggalSelesai - model.TanggalMulai).TotalDays + 1
            );

            if (model.JumlahHari <= 0)
                model.JumlahHari = defaultHari;

            if (model.JumlahHari <= 0)
                ModelState.AddModelError("JumlahHari", "Jumlah hari cuti harus minimal 1 hari.");

            if (await HasOverlapAsync(model.EmployeeId, model.TanggalMulai, model.TanggalSelesai, excludeCutiId: model.Id))
            {
                ModelState.AddModelError(string.Empty,
                    "Tanggal cuti yang diubah bertumpukan dengan cuti yang sudah ada.");
            }

            int year = model.TanggalMulai.Year;
            int kuota = await GetQuotaAsync(model.EmployeeId);
            int used = await UsedDaysThisYearAsync(model.EmployeeId, year, excludeCutiId: model.Id);

            if (used + model.JumlahHari > kuota)
            {
                ModelState.AddModelError(string.Empty,
                    $"Kuota cuti tahun {year} sudah habis / terlampaui. " +
                    $"Terpakai {used} dari {kuota} hari. " +
                    $"Perubahan {model.JumlahHari} hari membuat total {used + model.JumlahHari} hari, melebihi kuota.");
            }

            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Data tidak valid.",
                    errors,
                    employeeId = model.EmployeeId
                });
            }

            _db.Update(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data cuti berhasil diupdate.",
                employeeId = model.EmployeeId
            });
        }

        // =========================================================
        // ======================= DELETE (POST) ====================
        // =========================================================
        [HttpPost]
        public async Task<IActionResult> Delete(int id)
        {
            var c = await _db.Cutis.FindAsync(id);
            if (c == null)
                return Json(new { success = false, message = "Data tidak ditemukan." });

            int empId = c.EmployeeId;

            _db.Cutis.Remove(c);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data cuti berhasil dihapus.",
                employeeId = empId
            });
        }

        // =========================================================
        // ================= TEMPLATE / EXPORT ======================
        // =========================================================

        [HttpGet]
        public IActionResult DownloadTemplateExcel()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("TemplateCuti");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Nopeg";
            ws.Cells[1, 3].Value = "TanggalMulai (yyyy-MM-dd)";
            ws.Cells[1, 4].Value = "TanggalSelesai (yyyy-MM-dd)";
            ws.Cells[1, 5].Value = "JumlahHari (optional)";
            ws.Cells[1, 6].Value = "Tujuan";
            ws.Cells[1, 7].Value = "Keterangan";

            using (var range = ws.Cells[1, 1, 1, 7])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                     .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            const string fileName = "Template_Import_Cuti.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(bytes, contentType, fileName);
        }

        [HttpGet]
        public IActionResult DownloadTemplateCsv()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Nama,Nopeg,TanggalMulai (yyyy-MM-dd),TanggalSelesai (yyyy-MM-dd),JumlahHari (optional),Tujuan,Keterangan");
            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            const string fileName = "Template_Import_Cuti.csv";
            return File(bytes, "text/csv", fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportExcel(string? q, int? employeeId)
        {
            var query = _db.Cutis
                .Include(c => c.Employee)
                .AsQueryable();

            if (employeeId.HasValue)
            {
                query = query.Where(c => c.EmployeeId == employeeId.Value);
            }

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(c =>
                    (c.Employee!.NamaLengkap ?? "").Contains(q) ||
                    (c.Employee!.NopegPersero ?? "").Contains(q));
            }

            var cutis = await query
                .OrderBy(c => c.Employee!.NamaLengkap)
                .ThenByDescending(c => c.TanggalMulai)
                .ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("DataCuti");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Nopeg";
            ws.Cells[1, 3].Value = "TanggalMulai";
            ws.Cells[1, 4].Value = "TanggalSelesai";
            ws.Cells[1, 5].Value = "JumlahHari";
            ws.Cells[1, 6].Value = "Tujuan";
            ws.Cells[1, 7].Value = "Keterangan";

            using (var range = ws.Cells[1, 1, 1, 7])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                     .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            int row = 2;
            foreach (var c in cutis)
            {
                ws.Cells[row, 1].Value = c.Employee?.NamaLengkap;
                ws.Cells[row, 2].Value = c.Employee?.NopegPersero;
                ws.Cells[row, 3].Value = c.TanggalMulai.ToString("yyyy-MM-dd");
                ws.Cells[row, 4].Value = c.TanggalSelesai.ToString("yyyy-MM-dd");
                ws.Cells[row, 5].Value = c.JumlahHari <= 0
                    ? (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1)
                    : c.JumlahHari;
                ws.Cells[row, 6].Value = c.Tujuan;
                ws.Cells[row, 7].Value = c.Keterangan;
                row++;
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            const string fileName = "Data_Cuti.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(bytes, contentType, fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportCsv(string? q, int? employeeId)
        {
            var query = _db.Cutis
                .Include(c => c.Employee)
                .AsQueryable();

            if (employeeId.HasValue)
            {
                query = query.Where(c => c.EmployeeId == employeeId.Value);
            }

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(c =>
                    (c.Employee!.NamaLengkap ?? "").Contains(q) ||
                    (c.Employee!.NopegPersero ?? "").Contains(q));
            }

            var cutis = await query
                .OrderBy(c => c.Employee!.NamaLengkap)
                .ThenByDescending(c => c.TanggalMulai)
                .ToListAsync();

            var sb = new StringBuilder();
            sb.AppendLine("Nama,Nopeg,TanggalMulai,TanggalSelesai,JumlahHari,Tujuan,Keterangan");

            foreach (var c in cutis)
            {
                int jml = c.JumlahHari <= 0
                    ? (int)Math.Max(1, (c.TanggalSelesai.Date - c.TanggalMulai.Date).TotalDays + 1)
                    : c.JumlahHari;

                string line = string.Join(",",
                    Quote(c.Employee?.NamaLengkap),
                    Quote(c.Employee?.NopegPersero),
                    c.TanggalMulai.ToString("yyyy-MM-dd"),
                    c.TanggalSelesai.ToString("yyyy-MM-dd"),
                    jml.ToString(),
                    Quote(c.Tujuan),
                    Quote(c.Keterangan)
                );

                sb.AppendLine(line);
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            const string fileName = "Data_Cuti.csv";
            return File(bytes, "text/csv", fileName);

            static string Quote(string? value)
            {
                value ??= "";
                if (value.Contains('"') || value.Contains(','))
                    value = "\"" + value.Replace("\"", "\"\"") + "\"";
                return value;
            }
        }

        // =========================================================
        // ======================== IMPORT ==========================
        // =========================================================
        [HttpGet]
        public IActionResult Import()
        {
            return PartialView("Import");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            var employees = await _db.Employees
                .Select(e => new { e.Id, e.NamaLengkap, e.NopegPersero, e.CutiLimit })
                .ToListAsync();

            var empDict = employees.ToDictionary(
                e => $"{(e.NamaLengkap ?? "").Trim()}|{(e.NopegPersero ?? "").Trim()}",
                e => e.Id,
                StringComparer.OrdinalIgnoreCase
            );

            var extraUsedDict = new Dictionary<(int EmpId, int Year), int>();
            var extraRangesDict = new Dictionary<(int EmpId, int Year), List<(DateTime Start, DateTime End)>>();

            var validList = new List<Cuti>();
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
                    var tMulaiStr = ws.Cells[row, 3].Text?.Trim();
                    var tSelesaiStr = ws.Cells[row, 4].Text?.Trim();
                    var jmlStr = ws.Cells[row, 5].Text?.Trim();
                    var tujuan = ws.Cells[row, 6].Text?.Trim();
                    var ket = ws.Cells[row, 7].Text?.Trim();

                    if (string.IsNullOrEmpty(nama) && string.IsNullOrEmpty(nopeg) &&
                        string.IsNullOrEmpty(tMulaiStr) && string.IsNullOrEmpty(tSelesaiStr))
                        break;

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

                    if (!DateTime.TryParse(tMulaiStr, out var tMulai) ||
                        !DateTime.TryParse(tSelesaiStr, out var tSelesai))
                    {
                        errorList.Add($"Baris {row}: Tanggal tidak valid (format harus yyyy-MM-dd).");
                        row++;
                        continue;
                    }

                    tMulai = tMulai.Date;
                    tSelesai = tSelesai.Date;

                    if (tSelesai < tMulai)
                    {
                        errorList.Add($"Baris {row}: Tanggal selesai sebelum tanggal mulai.");
                        row++;
                        continue;
                    }

                    int defaultHari = (int)Math.Max(1, (tSelesai - tMulai).TotalDays + 1);
                    int jumlahHari = defaultHari;

                    if (int.TryParse(jmlStr, out var tmp) && tmp > 0)
                        jumlahHari = tmp;

                    int year = tMulai.Year;
                    var keyYearEmp = (EmpId: empId, Year: year);

                    bool overlapDb = await HasOverlapAsync(empId, tMulai, tSelesai);
                    if (overlapDb)
                    {
                        errorList.Add($"Baris {row}: Range {tMulai:dd/MM/yyyy}–{tSelesai:dd/MM/yyyy} bertumpukan dengan data cuti yang sudah ada.");
                        row++;
                        continue;
                    }

                    if (extraRangesDict.TryGetValue(keyYearEmp, out var ranges))
                    {
                        bool overlapFile = ranges.Any(r =>
                            r.Start <= tSelesai && r.End >= tMulai);
                        if (overlapFile)
                        {
                            errorList.Add($"Baris {row}: Range {tMulai:dd/MM/yyyy}–{tSelesai:dd/MM/yyyy} bertumpukan dengan baris lain dalam file import.");
                            row++;
                            continue;
                        }
                    }
                    else
                    {
                        extraRangesDict[keyYearEmp] = new List<(DateTime Start, DateTime End)>();
                    }

                    int kuota = await GetQuotaAsync(empId);
                    int usedDb = await UsedDaysThisYearAsync(empId, year);
                    extraUsedDict.TryGetValue(keyYearEmp, out int usedFile);

                    if (usedDb + usedFile + jumlahHari > kuota)
                    {
                        errorList.Add(
                            $"Baris {row}: Kuota cuti tahun {year} terlampaui. " +
                            $"Terpakai DB {usedDb} hari, dari file sebelumnya {usedFile} hari, " +
                            $"tambah baris ini {jumlahHari} hari => total {usedDb + usedFile + jumlahHari} > kuota {kuota}."
                        );
                        row++;
                        continue;
                    }

                    extraUsedDict[keyYearEmp] = usedFile + jumlahHari;
                    extraRangesDict[keyYearEmp].Add((tMulai, tSelesai));

                    validList.Add(new Cuti
                    {
                        EmployeeId = empId,
                        TanggalMulai = tMulai,
                        TanggalSelesai = tSelesai,
                        JumlahHari = jumlahHari,
                        Tujuan = tujuan,
                        Keterangan = ket
                    });

                    row++;
                }
            }

            if (validList.Count == 0)
            {
                var msg = "Tidak ada data valid yang dapat diimport.";
                if (errorList.Count > 0)
                    msg += " Beberapa error: " + string.Join(" | ", errorList.Take(5));

                return Json(new { success = false, message = msg, errors = errorList });
            }

            await _db.Cutis.AddRangeAsync(validList);
            await _db.SaveChangesAsync();

            var summary = $"Berhasil import {validList.Count} baris cuti.";
            if (errorList.Count > 0)
                summary += $" {errorList.Count} baris gagal (lihat detail error).";

            return Json(new { success = true, message = summary, errors = errorList });
        }

        // =========================================================
        // ===================== VIEWMODELS =========================
        // =========================================================
        public class CutiSummaryVM
        {
            public int EmployeeId { get; set; }
            public string Nama { get; set; } = "";
            public string Nopeg { get; set; } = "-";
            public int Kuota { get; set; }
            public int Diambil { get; set; }
            public int Sisa => Kuota - Diambil;
        }

        public class CutiDetailsVM
        {
            public Employee Emp { get; set; } = new();
            public int Kuota { get; set; }
            public int Terpakai { get; set; }
            public int Sisa { get; set; }
            public List<Cuti> Items { get; set; } = new();
        }
    }
}
