using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize(Roles = "Admin,SuperAdmin,Supervisor")]
    public class WorkJobNamesController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly ILogger<WorkJobNamesController> _logger;

        // ===================== EPPlus v8 =====================
        private static bool _epplusReady;
        private static void EnsureEpplus()
        {
            if (_epplusReady) return;

            ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR - PT KPI RU II");
            _epplusReady = true;
        }

        public WorkJobNamesController(
            ApplicationDbContext db,
            ILogger<WorkJobNamesController> logger)
        {
            _db = db;
            _logger = logger;
        }

        // ============== INDEX ==============
        public async Task<IActionResult> Index(string? q)
        {
            var query = _db.WorkJobNames.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(x => x.NamaPekerjaan.Contains(q));
            }

            var list = await query
                .OrderBy(x => x.NamaPekerjaan)
                .ToListAsync();

            ViewBag.Search = q;
            ViewBag.Message = TempData["Message"] as string;
            ViewBag.Error = TempData["Error"] as string;

            return View(list);
        }

        // ============== CREATE ==============
        [HttpGet]
        public IActionResult Create()
        {
            var model = new WorkJobName();
            return View(model); // Layout = null di view
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(WorkJobName model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            try
            {
                model.NamaPekerjaan = model.NamaPekerjaan?.Trim() ?? string.Empty;
                model.CreatedAt = DateTime.UtcNow;

                var exists = await _db.WorkJobNames
                    .AnyAsync(x => x.NamaPekerjaan == model.NamaPekerjaan);

                if (exists)
                {
                    ModelState.AddModelError(nameof(model.NamaPekerjaan), "Nama pekerjaan sudah ada.");
                    return View(model);
                }

                _db.WorkJobNames.Add(model);
                await _db.SaveChangesAsync();

                TempData["Message"] = "Nama pekerjaan berhasil ditambahkan.";
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal menambah WorkJobName");
                ModelState.AddModelError(string.Empty, "Terjadi kesalahan saat menyimpan data.");
                return View(model);
            }
        }

        // ============== EDIT ==============
        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            var job = await _db.WorkJobNames.FindAsync(id);
            if (job == null)
                return NotFound();

            return View(job);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, WorkJobName model)
        {
            if (id != model.Id)
                return NotFound();

            if (!ModelState.IsValid)
                return View(model);

            var job = await _db.WorkJobNames.FindAsync(id);
            if (job == null)
                return NotFound();

            try
            {
                var newName = model.NamaPekerjaan?.Trim() ?? string.Empty;

                var exists = await _db.WorkJobNames
                    .AnyAsync(x => x.Id != id && x.NamaPekerjaan == newName);

                if (exists)
                {
                    ModelState.AddModelError(nameof(model.NamaPekerjaan), "Nama pekerjaan sudah ada.");
                    return View(model);
                }

                job.NamaPekerjaan = newName;
                await _db.SaveChangesAsync();

                TempData["Message"] = "Nama pekerjaan berhasil diperbarui.";
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal mengupdate WorkJobName Id={Id}", id);
                ModelState.AddModelError(string.Empty, "Terjadi kesalahan saat menyimpan data.");
                return View(model);
            }
        }

        // ============== DELETE ==============
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var job = await _db.WorkJobNames.FindAsync(id);
            if (job == null)
            {
                TempData["Error"] = "Data nama pekerjaan tidak ditemukan.";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                // CEK: apakah NamaPekerjaan ini masih dipakai di WorkTasks?
                var nama = job.NamaPekerjaan ?? string.Empty;

                bool inUse = await _db.WorkTasks
                    .AnyAsync(t => t.NamaPekerjaan == nama);

                if (inUse)
                {
                    TempData["Error"] =
                        "Nama pekerjaan tidak dapat dihapus karena masih digunakan pada data pekerjaan.";
                    return RedirectToAction(nameof(Index));
                }

                _db.WorkJobNames.Remove(job);
                await _db.SaveChangesAsync();

                TempData["Message"] = "Nama pekerjaan berhasil dihapus.";
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal menghapus WorkJobName Id={Id}", id);
                TempData["Error"] = "Terjadi kesalahan saat menghapus data.";
                return RedirectToAction(nameof(Index));
            }
        }

        // ============== TEMPLATE IMPORT MASTER ==============
        [HttpGet]
        public IActionResult TemplateImport()
        {
            EnsureEpplus();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Template");

            ws.Cells[1, 1].Value = "TEMPLATE MASTER NAMA PEKERJAAN UTILITIES";
            ws.Cells[1, 1, 1, 2].Merge = true;
            ws.Cells[1, 1, 1, 2].Style.Font.Bold = true;

            ws.Cells[2, 1].Value = "No";
            ws.Cells[2, 2].Value = "NamaPekerjaan*";

            ws.Cells[2, 1, 2, 2].Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            var fileName = $"template-master-nama-pekerjaan-{DateTime.Now:yyyyMMdd}.xlsx";

            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ============== IMPORT DARI EXCEL (JSON) ==============
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile? file)
        {
            EnsureEpplus();

            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            int imported = 0;
            int row = 3;

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

                if (!title.Contains("TEMPLATE MASTER NAMA PEKERJAAN", StringComparison.OrdinalIgnoreCase) ||
                    !hNama.Contains("NamaPekerjaan", StringComparison.OrdinalIgnoreCase))
                {
                    return Json(new
                    {
                        success = false,
                        message = "Template tidak dikenali sebagai template master nama pekerjaan."
                    });
                }

                while (true)
                {
                    var nama = ws.Cells[row, 2].GetValue<string>()?.Trim();

                    if (string.IsNullOrWhiteSpace(nama))
                        break;

                    var exists = await _db.WorkJobNames
                        .AnyAsync(x => x.NamaPekerjaan == nama);

                    if (!exists)
                    {
                        _db.WorkJobNames.Add(new WorkJobName
                        {
                            NamaPekerjaan = nama,
                            CreatedAt = DateTime.UtcNow
                        });
                        imported++;
                    }

                    row++;
                }

                await _db.SaveChangesAsync();

                return Json(new
                {
                    success = true,
                    imported,
                    message = $"Berhasil mengimpor {imported} nama pekerjaan."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Import WorkJobNames gagal pada baris {Row}", row);
                return Json(new
                {
                    success = false,
                    message = $"Import gagal pada baris {row}: {ex.Message}"
                });
            }
        }

        // ============== IMPORT DIALOG (POPUP) ==============
        [HttpGet]
        public IActionResult ImportDialog()
        {
            // View ini hanya berisi form upload (Layout = null)
            return PartialView("ImportDialog");
        }

        // ============== OPTIONS UNTUK SELECT2 (FORM WORKTASK) ==============
        [HttpGet]
        [AllowAnonymous]
        public async Task<IActionResult> Options(string? term)
        {
            var q = _db.WorkJobNames.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(term))
            {
                term = term.Trim();
                q = q.Where(x => x.NamaPekerjaan.Contains(term));
            }

            var list = await q
                .OrderBy(x => x.NamaPekerjaan)
                .Take(50)
                .ToListAsync();

            var result = list.Select(x => new
            {
                id = x.NamaPekerjaan,
                text = x.NamaPekerjaan
            });

            return Json(result);
        }
    }
}
