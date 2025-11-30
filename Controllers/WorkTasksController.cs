using System;
using System.Collections.Generic;
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
    [Authorize]
    public class WorkTasksController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly ILogger<WorkTasksController> _logger;

        // ===================== EPPlus v8 =====================
        private static bool _epplusReady;
        private static void EnsureEpplus()
        {
            if (_epplusReady) return;
            ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR - PT KPI RU II");
            _epplusReady = true;
        }

        public WorkTasksController(ApplicationDbContext db, ILogger<WorkTasksController> logger)
        {
            _db = db;
            _logger = logger;
        }

        private bool IsAdminLike =>
            User.IsInRole("Admin") ||
            User.IsInRole("SuperAdmin") ||
            User.IsInRole("Supervisor");

        // ===================== MESSAGE HELPER (SweetAlert) =====================
        private const string TaskMessageKey = "TaskMessage";
        private const string TaskMessageTypeKey = "TaskMessageType";

        private void SetTaskMessage(string message, string type = "success")
        {
            TempData[TaskMessageKey] = message;
            TempData[TaskMessageTypeKey] = type;
        }

        // helper: cek request AJAX / fetch
        private bool IsAjaxRequest()
        {
            var xReq = Request.Headers["X-Requested-With"].ToString();
            if (!string.IsNullOrEmpty(xReq) &&
                xReq.Equals("XMLHttpRequest", StringComparison.OrdinalIgnoreCase))
                return true;

            var accept = Request.Headers["Accept"].ToString();
            if (!string.IsNullOrEmpty(accept) &&
                accept.IndexOf("application/json", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            return false;
        }

        // ============== INDEX (LIST + SUMMARY PER PIC) ==============
        public async Task<IActionResult> Index(
            string? pic,
            WorkTaskStatus? status,
            WorkTaskPriority? priority,
            string? q,
            bool onlyOutstanding = false)
        {
            var tasks = _db.WorkTasks.AsQueryable();

            // filter search
            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                tasks = tasks.Where(t =>
                    (t.NamaPekerjaan ?? "").Contains(q) ||
                    (t.NamaRequest ?? "").Contains(q) ||
                    (t.NamaPIC ?? "").Contains(q) ||
                    (t.KeteranganTindakLanjut ?? "").Contains(q));
            }

            // filter PIC
            if (!string.IsNullOrWhiteSpace(pic))
            {
                pic = pic.Trim();
                tasks = tasks.Where(t => (t.NamaPIC ?? "").Contains(pic));
            }

            // filter status
            if (status.HasValue)
            {
                tasks = tasks.Where(t => t.Status == status.Value);
            }

            // filter prioritas
            if (priority.HasValue)
            {
                tasks = tasks.Where(t => t.Prioritas == priority.Value);
            }

            // hanya non-done (outstanding)
            if (onlyOutstanding)
            {
                tasks = tasks.Where(t => t.Status != WorkTaskStatus.Done);
            }

            // ===== CARD SUMMARY =====
            var totalTasks = await tasks.CountAsync();
            var pendingTasks = await tasks.Where(t => t.Status == WorkTaskStatus.Pending).CountAsync();
            var inProgressTasks = await tasks.Where(t => t.Status == WorkTaskStatus.InProgress).CountAsync();
            var doneTasks = await tasks.Where(t => t.Status == WorkTaskStatus.Done).CountAsync();

            ViewBag.TotalTasks = totalTasks;
            ViewBag.PendingTasks = pendingTasks;
            ViewBag.InProgressTasks = inProgressTasks;
            ViewBag.DoneTasks = doneTasks;

            // ===== LIST UNTUK VIEW (nanti dikelompokkan per PIC di Razor) =====
            var list = await tasks
                .AsNoTracking()
                .OrderBy(t => t.Status)
                .ThenByDescending(t => t.Prioritas)
                .ThenBy(t => t.DueDate ?? t.StartDate)
                .ToListAsync();

            // list PIC untuk filter dropdown (kalau mau dipakai nanti)
            var picList = await _db.WorkTasks
                .Where(t => t.NamaPIC != null && t.NamaPIC != "")
                .Select(t => t.NamaPIC!)
                .Distinct()
                .OrderBy(x => x)
                .ToListAsync();
            ViewBag.PicList = picList;

            // simpan filter ke ViewBag
            ViewBag.Search = q;
            ViewBag.FilterPIC = pic;
            ViewBag.FilterStatus = status;
            ViewBag.FilterPrioritas = priority;
            ViewBag.OnlyOutstanding = onlyOutstanding;

            // query string utk Export
            ViewBag.qsPic = pic;
            ViewBag.qsStatus = status?.ToString();
            ViewBag.qsPriority = priority?.ToString();
            ViewBag.qsSearch = q;
            ViewBag.qsOnlyOutstanding = onlyOutstanding;

            // pesan dari TempData (untuk SweetAlert toast di Index)
            ViewBag.TaskMessage = TempData[TaskMessageKey] as string;
            ViewBag.TaskMessageType = TempData[TaskMessageTypeKey] as string;

            return View(list);
        }

        // ============== DETAIL SATU TASK (kalau masih mau pakai modal) ==============
        public async Task<IActionResult> Details(int id)
        {
            var task = await _db.WorkTasks
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id);

            if (task == null) return NotFound();
            return PartialView("Details", task);
        }

        // ============== DETAIL PER PIC (HALAMAN PISAH) ==============
        [HttpGet]
        public async Task<IActionResult> ByPic(string pic)
        {
            if (string.IsNullOrWhiteSpace(pic))
                return RedirectToAction(nameof(Index));

            var trimmedPic = pic.Trim();

            var tasks = await _db.WorkTasks
                .AsNoTracking()
                .Where(t => (t.NamaPIC ?? "").Trim() == trimmedPic)
                .OrderBy(t => t.Status)
                .ThenByDescending(t => t.Prioritas)
                .ThenBy(t => t.DueDate ?? t.StartDate)
                .ToListAsync();

            var vm = new WorkTasksByPicVM
            {
                PicName = trimmedPic,
                Tasks = tasks
            };

            // untuk tombol "Tambah Pekerjaan" dari halaman ByPic -> PIC locked
            ViewBag.FixedPicName = trimmedPic;

            // pesan SweetAlert setelah redirect ke ByPic
            ViewBag.TaskMessage = TempData[TaskMessageKey] as string;
            ViewBag.TaskMessageType = TempData[TaskMessageTypeKey] as string;

            return View(vm);
        }

        // ============== CREATE (MODAL / PAGE) ==============
        [HttpGet]
        public IActionResult Create(string? pic)
        {
            var fixedPic = string.IsNullOrWhiteSpace(pic) ? null : pic.Trim();

            var model = new WorkTask
            {
                StartDate = DateTime.Today,
                DueDate = DateTime.Today,
                Status = WorkTaskStatus.Pending,
                Prioritas = WorkTaskPriority.Medium,
                ProgressPercent = 0,
                NamaPIC = fixedPic
            };

            // kalau dipanggil dari ByPic, kita lock PIC ini (read-only di form)
            ViewBag.FixedPicName = fixedPic;

            return PartialView("Create", model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(WorkTask model)
        {
            var isAjax = IsAjaxRequest();

            if (!ModelState.IsValid)
            {
                if (isAjax)
                {
                    var errs = ModelState.Values.SelectMany(v => v.Errors)
                        .Select(e => e.ErrorMessage).ToList();
                    return Json(new { success = false, message = "Data tidak valid.", errors = errs });
                }

                // submit biasa (page Create penuh)
                return View(model);
            }

            try
            {
                SyncStatusAndProgress(model);

                _db.WorkTasks.Add(model);
                await _db.SaveChangesAsync();

                // upsert nama pekerjaan ke master
                await UpsertJobNameAsync(model.NamaPekerjaan);

                if (isAjax)
                {
                    // mode modal / fetch: tetap balikan JSON
                    return Json(new { success = true, message = "Pekerjaan berhasil ditambahkan." });
                }

                // mode normal form: pakai TempData + redirect ke ByPic (kalau ada PIC)
                SetTaskMessage("Pekerjaan berhasil ditambahkan.", "success");

                if (!string.IsNullOrWhiteSpace(model.NamaPIC))
                {
                    var picName = model.NamaPIC.Trim();
                    return RedirectToAction(nameof(ByPic), new { pic = picName });
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal create WorkTask");

                if (isAjax)
                {
                    return Json(new { success = false, message = "Terjadi kesalahan saat menyimpan data." });
                }

                ModelState.AddModelError(string.Empty, "Terjadi kesalahan saat menyimpan data.");
                return View(model);
            }
        }

        // ============== EDIT (MODAL / PAGE) ==============
        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            var task = await _db.WorkTasks.FindAsync(id);
            if (task == null) return NotFound();

            return PartialView("Edit", task);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, WorkTask model)
        {
            var isAjax = IsAjaxRequest();

            if (id != model.Id)
            {
                if (isAjax)
                    return Json(new { success = false, message = "Permintaan tidak valid." });

                return BadRequest();
            }

            if (!ModelState.IsValid)
            {
                if (isAjax)
                {
                    var errs = ModelState.Values.SelectMany(v => v.Errors)
                        .Select(e => e.ErrorMessage).ToList();
                    return Json(new { success = false, message = "Data tidak valid.", errors = errs });
                }

                return View(model);
            }

            var task = await _db.WorkTasks.FindAsync(id);
            if (task == null)
            {
                if (isAjax)
                    return Json(new { success = false, message = "Data tidak ditemukan." });

                return NotFound();
            }

            try
            {
                task.NamaPekerjaan = model.NamaPekerjaan;
                task.StartDate = model.StartDate;
                task.DueDate = model.DueDate;
                task.CompletedDate = model.CompletedDate;
                task.NamaRequest = model.NamaRequest;
                task.ProgressPercent = model.ProgressPercent;
                task.Status = model.Status;
                task.KeteranganTindakLanjut = model.KeteranganTindakLanjut;
                task.Prioritas = model.Prioritas;
                task.NamaPIC = model.NamaPIC;

                SyncStatusAndProgress(task);

                await _db.SaveChangesAsync();
                await UpsertJobNameAsync(task.NamaPekerjaan);

                if (isAjax)
                {
                    return Json(new { success = true, message = "Pekerjaan berhasil diperbarui." });
                }

                SetTaskMessage("Pekerjaan berhasil diperbarui.", "success");

                var picName = (task.NamaPIC ?? "").Trim();
                if (!string.IsNullOrWhiteSpace(picName))
                {
                    return RedirectToAction(nameof(ByPic), new { pic = picName });
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal update WorkTask Id={Id}", id);

                if (isAjax)
                    return Json(new { success = false, message = "Terjadi kesalahan saat menyimpan data." });

                ModelState.AddModelError(string.Empty, "Terjadi kesalahan saat menyimpan data.");
                return View(model);
            }
        }

        // ============== DELETE (AJAX / NON-AJAX) ==============
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var isAjax = IsAjaxRequest();

            var task = await _db.WorkTasks.FindAsync(id);
            if (task == null)
            {
                if (isAjax)
                    return Json(new { success = false, message = "Data pekerjaan tidak ditemukan." });

                SetTaskMessage("Data pekerjaan tidak ditemukan.", "error");
                return RedirectToAction(nameof(Index));
            }

            var picName = (task.NamaPIC ?? "").Trim();

            try
            {
                _db.WorkTasks.Remove(task);
                await _db.SaveChangesAsync();

                if (isAjax)
                {
                    return Json(new { success = true, message = "Pekerjaan berhasil dihapus." });
                }

                SetTaskMessage("Pekerjaan berhasil dihapus.", "success");

                if (!string.IsNullOrWhiteSpace(picName))
                    return RedirectToAction(nameof(ByPic), new { pic = picName });

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gagal menghapus WorkTask Id={Id}", id);

                if (isAjax)
                    return Json(new { success = false, message = "Terjadi kesalahan saat menghapus data." });

                SetTaskMessage("Terjadi kesalahan saat menghapus data.", "error");
                if (!string.IsNullOrWhiteSpace(picName))
                    return RedirectToAction(nameof(ByPic), new { pic = picName });

                return RedirectToAction(nameof(Index));
            }
        }

        // ============== BULK DELETE ==============
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> BulkDelete([FromForm] int[] ids)
        {
            var isAjax = IsAjaxRequest();

            if (ids == null || ids.Length == 0)
            {
                if (isAjax)
                    return Json(new { success = false, message = "Tidak ada data yang dipilih." });

                SetTaskMessage("Tidak ada data yang dipilih.", "warning");
                return RedirectToAction(nameof(Index));
            }

            try
            {
                var tasks = await _db.WorkTasks
                    .Where(t => ids.Contains(t.Id))
                    .ToListAsync();

                if (tasks.Count == 0)
                {
                    if (isAjax)
                        return Json(new { success = false, message = "Data pekerjaan tidak ditemukan." });

                    SetTaskMessage("Data pekerjaan tidak ditemukan.", "error");
                    return RedirectToAction(nameof(Index));
                }

                // diasumsikan BulkDelete dipakai dari ByPic -> semua PIC sama
                var picName = (tasks.First().NamaPIC ?? "").Trim();

                _db.WorkTasks.RemoveRange(tasks);
                await _db.SaveChangesAsync();

                var msg = $"Berhasil menghapus {tasks.Count} pekerjaan.";

                if (isAjax)
                {
                    return Json(new
                    {
                        success = true,
                        deleted = tasks.Count,
                        message = msg
                    });
                }

                SetTaskMessage(msg, "success");

                if (!string.IsNullOrWhiteSpace(picName))
                    return RedirectToAction(nameof(ByPic), new { pic = picName });

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "BulkDelete WorkTasks gagal");

                if (isAjax)
                    return Json(new { success = false, message = "Terjadi kesalahan saat menghapus data." });

                SetTaskMessage("Terjadi kesalahan saat menghapus data.", "error");
                return RedirectToAction(nameof(Index));
            }
        }

        // ============== TEMPLATE IMPORT WORKTASK ==============
        [HttpGet]
        public IActionResult TemplateImport()
        {
            EnsureEpplus();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Template");

            ws.Cells[1, 1].Value = "TEMPLATE OUTSTANDING PEKERJAAN UTILITIES";
            ws.Cells[1, 1, 1, 11].Merge = true;
            ws.Cells[1, 1, 1, 11].Style.Font.Bold = true;

            int c = 1;
            ws.Cells[2, c++].Value = "No";
            ws.Cells[2, c++].Value = "NamaPekerjaan*";
            ws.Cells[2, c++].Value = "StartDate (yyyy-MM-dd)";
            ws.Cells[2, c++].Value = "DueDate (yyyy-MM-dd)";
            ws.Cells[2, c++].Value = "NamaRequest";
            ws.Cells[2, c++].Value = "ProgressPercent (0-100)";
            ws.Cells[2, c++].Value = "Status (Pending/InProgress/Done)";
            ws.Cells[2, c++].Value = "KeteranganTindakLanjut";
            ws.Cells[2, c++].Value = "Prioritas (Low/Medium/High/Emergency)";
            ws.Cells[2, c++].Value = "NamaPIC";
            ws.Cells[2, c++].Value = "CompletedDate (yyyy-MM-dd)";

            ws.Cells[2, 1, 2, 11].Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            var fileName = $"template-outstanding-pekerjaan-{DateTime.Now:yyyyMMdd}.xlsx";

            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ============== IMPORT WORKTASK (EXCEL) ==============
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile? file)
        {
            EnsureEpplus();

            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx")
                return Json(new { success = false, message = "Gunakan file Excel (.xlsx)." });

            int imported = 0;
            int row = 3;
            var errors = new List<string>();

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

                if (!title.Contains("TEMPLATE OUTSTANDING PEKERJAAN", StringComparison.OrdinalIgnoreCase) ||
                    !hNama.Contains("NamaPekerjaan", StringComparison.OrdinalIgnoreCase))
                {
                    return Json(new
                    {
                        success = false,
                        message = "Template tidak dikenali sebagai template outstanding pekerjaan."
                    });
                }

                DateTime? ParseDate(string? s)
                {
                    if (string.IsNullOrWhiteSpace(s)) return null;
                    if (DateTime.TryParse(s, out var d)) return d;
                    return null;
                }

                while (true)
                {
                    var namaPekerjaan = ws.Cells[row, 2].GetValue<string>()?.Trim();
                    if (string.IsNullOrWhiteSpace(namaPekerjaan))
                        break;

                    string? startStr = ws.Cells[row, 3].GetValue<string>()?.Trim();
                    string? dueStr = ws.Cells[row, 4].GetValue<string>()?.Trim();
                    string? namaRequest = ws.Cells[row, 5].GetValue<string>()?.Trim();
                    string? progressStr = ws.Cells[row, 6].GetValue<string>()?.Trim();
                    string? statusStr = ws.Cells[row, 7].GetValue<string>()?.Trim();
                    string? keterangan = ws.Cells[row, 8].GetValue<string>()?.Trim();
                    string? prioritasStr = ws.Cells[row, 9].GetValue<string>()?.Trim();
                    string? namaPIC = ws.Cells[row, 10].GetValue<string>()?.Trim();
                    string? completedStr = ws.Cells[row, 11].GetValue<string>()?.Trim();

                    int progress = 0;
                    if (!string.IsNullOrWhiteSpace(progressStr))
                    {
                        int.TryParse(progressStr, out progress);
                        if (progress < 0) progress = 0;
                        if (progress > 100) progress = 100;
                    }

                    WorkTaskStatus statusVal = WorkTaskStatus.Pending;
                    if (!string.IsNullOrWhiteSpace(statusStr))
                    {
                        if (!Enum.TryParse<WorkTaskStatus>(statusStr.Replace(" ", ""), true, out statusVal))
                        {
                            statusVal = WorkTaskStatus.Pending;
                        }
                    }

                    WorkTaskPriority prioritasVal = WorkTaskPriority.Medium;
                    if (!string.IsNullOrWhiteSpace(prioritasStr))
                    {
                        if (!Enum.TryParse<WorkTaskPriority>(prioritasStr, true, out prioritasVal))
                        {
                            prioritasVal = WorkTaskPriority.Medium;
                        }
                    }

                    var task = new WorkTask
                    {
                        NamaPekerjaan = namaPekerjaan!,
                        StartDate = ParseDate(startStr),
                        DueDate = ParseDate(dueStr),
                        NamaRequest = namaRequest,
                        ProgressPercent = progress,
                        Status = statusVal,
                        KeteranganTindakLanjut = keterangan,
                        Prioritas = prioritasVal,
                        NamaPIC = namaPIC,
                        CompletedDate = ParseDate(completedStr)
                    };

                    SyncStatusAndProgress(task);

                    _db.WorkTasks.Add(task);
                    await UpsertJobNameAsync(task.NamaPekerjaan);

                    imported++;
                    row++;
                }

                await _db.SaveChangesAsync();

                if (imported == 0)
                    return Json(new { success = false, message = "Tidak ada baris yang berhasil diimpor." });

                return Json(new
                {
                    success = true,
                    imported,
                    message = $"Berhasil mengimpor {imported} pekerjaan."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Import WorkTasks gagal pada baris {Row}", row);
                var msg = $"Import gagal pada baris {row}: {ex.Message}";
                if (errors.Count > 0) msg += $" ({errors.Count} baris lain dilewati)";
                return Json(new { success = false, message = msg });
            }
        }

        // ============== EXPORT EXCEL WORKTASK ==============
        [HttpGet]
        public async Task<IActionResult> ExportExcel(
            string? pic,
            WorkTaskStatus? status,
            WorkTaskPriority? priority,
            string? q,
            bool onlyOutstanding = false)
        {
            EnsureEpplus();

            var tasks = _db.WorkTasks.AsQueryable();

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                tasks = tasks.Where(t =>
                    (t.NamaPekerjaan ?? "").Contains(q) ||
                    (t.NamaRequest ?? "").Contains(q) ||
                    (t.NamaPIC ?? "").Contains(q) ||
                    (t.KeteranganTindakLanjut ?? "").Contains(q));
            }

            if (!string.IsNullOrWhiteSpace(pic))
            {
                pic = pic.Trim();
                tasks = tasks.Where(t => (t.NamaPIC ?? "").Contains(pic));
            }

            if (status.HasValue)
            {
                tasks = tasks.Where(t => t.Status == status.Value);
            }

            if (priority.HasValue)
            {
                tasks = tasks.Where(t => t.Prioritas == priority.Value);
            }

            if (onlyOutstanding)
            {
                tasks = tasks.Where(t => t.Status != WorkTaskStatus.Done);
            }

            var data = await tasks
                .AsNoTracking()
                .OrderBy(t => t.Status)
                .ThenByDescending(t => t.Prioritas)
                .ThenBy(t => t.DueDate ?? t.StartDate)
                .ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Outstanding");

            ws.Cells[1, 1].Value = "OUTSTANDING PEKERJAAN UTILITIES";
            ws.Cells[1, 1, 1, 11].Merge = true;
            ws.Cells[1, 1, 1, 11].Style.Font.Bold = true;

            int c = 1;
            ws.Cells[2, c++].Value = "No";
            ws.Cells[2, c++].Value = "NamaPekerjaan";
            ws.Cells[2, c++].Value = "StartDate";
            ws.Cells[2, c++].Value = "DueDate";
            ws.Cells[2, c++].Value = "NamaRequest";
            ws.Cells[2, c++].Value = "ProgressPercent";
            ws.Cells[2, c++].Value = "Status";
            ws.Cells[2, c++].Value = "KeteranganTindakLanjut";
            ws.Cells[2, c++].Value = "Prioritas";
            ws.Cells[2, c++].Value = "NamaPIC";
            ws.Cells[2, c++].Value = "CompletedDate";

            ws.Cells[2, 1, 2, 11].Style.Font.Bold = true;

            int row = 3;
            int no = 1;
            foreach (var t in data)
            {
                c = 1;
                ws.Cells[row, c++].Value = no;
                ws.Cells[row, c++].Value = t.NamaPekerjaan;
                ws.Cells[row, c++].Value = t.StartDate?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = t.DueDate?.ToString("yyyy-MM-dd");
                ws.Cells[row, c++].Value = t.NamaRequest;
                ws.Cells[row, c++].Value = t.ProgressPercent;
                ws.Cells[row, c++].Value = t.Status.ToString();
                ws.Cells[row, c++].Value = t.KeteranganTindakLanjut;
                ws.Cells[row, c++].Value = t.Prioritas.ToString();
                ws.Cells[row, c++].Value = t.NamaPIC;
                ws.Cells[row, c++].Value = t.CompletedDate?.ToString("yyyy-MM-dd");

                row++;
                no++;
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            var fileName = $"outstanding-pekerjaan-{DateTime.Now:yyyyMMdd}.xlsx";

            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        // ============== VIEWMODEL DETAIL PER PIC ==============
        public class WorkTasksByPicVM
        {
            public string PicName { get; set; } = "";
            public List<WorkTask> Tasks { get; set; } = new();

            public int Total => Tasks.Count;
            public int Pending => Tasks.Count(t => t.Status == WorkTaskStatus.Pending);
            public int InProgress => Tasks.Count(t => t.Status == WorkTaskStatus.InProgress);
            public int Done => Tasks.Count(t => t.Status == WorkTaskStatus.Done);
        }

        // ============== HELPER: SYNC STATUS & PROGRESS ==============
        private static void SyncStatusAndProgress(WorkTask task)
        {
            if (task.ProgressPercent == 100 && task.Status != WorkTaskStatus.Done)
            {
                task.Status = WorkTaskStatus.Done;
            }
            else if (task.Status == WorkTaskStatus.Done && task.ProgressPercent < 100)
            {
                task.ProgressPercent = 100;
            }

            // Untuk sekarang, CompletedDate dikontrol manual oleh user/import.
            // Kalau mau auto-set saat status Done, bisa tambahkan logika di sini.
        }

        // ============== HELPER: UPSERT MASTER NAMA PEKERJAAN ==========
        private async Task UpsertJobNameAsync(string? nama)
        {
            if (string.IsNullOrWhiteSpace(nama)) return;
            var val = nama.Trim();
            if (val.Length == 0) return;

            try
            {
                var exists = await _db.WorkJobNames
                    .AnyAsync(x => x.NamaPekerjaan == val);

                if (!exists)
                {
                    _db.WorkJobNames.Add(new WorkJobName
                    {
                        NamaPekerjaan = val,
                        CreatedAt = DateTime.UtcNow
                    });

                    await _db.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Gagal upsert WorkJobName {Nama}", val);
            }
        }

        // ============== LOOKUP EMPLOYEE UNTUK SELECT2 NAMA PIC ==========
        [HttpGet]
        public async Task<IActionResult> EmployeeOptions(string? term)
        {
            var q = _db.Employees.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(term))
            {
                term = term.Trim();
                q = q.Where(e => (e.NamaLengkap ?? "").Contains(term));
            }

            var list = await q
                .OrderBy(e => e.NamaLengkap)
                .Take(20)
                .ToListAsync();

            var result = list.Select(e => new
            {
                id = e.NamaLengkap,
                text = $"{e.NamaLengkap}"
            });

            return Json(result);
        }
    }
}
