using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Data;
using Microsoft.AspNetCore.Authorization;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class CutiLimitsController : Controller
    {
        private readonly ApplicationDbContext _db;
        public CutiLimitsController(ApplicationDbContext db) => _db = db;

        public async Task<IActionResult> Index(string? q)
        {
            var query = _db.Employees.AsNoTracking();

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(e =>
                    e.NamaLengkap.Contains(q) ||
                    e.NopegPersero.Contains(q) ||
                    (e.NopegKPI ?? "").Contains(q));
            }

            var items = await query
                .OrderBy(e => e.NamaLengkap)
                .Select(e => new VM
                {
                    Id = e.Id,
                    Nama = e.NamaLengkap,
                    Nopeg = e.NopegPersero,
                    Limit = e.CutiLimit
                })
                .ToListAsync();

            return View(items);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> BulkSet(int? limit, int[]? selectedIds /*, bool selectAll = false*/)
        {
            // Validasi dasar
            if (limit is null || limit < 0 || limit > 365)
            {
                SetSwal("warning", "Limit tidak valid", "Nilai harus 0–365 hari.", toast: true);
                TempData["Msg"] = "Limit harus antara 0 sampai 365 hari.";
                return RedirectToAction(nameof(Index));
            }

            // (Fitur applyToAll/selectAll sedang tidak dipakai pada view ini)
            if (selectedIds is not { Length: > 0 })
            {
                SetSwal("warning", "Tidak ada yang dipilih", "Centang minimal 1 karyawan terlebih dahulu.", toast: true);
                TempData["Msg"] = "⚠️ Tidak ada karyawan yang dipilih.";
                return RedirectToAction(nameof(Index));
            }

            // Update
            var selected = await _db.Employees.Where(e => selectedIds.Contains(e.Id)).ToListAsync();
            if (selected.Count == 0)
            {
                SetSwal("error", "Gagal menyimpan", "Data karyawan tidak ditemukan.", toast: false);
                TempData["Msg"] = "❌ Data karyawan tidak ditemukan.";
                return RedirectToAction(nameof(Index));
            }

            selected.ForEach(e => e.CutiLimit = limit);
            await _db.SaveChangesAsync();

            SetSwal("success", "Berhasil disimpan", $"Limit <b>{limit} hari</b> diterapkan ke <b>{selected.Count}</b> karyawan terpilih.", toast: false);
            TempData["Msg"] = $"✅ Limit {limit} hari diterapkan ke {selected.Count} karyawan terpilih.";
            return RedirectToAction(nameof(Index));
        }

        // Helper untuk SweetAlert via TempData
        private void SetSwal(string icon, string title, string html, bool toast)
        {
            TempData["SwalIcon"] = icon;          // success | error | warning | info | question
            TempData["SwalTitle"] = title;
            TempData["SwalText"] = html;          // boleh HTML pendek
            TempData["SwalToast"] = toast ? "1" : "0";
        }

        public class VM
        {
            public int Id { get; set; }
            public string Nama { get; set; } = "";
            public string Nopeg { get; set; } = "";
            public int? Limit { get; set; }
        }
    }
}
