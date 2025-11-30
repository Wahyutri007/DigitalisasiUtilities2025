using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Text.RegularExpressions;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class JenisDinasController : Controller
    {
        private readonly ApplicationDbContext _db;

        public JenisDinasController(ApplicationDbContext db)
        {
            _db = db;
        }

        // Helper untuk normalisasi nama (trim + rapikan spasi)
        private static string NormalizeNama(string? nama)
        {
            if (string.IsNullOrWhiteSpace(nama)) return string.Empty;
            var t = nama.Trim();
            t = Regex.Replace(t, @"\s+", " ");
            return t;
        }

        // ========== INDEX (HALAMAN PENUH) ==========
        public async Task<IActionResult> Index()
        {
            var list = await _db.JenisDinas
                .AsNoTracking()
                .OrderBy(j => j.Nama)
                .ToListAsync();

            return View(list); // Views/JenisDinas/Index.cshtml
        }

        // ========== CREATE (GET – MODAL) ==========
        [HttpGet]
        public IActionResult Create()
        {
            return PartialView("_CreateModal", new JenisDinas());
        }

        // ========== CREATE (POST – AJAX) ==========
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(JenisDinas model)
        {
            if (!ModelState.IsValid)
            {
                var errs = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Validasi gagal.",
                    errors = errs
                });
            }

            var namaNorm = NormalizeNama(model.Nama);
            if (string.IsNullOrEmpty(namaNorm))
            {
                return Json(new
                {
                    success = false,
                    message = "Nama jenis dinas wajib diisi."
                });
            }

            model.Nama = namaNorm;

            var namaLower = namaNorm.ToLowerInvariant();

            var exists = await _db.JenisDinas
                .AnyAsync(x =>
                    x.Nama != null &&
                    x.Nama.Trim().ToLower() == namaLower);

            if (exists)
            {
                return Json(new
                {
                    success = false,
                    message = "Nama jenis dinas sudah ada."
                });
            }

            _db.JenisDinas.Add(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Jenis dinas berhasil ditambahkan."
            });
        }

        // ========== EDIT (GET – MODAL) ==========
        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            var it = await _db.JenisDinas.FindAsync(id);
            if (it == null)
            {
                return Content("<div class='alert alert-danger'>Data tidak ditemukan.</div>");
            }

            return PartialView("_EditModal", it);
        }

        // ========== EDIT (POST – AJAX) ==========
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, JenisDinas model)
        {
            if (id != model.Id)
                return Json(new { success = false, message = "ID tidak cocok." });

            if (!ModelState.IsValid)
            {
                var errs = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToList();

                return Json(new
                {
                    success = false,
                    message = "Validasi gagal.",
                    errors = errs
                });
            }

            var namaNorm = NormalizeNama(model.Nama);
            if (string.IsNullOrEmpty(namaNorm))
            {
                return Json(new
                {
                    success = false,
                    message = "Nama jenis dinas wajib diisi."
                });
            }

            model.Nama = namaNorm;
            var namaLower = namaNorm.ToLowerInvariant();

            var duplicate = await _db.JenisDinas
                .AnyAsync(x =>
                    x.Id != id &&
                    x.Nama != null &&
                    x.Nama.Trim().ToLower() == namaLower);

            if (duplicate)
            {
                return Json(new
                {
                    success = false,
                    message = "Nama jenis dinas sudah digunakan."
                });
            }

            _db.JenisDinas.Update(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Jenis dinas berhasil diperbarui."
            });
        }

        // ========== DELETE (POST – AJAX) ==========
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var it = await _db.JenisDinas.FindAsync(id);
            if (it == null)
                return Json(new { success = false, message = "Data tidak ditemukan." });

            // Cek relasi dengan Dinas
            var inUse = await _db.Dinas
                .AnyAsync(d => d.JenisDinasId == id);

            if (inUse)
            {
                return Json(new
                {
                    success = false,
                    message = "Tidak bisa dihapus karena sudah dipakai oleh data Dinas."
                });
            }

            _db.JenisDinas.Remove(it);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Jenis dinas berhasil dihapus."
            });
        }
    }
}
