using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using Microsoft.AspNetCore.Authorization;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class MasterSertifikasisController : Controller
    {
        private readonly ApplicationDbContext _db;
        public MasterSertifikasisController(ApplicationDbContext db) => _db = db;

        // ========================= INDEX =========================
        public async Task<IActionResult> Index(string? q)
        {
            var query = _db.MasterSertifikasis.AsQueryable();

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(m =>
                    (m.Nama ?? "").Contains(q) ||
                    (m.Kategori ?? "").Contains(q));
            }

            var data = await query
                .OrderBy(m => m.Nama)
                .ToListAsync();

            return View(data);
        }

        // ========================= CREATE ========================
        // GET: partial untuk popup
        public IActionResult Create()
        {
            return PartialView("Create", new MasterSertifikasi());
        }

        // POST: AJAX JSON
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(MasterSertifikasi model)
        {
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
                    errors
                });
            }

            _db.Add(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data master sertifikasi berhasil disimpan."
            });
        }

        // ========================== EDIT =========================
        // GET: partial untuk popup
        public async Task<IActionResult> Edit(int id)
        {
            var data = await _db.MasterSertifikasis.FindAsync(id);
            if (data == null) return NotFound();

            return PartialView("Edit", data);
        }

        // POST: AJAX JSON
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, MasterSertifikasi model)
        {
            if (id != model.Id) return BadRequest();

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
                    errors
                });
            }

            _db.Update(model);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data master sertifikasi berhasil diupdate."
            });
        }

        // ========================= DELETE ========================
        // Tidak perlu GET Delete view, kita hapus via AJAX dari tabel.

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var data = await _db.MasterSertifikasis.FindAsync(id);
            if (data == null)
            {
                return Json(new
                {
                    success = false,
                    message = "Data tidak ditemukan."
                });
            }

            _db.MasterSertifikasis.Remove(data);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = "Data master sertifikasi berhasil dihapus."
            });
        }
    }
}
