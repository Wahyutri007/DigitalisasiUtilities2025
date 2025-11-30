using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using Microsoft.AspNetCore.Authorization;
namespace UtilitiesHR.Controllers
{
    // FIX: Menggunakan Primary Constructor
    [Authorize]
    public class MaterialJenisController(ApplicationDbContext db) : Controller
    {
        private readonly ApplicationDbContext _db = db;

        // GET: MaterialJenis
        public async Task<IActionResult> Index()
        {
            var jenisBarang = await _db.JenisBarangs
                .OrderBy(j => j.Nama)
                .ToListAsync();

            return View(jenisBarang);
        }

        // GET: MaterialJenis/Create (Modal)
        public IActionResult Create()
        {
            return PartialView("_JenisBarangCreateModal", new JenisBarang());
        }

        // POST: MaterialJenis/Create (Modal)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(JenisBarang jenisBarang)
        {
            using var transaction = await _db.Database.BeginTransactionAsync();
            try
            {
                if (ModelState.IsValid)
                {
                    if (await _db.JenisBarangs.CountAsync(j => j.Nama == jenisBarang.Nama) > 0)
                    {
                        return Json(new { success = false, message = "Nama jenis barang sudah ada" });
                    }

                    _db.Add(jenisBarang);
                    await _db.SaveChangesAsync();
                    await transaction.CommitAsync();

                    return Json(new { success = true, message = "Jenis barang berhasil ditambahkan" });
                }

                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .ToList();

                return Json(new { success = false, message = "Validasi gagal", errors });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // GET: MaterialJenis/Edit/5 (Modal)
        public async Task<IActionResult> Edit(int id)
        {
            var jenisBarang = await _db.JenisBarangs.FindAsync(id);
            if (jenisBarang == null)
            {
                return Content("<div class='alert alert-danger'>Jenis barang tidak ditemukan.</div>");
            }
            return PartialView("_JenisBarangEditModal", jenisBarang);
        }

        // POST: MaterialJenis/Edit/5 (Modal)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, JenisBarang jenisBarang)
        {
            using var transaction = await _db.Database.BeginTransactionAsync();
            try
            {
                if (id != jenisBarang.Id)
                {
                    return Json(new { success = false, message = "ID tidak valid" });
                }

                if (ModelState.IsValid)
                {
                    if (await _db.JenisBarangs.CountAsync(j => j.Nama == jenisBarang.Nama && j.Id != id) > 0)
                    {
                        return Json(new { success = false, message = "Nama jenis barang sudah ada" });
                    }

                    var existingJenis = await _db.JenisBarangs.FindAsync(id);
                    if (existingJenis == null)
                    {
                        return Json(new { success = false, message = "Jenis barang tidak ditemukan" });
                    }

                    existingJenis.Nama = jenisBarang.Nama;
                    existingJenis.Keterangan = jenisBarang.Keterangan;
                    await _db.SaveChangesAsync();
                    await transaction.CommitAsync();

                    return Json(new { success = true, message = "Jenis barang berhasil diupdate" });
                }

                var errors = ModelState.Values
                    .SelectMany(v => v.Errors)
                    .Select(e => e.ErrorMessage)
                    .ToList();

                return Json(new { success = false, message = "Validasi gagal", errors });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // POST: MaterialJenis/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            using var transaction = await _db.Database.BeginTransactionAsync();
            try
            {
                var jenisBarang = await _db.JenisBarangs.FindAsync(id);
                if (jenisBarang == null)
                {
                    return Json(new { success = false, message = "Jenis barang tidak ditemukan" });
                }

                var materialCount = await _db.Materials
                    .Where(m => m.JenisBarangId == id)
                    .CountAsync();

                if (materialCount > 0)
                {
                    return Json(new { success = false, message = "Tidak dapat menghapus jenis barang karena masih digunakan" });
                }

                _db.JenisBarangs.Remove(jenisBarang);
                await _db.SaveChangesAsync();
                await transaction.CommitAsync();

                return Json(new { success = true, message = "Jenis barang berhasil dihapus" });
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }
    }
}