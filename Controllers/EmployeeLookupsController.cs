using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using Microsoft.AspNetCore.Authorization;

namespace UtilitiesHR.Controllers
{
    //[Authorize]
    public class EmployeeLookupsController : Controller
    {
        private readonly ApplicationDbContext _db;

        public EmployeeLookupsController(ApplicationDbContext db)
        {
            _db = db;
        }

        // ========== INDEX ==========

        public async Task<IActionResult> Index()
        {
            var items = await _db.EmployeeLookups
                .AsNoTracking()
                .OrderBy(x => x.Category)
                .ThenBy(x => x.Value)
                .ToListAsync();

            var categories = new[]
            {
                "Jabatan",
                "Pendidikan",
                "Jurusan",
                "Institusi",
                "StatusNikah",
                "UkWearpack",
                "UkKemejaPutih",
                "UkKemejaProduk",
                "UkSepatuSafety"
            };

            ViewBag.Categories = categories;

            return View(items);
        }

        // ========== ADD (AJAX + JSON) ==========

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Add(string category, string value)
        {
            if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(value))
            {
                return Json(new
                {
                    success = false,
                    message = "Kategori dan nilai wajib diisi."
                });
            }

            category = category.Trim();
            value = value.Trim();

            // Cek duplikat per kategori (case-insensitive kalau collate db-nya CI)
            bool exists = await _db.EmployeeLookups
                .AnyAsync(x => x.Category == category && x.Value == value);

            if (exists)
            {
                return Json(new
                {
                    success = false,
                    message = $"Nilai \"{value}\" sudah ada pada kategori {category}."
                });
            }

            _db.EmployeeLookups.Add(new EmployeeLookup
            {
                Category = category,
                Value = value
            });

            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = $"Berhasil menambah \"{value}\" ke kategori {category}."
            });
        }

        // ========== DELETE (AJAX + JSON) ==========

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var item = await _db.EmployeeLookups.FindAsync(id);
            if (item == null)
            {
                return Json(new
                {
                    success = false,
                    message = "Data tidak ditemukan."
                });
            }

            string name = item.Value;

            _db.EmployeeLookups.Remove(item);
            await _db.SaveChangesAsync();

            return Json(new
            {
                success = true,
                message = $"Berhasil menghapus \"{name}\"."
            });
        }
    }
}
