using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using Microsoft.AspNetCore.Authorization;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class JenisBarangsController : Controller
    {
        private readonly ApplicationDbContext _db;
        public JenisBarangsController(ApplicationDbContext db) => _db = db;

        public async Task<IActionResult> Index(string? q)
        {
            var qry = _db.JenisBarangs.AsQueryable();
            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                qry = qry.Where(x => x.Nama.Contains(q) || (x.Keterangan ?? "").Contains(q));
            }
            return View(await qry.OrderBy(x => x.Nama).ToListAsync());
        }

        public IActionResult Create() => View();

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(JenisBarang m)
        {
            if (!ModelState.IsValid) return View(m);
            _db.Add(m); await _db.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        public async Task<IActionResult> Edit(int id)
        {
            var m = await _db.JenisBarangs.FindAsync(id);
            if (m == null) return NotFound();
            return View(m);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, JenisBarang m)
        {
            if (id != m.Id) return BadRequest();
            if (!ModelState.IsValid) return View(m);
            _db.Update(m); await _db.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        public async Task<IActionResult> Delete(int id)
        {
            var m = await _db.JenisBarangs.FindAsync(id);
            if (m == null) return NotFound();
            return View(m);
        }

        [HttpPost, ActionName("Delete"), ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var m = await _db.JenisBarangs.FindAsync(id);
            if (m != null) { _db.JenisBarangs.Remove(m); await _db.SaveChangesAsync(); }
            return RedirectToAction(nameof(Index));
        }
    }
}
