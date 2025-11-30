using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using UtilitiesHR.Data;
using UtilitiesHR.Models;
using Microsoft.AspNetCore.Authorization;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class MaterialsController : Controller
    {
        private readonly ApplicationDbContext _db;
        private static bool _epplusConfigured;

        public MaterialsController(ApplicationDbContext db)
        {
            _db = db;
        }

        // ================= EPPLUS LICENSE =================

        private static void ConfigureEpplusLicense()
        {
            if (_epplusConfigured)
                return;

            var env = Environment.GetEnvironmentVariable("EPPlusLicense");
            if (!string.IsNullOrWhiteSpace(env))
            {
                _epplusConfigured = true;
                return;
            }

            // Sesuaikan dengan versi EPPlus yang dipakai
            ExcelPackage.License.SetNonCommercialPersonal("UtilitiesHR Dev");
            _epplusConfigured = true;
        }

        // ================= HELPERS =================

        private async Task LoadJenisDropdown(int? selected = null)
        {
            ViewBag.JenisBarangId = new SelectList(
                await _db.JenisBarangs
                    .OrderBy(x => x.Nama)
                    .ToListAsync(),
                "Id", "Nama", selected
            );
        }

        private IQueryable<Material> GetMaterialsQuery(int? jenisBarangId, string? q)
        {
            var query = _db.Materials
                .Include(m => m.JenisBarang)
                .AsQueryable();

            if (jenisBarangId.HasValue && jenisBarangId.Value > 0)
                query = query.Where(m => m.JenisBarangId == jenisBarangId.Value);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(m =>
                    m.NamaBarang.Contains(q) ||
                    m.KodeBarang.Contains(q) ||
                    (m.PosisiBarang ?? "").Contains(q) ||
                    (m.JenisBarang != null && m.JenisBarang.Nama.Contains(q))
                );
            }

            return query;
        }

        private static List<MaterialTxnRow> CalculateRunningBalance(List<MaterialTxn> txns)
        {
            var result = new List<MaterialTxnRow>();
            var balance = 0;

            foreach (var t in txns.OrderBy(x => x.Tanggal).ThenBy(x => x.Id))
            {
                balance += (t.Jenis == JenisTransaksi.Masuk ? t.Jumlah : -t.Jumlah);

                result.Add(new MaterialTxnRow
                {
                    Id = t.Id,
                    Tanggal = t.Tanggal,
                    Jenis = t.Jenis,
                    Jumlah = t.Jumlah,
                    PenanggungJawab = t.PenanggungJawab,
                    Keterangan = t.Keterangan,
                    SisaStok = balance
                });
            }

            return result;
        }

        private static MaterialDetailsVM BuildMaterialDetailsVM(
            Material m,
            DateTime? startDate = null,
            DateTime? endDate = null)
        {
            var bulanIni = DateTime.Today.Month;
            var tahunIni = DateTime.Today.Year;

            var txBulanIni = m.Txns
                .Where(t => t.Tanggal.Month == bulanIni && t.Tanggal.Year == tahunIni)
                .ToList();

            var allWithBalance = CalculateRunningBalance(m.Txns.ToList());

            var filtered = allWithBalance.AsEnumerable();

            if (startDate.HasValue)
            {
                var s = startDate.Value.Date;
                filtered = filtered.Where(t => t.Tanggal.Date >= s);
            }

            if (endDate.HasValue)
            {
                var e = endDate.Value.Date;
                filtered = filtered.Where(t => t.Tanggal.Date <= e);
            }

            return new MaterialDetailsVM
            {
                Id = m.Id,
                KodeBarang = m.KodeBarang,
                NamaBarang = m.NamaBarang,
                JenisBarang = m.JenisBarang?.Nama ?? "-",
                PosisiBarang = m.PosisiBarang,
                Satuan = m.Satuan,
                Stok = m.JumlahBarang,
                MasukBulanIni = txBulanIni
                    .Where(t => t.Jenis == JenisTransaksi.Masuk)
                    .Sum(t => t.Jumlah),
                KeluarBulanIni = txBulanIni
                    .Where(t => t.Jenis == JenisTransaksi.Keluar)
                    .Sum(t => t.Jumlah),
                TotalMasuk = m.Txns
                    .Where(t => t.Jenis == JenisTransaksi.Masuk)
                    .Sum(t => t.Jumlah),
                TotalKeluar = m.Txns
                    .Where(t => t.Jenis == JenisTransaksi.Keluar)
                    .Sum(t => t.Jumlah),
                Txns = filtered
                    .OrderByDescending(x => x.Tanggal)
                    .ThenByDescending(x => x.Id)
                    .ToList()
            };
        }

        private static string EscapeCsv(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
                return "\"" + value.Replace("\"", "\"\"") + "\"";

            return value;
        }

        private static readonly string[] DateFormats =
        {
            "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"
        };

        // ================= INDEX =================

        public async Task<IActionResult> Index(int? jenisBarangId, string? q)
        {
            var list = await GetMaterialsQuery(jenisBarangId, q)
                .OrderBy(m => m.NamaBarang)
                .ToListAsync();

            return View(list);
        }

        // ================= MATERIAL CRUD (MODAL) =================

        public async Task<IActionResult> Create()
        {
            await LoadJenisDropdown();
            return PartialView("_CreateModal", new Material());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Material model)
        {
            // handle Select2 tag JenisBarang
            if (model.JenisBarangId == 0)
            {
                var input = Request.Form["JenisBarangId"].ToString();
                if (!string.IsNullOrWhiteSpace(input) && !int.TryParse(input, out _))
                {
                    var existing = await _db.JenisBarangs
                        .FirstOrDefaultAsync(x => x.Nama == input);
                    if (existing != null)
                        model.JenisBarangId = existing.Id;
                    else
                    {
                        var baru = new JenisBarang { Nama = input };
                        _db.JenisBarangs.Add(baru);
                        await _db.SaveChangesAsync();
                        model.JenisBarangId = baru.Id;
                    }
                }
            }

            if (!ModelState.IsValid)
            {
                await LoadJenisDropdown(model.JenisBarangId);
                return PartialView("_CreateModal", model);
            }

            var kodeExist = await _db.Materials
                .AnyAsync(x => x.KodeBarang == model.KodeBarang);
            if (kodeExist)
            {
                ModelState.AddModelError("KodeBarang", "Kode Barang sudah terdaftar.");
                await LoadJenisDropdown(model.JenisBarangId);
                return PartialView("_CreateModal", model);
            }

            model.JumlahBarang = 0;
            _db.Materials.Add(model);
            await _db.SaveChangesAsync();

            return Json(new { success = true, message = "Material berhasil ditambahkan." });
        }

        public async Task<IActionResult> Edit(int id)
        {
            var material = await _db.Materials
                .Include(m => m.JenisBarang)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (material == null)
                return Content("<div class='alert alert-danger'>Material tidak ditemukan.</div>");

            await LoadJenisDropdown(material.JenisBarangId);
            return PartialView("_EditModal", material);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Material model)
        {
            if (id != model.Id)
                return Json(new { success = false, message = "ID tidak valid." });

            if (model.JenisBarangId == 0)
            {
                var input = Request.Form["JenisBarangId"].ToString();
                if (!string.IsNullOrWhiteSpace(input) && !int.TryParse(input, out _))
                {
                    var existing = await _db.JenisBarangs
                        .FirstOrDefaultAsync(x => x.Nama == input);
                    if (existing != null)
                        model.JenisBarangId = existing.Id;
                    else
                    {
                        var baru = new JenisBarang { Nama = input };
                        _db.JenisBarangs.Add(baru);
                        await _db.SaveChangesAsync();
                        model.JenisBarangId = baru.Id;
                    }
                }
            }

            if (!ModelState.IsValid)
            {
                await LoadJenisDropdown(model.JenisBarangId);
                return PartialView("_EditModal", model);
            }

            var existingMaterial = await _db.Materials.FindAsync(id);
            if (existingMaterial == null)
                return Json(new { success = false, message = "Material tidak ditemukan." });

            var duplicateKode = await _db.Materials
                .AnyAsync(x => x.KodeBarang == model.KodeBarang && x.Id != id);
            if (duplicateKode)
            {
                ModelState.AddModelError("KodeBarang", "Kode Barang sudah terdaftar.");
                await LoadJenisDropdown(model.JenisBarangId);
                return PartialView("_EditModal", model);
            }

            existingMaterial.KodeBarang = model.KodeBarang;
            existingMaterial.NamaBarang = model.NamaBarang;
            existingMaterial.JenisBarangId = model.JenisBarangId;
            existingMaterial.PosisiBarang = model.PosisiBarang;
            existingMaterial.Satuan = model.Satuan;

            await _db.SaveChangesAsync();
            return Json(new { success = true, message = "Material berhasil diupdate." });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            try
            {
                var m = await _db.Materials
                    .Include(x => x.Txns)
                    .FirstOrDefaultAsync(x => x.Id == id);

                if (m == null)
                    return Json(new { success = false, message = "Material tidak ditemukan." });

                if (m.Txns.Count > 0)
                    return Json(new { success = false, message = "Tidak dapat menghapus material yang sudah memiliki transaksi." });

                _db.Materials.Remove(m);
                await _db.SaveChangesAsync();

                return Json(new { success = true, message = "Material berhasil dihapus." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // ================= DETAIL MATERIAL =================

        public async Task<IActionResult> Details(int id, DateTime? startDate, DateTime? endDate)
        {
            var m = await _db.Materials
                .Include(x => x.JenisBarang)
                .Include(x => x.Txns)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (m == null)
                return NotFound();

            var vm = BuildMaterialDetailsVM(m, startDate, endDate);

            ViewBag.StartDate = startDate?.ToString("yyyy-MM-dd");
            ViewBag.EndDate = endDate?.ToString("yyyy-MM-dd");

            return View(vm);
        }

        // ================= TRANSAKSI (MODAL) =================

        public async Task<IActionResult> TxnCreate(int materialId, int jenis = 1)
        {
            var m = await _db.Materials.FindAsync(materialId);
            if (m == null)
                return Content("<div class='alert alert-danger'>Material tidak ditemukan.</div>");

            ViewBag.Material = m;

            var dto = new TxnDto
            {
                MaterialId = materialId,
                Jenis = (JenisTransaksi)jenis,
                Tanggal = DateTime.Today
            };

            return PartialView("_TxnModal", dto);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> TxnCreate(TxnDto dto)
        {
            using var tr = await _db.Database.BeginTransactionAsync();
            try
            {
                var m = await _db.Materials.FindAsync(dto.MaterialId);
                if (m == null)
                    return Json(new { success = false, message = "Material tidak ditemukan." });

                if (!ModelState.IsValid)
                {
                    ViewBag.Material = m;
                    return PartialView("_TxnModal", dto);
                }

                if (dto.Jumlah <= 0)
                    return Json(new { success = false, message = "Jumlah harus lebih dari 0." });

                if (dto.Jenis == JenisTransaksi.Keluar && m.JumlahBarang < dto.Jumlah)
                    return Json(new { success = false, message = "Stok tidak mencukupi." });

                var txn = new MaterialTxn
                {
                    MaterialId = dto.MaterialId,
                    Jenis = dto.Jenis,
                    Tanggal = dto.Tanggal.Date,
                    Jumlah = dto.Jumlah,
                    PenanggungJawab = dto.PenanggungJawab,
                    Keterangan = dto.Keterangan
                };

                m.JumlahBarang += (dto.Jenis == JenisTransaksi.Masuk ? dto.Jumlah : -dto.Jumlah);

                _db.MaterialTxns.Add(txn);
                await _db.SaveChangesAsync();
                await tr.CommitAsync();

                return Json(new { success = true, message = "Transaksi berhasil disimpan." });
            }
            catch (Exception ex)
            {
                await tr.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        public async Task<IActionResult> EditTxn(int id)
        {
            var txn = await _db.MaterialTxns
                .Include(t => t.Material)
                .FirstOrDefaultAsync(t => t.Id == id);

            if (txn == null)
                return Content("<div class='alert alert-danger'>Transaksi tidak ditemukan.</div>");

            if (txn.Material == null)
                return Content("<div class='alert alert-danger'>Material terkait tidak ditemukan.</div>");

            ViewBag.Material = txn.Material;

            var dto = new TxnDto
            {
                Id = txn.Id,
                MaterialId = txn.MaterialId,
                Jenis = txn.Jenis,
                Tanggal = txn.Tanggal,
                Jumlah = txn.Jumlah,
                PenanggungJawab = txn.PenanggungJawab,
                Keterangan = txn.Keterangan
            };

            return PartialView("_EditTxnModal", dto);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> EditTxn(int id, TxnDto dto)
        {
            using var tr = await _db.Database.BeginTransactionAsync();
            try
            {
                if (id != dto.Id)
                    return Json(new { success = false, message = "ID transaksi tidak valid." });

                var txn = await _db.MaterialTxns
                    .Include(t => t.Material)
                    .FirstOrDefaultAsync(t => t.Id == id);

                if (txn == null || txn.Material == null)
                    return Json(new { success = false, message = "Transaksi / material tidak ditemukan." });

                var m = txn.Material;

                if (!ModelState.IsValid)
                {
                    ViewBag.Material = m;
                    return PartialView("_EditTxnModal", dto);
                }

                // rollback stok lama
                var oldDelta = (txn.Jenis == JenisTransaksi.Masuk ? txn.Jumlah : -txn.Jumlah);
                m.JumlahBarang -= oldDelta;

                if (dto.Jumlah <= 0)
                {
                    await tr.RollbackAsync();
                    return Json(new { success = false, message = "Jumlah harus lebih dari 0." });
                }

                if (dto.Jenis == JenisTransaksi.Keluar && m.JumlahBarang < dto.Jumlah)
                {
                    await tr.RollbackAsync();
                    return Json(new { success = false, message = "Stok tidak mencukupi." });
                }

                // apply baru
                txn.Tanggal = dto.Tanggal.Date;
                txn.Jenis = dto.Jenis;
                txn.Jumlah = dto.Jumlah;
                txn.PenanggungJawab = dto.PenanggungJawab;
                txn.Keterangan = dto.Keterangan;

                var newDelta = (dto.Jenis == JenisTransaksi.Masuk ? dto.Jumlah : -dto.Jumlah);
                m.JumlahBarang += newDelta;

                await _db.SaveChangesAsync();
                await tr.CommitAsync();

                return Json(new { success = true, message = "Transaksi berhasil diupdate." });
            }
            catch (Exception ex)
            {
                await tr.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteTxn(int id)
        {
            using var tr = await _db.Database.BeginTransactionAsync();
            try
            {
                var txn = await _db.MaterialTxns
                    .Include(t => t.Material)
                    .FirstOrDefaultAsync(t => t.Id == id);

                if (txn == null || txn.Material == null)
                    return Json(new { success = false, message = "Transaksi / material tidak ditemukan." });

                var m = txn.Material;
                var delta = (txn.Jenis == JenisTransaksi.Masuk ? txn.Jumlah : -txn.Jumlah);

                m.JumlahBarang -= delta;
                if (m.JumlahBarang < 0)
                    m.JumlahBarang = 0;

                _db.MaterialTxns.Remove(txn);
                await _db.SaveChangesAsync();
                await tr.CommitAsync();

                return Json(new { success = true, message = "Transaksi berhasil dihapus." });
            }
            catch (Exception ex)
            {
                await tr.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteTxnBulk([FromBody] List<int> ids)
        {
            if (ids == null || ids.Count == 0)
                return Json(new { success = false, message = "Tidak ada transaksi yang dipilih." });

            using var tr = await _db.Database.BeginTransactionAsync();
            try
            {
                var txns = await _db.MaterialTxns
                    .Include(t => t.Material)
                    .Where(t => ids.Contains(t.Id))
                    .ToListAsync();

                if (txns.Count == 0)
                    return Json(new { success = false, message = "Transaksi tidak ditemukan." });

                var material = txns.First().Material;
                if (material == null || txns.Any(t => t.MaterialId != material.Id))
                    return Json(new { success = false, message = "Bulk delete hanya untuk satu material yang sama." });

                foreach (var t in txns)
                {
                    var delta = (t.Jenis == JenisTransaksi.Masuk ? t.Jumlah : -t.Jumlah);
                    material.JumlahBarang -= delta;
                }

                if (material.JumlahBarang < 0)
                    material.JumlahBarang = 0;

                _db.MaterialTxns.RemoveRange(txns);
                await _db.SaveChangesAsync();
                await tr.CommitAsync();

                return Json(new { success = true, message = $"{txns.Count} transaksi berhasil dihapus." });
            }
            catch (Exception ex)
            {
                await tr.RollbackAsync();
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // ================= EXPORT MATERIAL & TRANSAKSI =================

        [HttpGet]
        public async Task<IActionResult> ExportMaterials(int? jenisBarangId, string? q)
        {
            ConfigureEpplusLicense();

            var items = await GetMaterialsQuery(jenisBarangId, q)
                .Include(m => m.JenisBarang)
                .OrderBy(m => m.NamaBarang)
                .Select(m => new
                {
                    m.KodeBarang,
                    m.NamaBarang,
                    JenisBarang = m.JenisBarang != null ? m.JenisBarang.Nama : "-",
                    m.PosisiBarang,
                    m.Satuan,
                    StokSaatIni = m.JumlahBarang
                })
                .ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Daftar Material");
            ws.Cells["A1"].LoadFromCollection(items, true);
            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            await package.SaveAsAsync(stream);
            stream.Position = 0;

            var fileName = $"DaftarMaterial-{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportMaterialsCsv(int? jenisBarangId, string? q)
        {
            var items = await GetMaterialsQuery(jenisBarangId, q)
                .Include(m => m.JenisBarang)
                .OrderBy(m => m.NamaBarang)
                .ToListAsync();

            var sb = new StringBuilder();
            sb.AppendLine("KodeBarang,NamaBarang,JenisBarang,PosisiBarang,Satuan,StokSaatIni");

            foreach (var m in items)
            {
                sb.AppendLine(string.Join(",",
                    EscapeCsv(m.KodeBarang),
                    EscapeCsv(m.NamaBarang),
                    EscapeCsv(m.JenisBarang?.Nama ?? "-"),
                    EscapeCsv(m.PosisiBarang),
                    EscapeCsv(m.Satuan),
                    m.JumlahBarang));
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var fileName = $"DaftarMaterial-{DateTime.Now:yyyyMMddHHmmss}.csv";
            return File(bytes, "text/csv", fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportTransactions(int id, DateTime? startDate, DateTime? endDate)
        {
            ConfigureEpplusLicense();

            var m = await _db.Materials
                .Include(x => x.JenisBarang)
                .Include(x => x.Txns)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (m == null)
                return NotFound();

            var vm = BuildMaterialDetailsVM(m, startDate, endDate);

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Riwayat Transaksi");

            ws.Cells["A1"].Value = $"Riwayat Transaksi: {vm.NamaBarang} ({vm.KodeBarang})";
            ws.Cells["A1:F1"].Merge = true;
            ws.Cells["A1"].Style.Font.Bold = true;

            var data = vm.Txns
                .OrderBy(t => t.Tanggal)
                .ThenBy(t => t.Id)
                .Select(t => new
                {
                    Tanggal = t.Tanggal.ToString("dd/MM/yyyy"),
                    Jenis = t.Jenis.ToString(),
                    Jumlah = (t.Jenis == JenisTransaksi.Masuk ? "+" : "-") + t.Jumlah,
                    t.SisaStok,
                    t.PenanggungJawab,
                    t.Keterangan
                })
                .ToList();

            ws.Cells["A3"].LoadFromCollection(data, true);
            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            await package.SaveAsAsync(stream);
            stream.Position = 0;

            var fileName = $"Riwayat-{vm.KodeBarang}-{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }

        [HttpGet]
        public async Task<IActionResult> ExportTransactionsCsv(int id, DateTime? startDate, DateTime? endDate)
        {
            var m = await _db.Materials
                .Include(x => x.Txns)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (m == null)
                return NotFound();

            var vm = BuildMaterialDetailsVM(m, startDate, endDate);

            var sb = new StringBuilder();
            sb.AppendLine("Tanggal,Jenis,Jumlah,SisaStok,PenanggungJawab,Keterangan");

            foreach (var t in vm.Txns.OrderBy(x => x.Tanggal).ThenBy(x => x.Id))
            {
                var jumlahStr = (t.Jenis == JenisTransaksi.Masuk ? "+" : "-") + t.Jumlah;
                sb.AppendLine(string.Join(",",
                    t.Tanggal.ToString("dd/MM/yyyy"),
                    t.Jenis,
                    jumlahStr,
                    t.SisaStok,
                    EscapeCsv(t.PenanggungJawab),
                    EscapeCsv(t.Keterangan)));
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var fileName = $"Riwayat-{vm.KodeBarang}-{DateTime.Now:yyyyMMddHHmmss}.csv";
            return File(bytes, "text/csv", fileName);
        }

        // ================= IMPORT MASTER + TRANSAKSI =================

        [HttpGet]
        public IActionResult DownloadImportTemplate()
        {
            ConfigureEpplusLicense();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("ImportMaterialTransaksi");

            ws.Cells[1, 1].Value = "KodeBarang";
            ws.Cells[1, 2].Value = "NamaBarang";
            ws.Cells[1, 3].Value = "JenisBarang";
            ws.Cells[1, 4].Value = "PosisiBarang";
            ws.Cells[1, 5].Value = "Satuan";
            ws.Cells[1, 6].Value = "Tanggal (dd/MM/yyyy)";
            ws.Cells[1, 7].Value = "Jenis (Masuk/Keluar)";
            ws.Cells[1, 8].Value = "Jumlah";
            ws.Cells[1, 9].Value = "PenanggungJawab";
            ws.Cells[1, 10].Value = "Keterangan";

            ws.Row(1).Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "TemplateImport_Material_Transaksi.xlsx");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ImportTransaksi(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            ConfigureEpplusLicense();

            var errors = new List<string>();
            var txns = new List<MaterialTxn>();
            var newJenis = new List<JenisBarang>();
            var newMaterials = new List<Material>();

            var jenisDict = (await _db.JenisBarangs.ToListAsync())
                .Where(j => !string.IsNullOrWhiteSpace(j.Nama))
                .ToDictionary(j => j.Nama.Trim(), StringComparer.OrdinalIgnoreCase);

            var materialDict = (await _db.Materials
                    .Include(m => m.JenisBarang)
                    .ToListAsync())
                .Where(m => !string.IsNullOrWhiteSpace(m.KodeBarang))
                .ToDictionary(m => m.KodeBarang.Trim(), StringComparer.OrdinalIgnoreCase);

            try
            {
                if (file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new StreamReader(file.OpenReadStream());
                    var header = await reader.ReadLineAsync();
                    if (header == null)
                        return Json(new { success = false, message = "File kosong." });

                    var row = 2;
                    while (!reader.EndOfStream)
                    {
                        var line = await reader.ReadLineAsync();
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            row++;
                            continue;
                        }

                        var cols = line.Split(',');
                        ProcessImportRowExtended(cols, row,
                            jenisDict, materialDict,
                            newJenis, newMaterials, txns, errors);
                        row++;
                    }
                }
                else if (file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    using var stream = file.OpenReadStream();
                    using var package = new ExcelPackage(stream);
                    var ws = package.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null)
                        return Json(new { success = false, message = "Sheet tidak ditemukan." });

                    var lastRow = ws.Dimension.End.Row;
                    for (var row = 2; row <= lastRow; row++)
                    {
                        var cols = new[]
                        {
                            ws.Cells[row, 1].Text,
                            ws.Cells[row, 2].Text,
                            ws.Cells[row, 3].Text,
                            ws.Cells[row, 4].Text,
                            ws.Cells[row, 5].Text,
                            ws.Cells[row, 6].Text,
                            ws.Cells[row, 7].Text,
                            ws.Cells[row, 8].Text,
                            ws.Cells[row, 9].Text,
                            ws.Cells[row,10].Text
                        };

                        if (cols.All(string.IsNullOrWhiteSpace))
                            continue;

                        ProcessImportRowExtended(cols, row,
                            jenisDict, materialDict,
                            newJenis, newMaterials, txns, errors);
                    }
                }
                else
                {
                    return Json(new
                    {
                        success = false,
                        message = "Format file tidak didukung (gunakan .xlsx atau .csv)."
                    });
                }

                if (errors.Count > 0)
                    return Json(new
                    {
                        success = false,
                        message = "Import gagal. Harap perbaiki error berikut:",
                        errors
                    });

                if (txns.Count == 0 && newMaterials.Count == 0 && newJenis.Count == 0)
                    return Json(new
                    {
                        success = false,
                        message = "Tidak ada baris valid untuk di-import."
                    });

                using var tr = await _db.Database.BeginTransactionAsync();

                if (newJenis.Count > 0)
                    _db.JenisBarangs.AddRange(newJenis);

                if (newMaterials.Count > 0)
                    _db.Materials.AddRange(newMaterials);

                if (txns.Count > 0)
                    _db.MaterialTxns.AddRange(txns);

                await _db.SaveChangesAsync();
                await tr.CommitAsync();

                var msg = $"Import selesai. {newJenis.Count} jenis baru, " +
                          $"{newMaterials.Count} material baru, " +
                          $"{txns.Count} transaksi.";

                return Json(new { success = true, message = msg });
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    message = $"Terjadi error saat memproses file: {ex.Message}"
                });
            }
        }

        // Core parser 1 baris
        private static void ProcessImportRowExtended(
            string[] cols,
            int row,
            Dictionary<string, JenisBarang> jenisDict,
            Dictionary<string, Material> materialDict,
            List<JenisBarang> newJenis,
            List<Material> newMaterials,
            List<MaterialTxn> txns,
            List<string> errors)
        {
            var kode = (cols.Length > 0 ? cols[0] : "")?.Trim();
            var nama = (cols.Length > 1 ? cols[1] : "")?.Trim();
            var jenisNama = (cols.Length > 2 ? cols[2] : "")?.Trim();
            var posisi = (cols.Length > 3 ? cols[3] : "")?.Trim();
            var satuan = (cols.Length > 4 ? cols[4] : "")?.Trim();
            var tglStr = (cols.Length > 5 ? cols[5] : "")?.Trim();
            var jenisTxnStr = (cols.Length > 6 ? cols[6] : "")?.Trim();
            var jmlStr = (cols.Length > 7 ? cols[7] : "")?.Trim();
            var pj = (cols.Length > 8 ? cols[8] : "")?.Trim();
            var ket = (cols.Length > 9 ? cols[9] : "")?.Trim();

            var allMasterEmpty =
                string.IsNullOrWhiteSpace(kode) &&
                string.IsNullOrWhiteSpace(nama) &&
                string.IsNullOrWhiteSpace(jenisNama) &&
                string.IsNullOrWhiteSpace(posisi) &&
                string.IsNullOrWhiteSpace(satuan);

            var allTxnEmpty =
                string.IsNullOrWhiteSpace(tglStr) &&
                string.IsNullOrWhiteSpace(jenisTxnStr) &&
                string.IsNullOrWhiteSpace(jmlStr) &&
                string.IsNullOrWhiteSpace(pj) &&
                string.IsNullOrWhiteSpace(ket);

            if (allMasterEmpty && allTxnEmpty)
                return;

            if (string.IsNullOrWhiteSpace(kode))
            {
                errors.Add($"Baris {row}: KodeBarang wajib diisi.");
                return;
            }

            // ===== Material (existing / baru) =====
            if (!materialDict.TryGetValue(kode, out var material))
            {
                if (string.IsNullOrWhiteSpace(nama))
                {
                    errors.Add($"Baris {row}: NamaBarang wajib diisi untuk KodeBarang baru '{kode}'.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(jenisNama))
                {
                    errors.Add($"Baris {row}: JenisBarang wajib diisi untuk KodeBarang baru '{kode}'.");
                    return;
                }

                if (!jenisDict.TryGetValue(jenisNama, out var jenisEntity))
                {
                    jenisEntity = new JenisBarang { Nama = jenisNama };
                    jenisDict[jenisNama] = jenisEntity;
                    newJenis.Add(jenisEntity);
                }

                material = new Material
                {
                    KodeBarang = kode,
                    NamaBarang = nama,
                    JenisBarang = jenisEntity,
                    PosisiBarang = posisi,
                    Satuan = satuan,
                    JumlahBarang = 0
                };

                materialDict[kode] = material;
                newMaterials.Add(material);
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(nama))
                    material.NamaBarang = nama;

                if (!string.IsNullOrWhiteSpace(jenisNama))
                {
                    if (!jenisDict.TryGetValue(jenisNama, out var jenisEntity))
                    {
                        jenisEntity = new JenisBarang { Nama = jenisNama };
                        jenisDict[jenisNama] = jenisEntity;
                        newJenis.Add(jenisEntity);
                    }

                    material.JenisBarang = jenisEntity;
                    material.JenisBarangId = jenisEntity.Id;
                }

                if (!string.IsNullOrWhiteSpace(posisi))
                    material.PosisiBarang = posisi;

                if (!string.IsNullOrWhiteSpace(satuan))
                    material.Satuan = satuan;
            }

            if (allTxnEmpty)
                return;

            // ===== Transaksi =====
            if (!DateTime.TryParseExact(
                    tglStr,
                    DateFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var tgl))
            {
                errors.Add($"Baris {row}: Tanggal tidak valid (gunakan format dd/MM/yyyy).");
                return;
            }

            JenisTransaksi jenisTxn;
            if (jenisTxnStr.Equals("Masuk", StringComparison.OrdinalIgnoreCase) ||
                jenisTxnStr.Equals("In", StringComparison.OrdinalIgnoreCase) ||
                jenisTxnStr.Equals("M", StringComparison.OrdinalIgnoreCase))
            {
                jenisTxn = JenisTransaksi.Masuk;
            }
            else if (jenisTxnStr.Equals("Keluar", StringComparison.OrdinalIgnoreCase) ||
                     jenisTxnStr.Equals("Out", StringComparison.OrdinalIgnoreCase) ||
                     jenisTxnStr.Equals("K", StringComparison.OrdinalIgnoreCase))
            {
                jenisTxn = JenisTransaksi.Keluar;
            }
            else
            {
                errors.Add($"Baris {row}: Jenis harus 'Masuk' atau 'Keluar'.");
                return;
            }

            if (!int.TryParse(jmlStr, out var jumlah) || jumlah <= 0)
            {
                errors.Add($"Baris {row}: Jumlah harus angka > 0.");
                return;
            }

            var delta = (jenisTxn == JenisTransaksi.Masuk ? jumlah : -jumlah);

            if (jenisTxn == JenisTransaksi.Keluar && material.JumlahBarang + delta < 0)
            {
                errors.Add($"Baris {row}: Stok tidak mencukupi untuk transaksi keluar (KodeBarang {kode}).");
                return;
            }

            material.JumlahBarang += delta;

            txns.Add(new MaterialTxn
            {
                Material = material,
                Tanggal = tgl.Date,
                Jenis = jenisTxn,
                Jumlah = jumlah,
                PenanggungJawab = pj,
                Keterangan = ket
            });
        }
    }
}
