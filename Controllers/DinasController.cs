using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.VisualBasic.FileIO; // CSV parser
using OfficeOpenXml;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class DinasController : Controller
    {
        private readonly ApplicationDbContext _db;
        public DinasController(ApplicationDbContext db) => _db = db;

        // ===================== EPPlus v8 =====================
        private static bool _epplusReady;
        private static void EnsureEpplus()
        {
            if (_epplusReady) return;
            ExcelPackage.License.SetNonCommercialPersonal("UtilitiesHR");
            _epplusReady = true;
        }

        // ===================== Helpers umum =====================
        private static string Norm(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var t = s.Trim();

            // buang penjelasan di dalam kurung pada header
            var idx = t.IndexOf('(');
            if (idx > 0) t = t[..idx];

            t = t.ToLowerInvariant();
            t = Regex.Replace(t, @"\s+", " ");
            t = Regex.Replace(t, @"[^a-z0-9]+", "");
            return t;
        }

        private static string Trunc(string? s, int max)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Trim();
            return s.Length <= max ? s : s[..max];
        }

        private static readonly string[] ExactFormats = new[]
        {
            "dd/MM/yyyy","d/M/yyyy","dd/MM/yy","d/M/yy",
            "yyyy-MM-dd","dd-MM-yyyy","d-M-yyyy","dd-MM-yy","d-M-yy",
            "MM/dd/yyyy","M/d/yyyy","MM/dd/yy","M/d/yy",
            "dd MMMM yyyy","d MMMM yyyy","dd MMM yyyy","d MMM yyyy",
            "MMMM dd, yyyy","MMM dd, yyyy",
            "dddd, dd MMMM yyyy","dddd, d MMMM yyyy",
            "dddd, dd MMM yyyy","dddd, d MMM yyyy"
        };

        private static readonly CultureInfo[] TryCultures = new[]
        {
            CultureInfo.GetCultureInfo("id-ID"),
            CultureInfo.GetCultureInfo("en-US")
        };

        private static string CleanDateString(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return s;
            var t = s.Trim();
            t = Regex.Replace(t, @"\b(\d{1,2})(st|nd|rd|th)\b", "$1", RegexOptions.IgnoreCase);
            var cut = t.IndexOf(" (", StringComparison.Ordinal);
            if (cut > 0) t = t[..cut];
            cut = t.IndexOf(" - ", StringComparison.Ordinal);
            if (cut > 0) t = t[..cut];
            cut = t.IndexOf(" – ", StringComparison.Ordinal);
            if (cut > 0) t = t[..cut];
            return t.Trim();
        }

        private static bool TryParseDateFlex(string? src, out DateTime dt)
        {
            dt = default;
            if (string.IsNullOrWhiteSpace(src)) return false;

            var s = CleanDateString(src);

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var oa))
            {
                try
                {
                    dt = DateTime.FromOADate(oa).Date;
                    return true;
                }
                catch { }
            }

            foreach (var c in TryCultures)
            {
                if (DateTime.TryParseExact(s, ExactFormats, c, DateTimeStyles.AllowWhiteSpaces, out dt))
                    return true;
            }

            foreach (var c in TryCultures)
            {
                if (DateTime.TryParse(s, c, DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    dt = dt.Date;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Pastikan JenisDinas ada. Kalau belum ada → buat baru.
        /// return Id atau null kalau gagal.
        /// </summary>
        private async Task<int?> EnsureJenisDinasFromName(string nama, string sumberKeterangan)
        {
            if (string.IsNullOrWhiteSpace(nama)) return null;

            // rapikan spasi & batasi panjang
            var trimmed = Trunc(nama, 120);
            if (string.IsNullOrWhiteSpace(trimmed)) return null;

            var namaNorm = Regex.Replace(trimmed, @"\s+", " ").Trim();
            var namaLower = namaNorm.ToLowerInvariant();

            // cari existing secara case-insensitive + trim
            var existing = await _db.JenisDinas
                .FirstOrDefaultAsync(j =>
                    j.Nama != null &&
                    j.Nama.Trim().ToLower() == namaLower);

            if (existing != null)
                return existing.Id;

            // belum ada → buat baru
            var jd = new JenisDinas
            {
                Nama = namaNorm,
                Keterangan = sumberKeterangan
            };

            _db.JenisDinas.Add(jd);
            await _db.SaveChangesAsync(); // kalau gagal, biar lempar exception dan import kebaca error
            return jd.Id;
        }

        // ===================== INDEX =====================
        public async Task<IActionResult> Index(int? employeeId, string? q, int? year)
        {
            var baseQuery = _db.Dinas
                .Include(d => d.Employee)
                .Include(d => d.JenisDinas)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(q))
            {
                var s = q.Trim();
                baseQuery = baseQuery.Where(d =>
                    d.Kegiatan.Contains(s) ||
                    d.Lokasi!.Contains(s) ||
                    (d.Employee != null &&
                     (d.Employee.NamaLengkap.Contains(s) || d.Employee.NopegPersero.Contains(s))));
            }

            var employees = await _db.Employees
                .OrderBy(e => e.NamaLengkap)
                .Select(e => new { e.Id, e.NamaLengkap })
                .ToListAsync();
            ViewBag.Employees = new SelectList(employees, "Id", "NamaLengkap", employeeId);

            var tahunList = await _db.Dinas
                .Select(d => d.TanggalBerangkat.Year)
                .Distinct().OrderByDescending(y => y).ToListAsync();
            ViewBag.TahunList = tahunList;

            int? selectedYear = (year.HasValue && tahunList.Contains(year.Value)) ? year.Value : (int?)null;
            ViewBag.SelectedYear = selectedYear;

            var query = baseQuery;
            if (employeeId.HasValue) query = query.Where(d => d.EmployeeId == employeeId.Value);
            if (selectedYear.HasValue) query = query.Where(d => d.TanggalBerangkat.Year == selectedYear.Value);

            var today = DateTime.Today;
            ViewBag.TotalTahun = await query.CountAsync();
            ViewBag.BulanIni = await query.Where(d => d.TanggalBerangkat.Year == today.Year && d.TanggalBerangkat.Month == today.Month).CountAsync();
            ViewBag.MasihBerlaku = await query.Where(d => d.TanggalBerangkat <= today && d.TanggalPulang >= today).CountAsync();
            ViewBag.SudahExpired = await query.Where(d => d.TanggalPulang < today).CountAsync();

            if (!employeeId.HasValue)
            {
                var flat = await query.AsNoTracking()
                    .Select(d => new
                    {
                        d.EmployeeId,
                        Nama = d.Employee != null ? d.Employee.NamaLengkap : null,
                        Nopeg = d.Employee != null ? d.Employee.NopegPersero : null
                    })
                    .ToListAsync();

                var groups = flat
                    .GroupBy(x => new
                    {
                        x.EmployeeId,
                        Nama = string.IsNullOrWhiteSpace(x.Nama) ? "(tanpa nama)" : x.Nama!,
                        Nopeg = string.IsNullOrWhiteSpace(x.Nopeg) ? "-" : x.Nopeg!
                    })
                    .Select(g => GroupRow.Create(g.Key.EmployeeId, g.Key.Nama, g.Key.Nopeg, g.Count()))
                    .OrderBy(x => x.Nama)
                    .ToList();

                ViewBag.GroupRows = groups;
            }

            var list = await query
                .OrderByDescending(d => d.TanggalBerangkat)
                .ThenByDescending(d => d.Id)
                .ToListAsync();

            ViewBag.IsDetailMode = employeeId.HasValue;
            return View(list);
        }

        // ===================== DETAILS =====================
        public async Task<IActionResult> Details(int id)
        {
            var dinas = await _db.Dinas
                .Include(d => d.Employee)
                .Include(d => d.JenisDinas)
                .FirstOrDefaultAsync(d => d.Id == id);
            if (dinas == null) return NotFound();
            return PartialView("Details", dinas);
        }

        // ===================== CREATE (GET) =====================
        public async Task<IActionResult> Create(int? employeeId)
        {
            ViewBag.EmployeeId = new SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", employeeId);

            ViewBag.JenisDinasId = new SelectList(
                await _db.JenisDinas.OrderBy(j => j.Nama).ToListAsync(),
                "Id", "Nama");

            return PartialView("Create", new Dinas
            {
                EmployeeId = employeeId ?? 0,
                TanggalBerangkat = DateTime.Today,
                TanggalPulang = DateTime.Today,
                Sifat = null
            });
        }

        // ===================== CREATE (POST) =====================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Dinas model, string? newJenisNama)
        {
            if (model.TanggalPulang < model.TanggalBerangkat)
                ModelState.AddModelError("TanggalPulang", "Tanggal pulang tidak boleh sebelum berangkat.");

            if (model.JenisDinasId == 0 && string.IsNullOrWhiteSpace(newJenisNama))
                ModelState.AddModelError("JenisDinasId", "Jenis dinas wajib diisi.");

            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
                return Json(new { success = false, message = "Data tidak valid.", errors });
            }

            if (model.JenisDinasId == 0 && !string.IsNullOrWhiteSpace(newJenisNama))
            {
                var jenisId = await EnsureJenisDinasFromName(newJenisNama, "Manual input dari form dinas");
                if (jenisId == null)
                    return Json(new { success = false, message = "Gagal menyimpan jenis dinas baru." });

                model.JenisDinasId = jenisId.Value;
            }

            _db.Dinas.Add(model);
            await _db.SaveChangesAsync();
            return Json(new { success = true, message = "Data dinas berhasil disimpan.", employeeId = model.EmployeeId });
        }

        // Quick add Jenis Dinas dari tombol (+)
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> CreateJenisDinas(string nama, string? keterangan)
        {
            if (string.IsNullOrWhiteSpace(nama))
                return Json(new { success = false, message = "Nama jenis dinas wajib diisi." });

            var jenisId = await EnsureJenisDinasFromName(nama, Trunc(keterangan, 400));
            if (jenisId == null)
            {
                return Json(new { success = false, message = "Gagal menyimpan jenis dinas baru." });
            }

            var list = await _db.JenisDinas
                .OrderBy(x => x.Nama)
                .Select(x => new { x.Id, x.Nama })
                .ToListAsync();

            return Json(new { success = true, id = jenisId.Value, items = list });
        }

        // ===================== EDIT (GET) =====================
        public async Task<IActionResult> Edit(int id)
        {
            var dinas = await _db.Dinas.FindAsync(id);
            if (dinas == null) return NotFound();

            ViewBag.EmployeeId = new SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", dinas.EmployeeId);

            ViewBag.JenisDinasId = new SelectList(
                await _db.JenisDinas.OrderBy(j => j.Nama).ToListAsync(),
                "Id", "Nama", dinas.JenisDinasId);

            return PartialView("Edit", dinas);
        }

        // ===================== EDIT (POST) =====================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, Dinas model, string? newJenisNama)
        {
            if (id != model.Id) return BadRequest();

            if (model.TanggalPulang < model.TanggalBerangkat)
                ModelState.AddModelError("TanggalPulang", "Tanggal pulang tidak boleh sebelum berangkat.");

            if (model.JenisDinasId == 0 && string.IsNullOrWhiteSpace(newJenisNama))
                ModelState.AddModelError("JenisDinasId", "Jenis dinas wajib diisi.");

            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).ToList();
                return Json(new { success = false, message = "Data tidak valid.", errors });
            }

            if (model.JenisDinasId == 0 && !string.IsNullOrWhiteSpace(newJenisNama))
            {
                var jenisId = await EnsureJenisDinasFromName(newJenisNama, "Manual input dari form dinas (edit)");
                if (jenisId == null)
                    return Json(new { success = false, message = "Gagal menyimpan jenis dinas baru." });

                model.JenisDinasId = jenisId.Value;
            }

            _db.Dinas.Update(model);
            await _db.SaveChangesAsync();
            return Json(new { success = true, message = "Data dinas berhasil diupdate.", employeeId = model.EmployeeId });
        }

        // ===================== DELETE =====================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(int id)
        {
            var dinas = await _db.Dinas.FindAsync(id);
            if (dinas == null) return Json(new { success = false, message = "Data tidak ditemukan." });

            _db.Dinas.Remove(dinas);
            await _db.SaveChangesAsync();
            return Json(new { success = true, message = "Data dinas berhasil dihapus.", employeeId = dinas.EmployeeId });
        }

        // ===================== IMPORT DIALOG =====================
        public IActionResult ImportDialog() => PartialView("ImportDialog");

        // ===================== IMPORT (Excel/CSV) =====================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File belum dipilih." });

            EnsureEpplus();

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".csv")
                return Json(new { success = false, message = "Format file tidak didukung. Gunakan .xlsx atau .csv." });

            var errors = new List<string>();
            var imported = 0;

            try
            {
                using var trx = await _db.Database.BeginTransactionAsync();

                if (ext == ".xlsx")
                {
                    using var ms = new MemoryStream();
                    await file.CopyToAsync(ms);
                    ms.Position = 0;

                    using var p = new ExcelPackage(ms);
                    var ws = p.Workbook.Worksheets.FirstOrDefault();
                    if (ws == null) return Json(new { success = false, message = "Sheet tidak ditemukan." });

                    int cMax = ws.Dimension.End.Column;
                    int rMax = ws.Dimension.End.Row;
                    var map = new Dictionary<string, int>();
                    for (int c = 1; c <= cMax; c++)
                    {
                        var h = Norm(ws.Cells[1, c].Text);
                        if (!string.IsNullOrEmpty(h) && !map.ContainsKey(h))
                            map[h] = c;
                    }

                    int Col(params string[] aliases)
                    {
                        foreach (var a in aliases)
                        {
                            var k = Norm(a);
                            if (map.TryGetValue(k, out var idx)) return idx;
                        }
                        return -1;
                    }

                    int colNopeg = Col("NopegPersero", "Nopeg", "Person Id", "PersonId");
                    int colNama = Col("NamaLengkap", "Nama");
                    int colJudul = Col("Judul Dinas", "Judul Dinas (Kegiatan)", "Judul", "Kegiatan", "Title");
                    int colLok = Col("Lokasi", "Lokasi (opsional)", "Location");
                    int colJenis = Col("JenisDinas", "Jenis Dinas", "Jenis");
                    int colSifat = Col("Sifat", "Sifat Dinas", "Sifat (PCU/Utilities, opsional)");
                    int colStart = Col("TanggalBerangkat", "Start Date", "StartDate", "Berangkat");
                    int colEnd = Col("TanggalPulang", "End Date", "EndDate", "Pulang");

                    if (colNama < 0 || colJenis < 0 || colStart < 0 || colEnd < 0)
                        return Json(new { success = false, message = "Header minimal: Nama, Jenis Dinas, Start Date, End Date." });

                    for (int r = 2; r <= rMax; r++)
                    {
                        string? get(int col) => col > 0 ? ws.Cells[r, col].Text?.Trim() : null;

                        var nopeg = get(colNopeg);
                        var nama = get(colNama);
                        var jenis = get(colJenis);
                        var sifat = get(colSifat);
                        var t1 = get(colStart);
                        var t2 = get(colEnd);
                        var lokasi = get(colLok);
                        var keg = get(colJudul);

                        if (new[] { nopeg, nama, jenis, sifat, t1, t2, lokasi, keg }.All(string.IsNullOrWhiteSpace))
                            continue;

                        if (await ProcessRowDynamic(r, nopeg, nama, jenis, sifat, t1, t2, keg, lokasi, errors))
                            imported++;
                    }
                }
                else // CSV
                {
                    using var sr = new StreamReader(file.OpenReadStream());
                    using var parser = new TextFieldParser(sr);
                    parser.SetDelimiters(",");
                    parser.HasFieldsEnclosedInQuotes = true;

                    if (parser.EndOfData)
                        return Json(new { success = false, message = "File kosong." });

                    var headers = parser.ReadFields() ?? Array.Empty<string>();
                    var map = new Dictionary<string, int>();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var h = Norm(headers[i]);
                        if (!string.IsNullOrEmpty(h) && !map.ContainsKey(h))
                            map[h] = i;
                    }

                    int Idx(params string[] aliases)
                    {
                        foreach (var a in aliases)
                        {
                            var k = Norm(a);
                            if (map.TryGetValue(k, out var id)) return id;
                        }
                        return -1;
                    }

                    int iNopeg = Idx("NopegPersero", "Nopeg", "Person Id", "PersonId");
                    int iNama = Idx("NamaLengkap", "Nama");
                    int iJudul = Idx("Judul Dinas", "Judul Dinas (Kegiatan)", "Judul", "Kegiatan", "Title");
                    int iLok = Idx("Lokasi", "Lokasi (opsional)", "Location");
                    int iJenis = Idx("JenisDinas", "Jenis Dinas", "Jenis");
                    int iSifat = Idx("Sifat", "Sifat Dinas", "Sifat (PCU/Utilities, opsional)");
                    int iStart = Idx("TanggalBerangkat", "Start Date", "StartDate", "Berangkat");
                    int iEnd = Idx("TanggalPulang", "End Date", "EndDate", "Pulang");

                    if (iNama < 0 || iJenis < 0 || iStart < 0 || iEnd < 0)
                        return Json(new { success = false, message = "Header minimal: Nama, Jenis Dinas, Start Date, End Date." });

                    int row = 2;
                    while (!parser.EndOfData)
                    {
                        var fields = parser.ReadFields() ?? Array.Empty<string>();
                        string? get(int i) => i >= 0 && i < fields.Length ? fields[i]?.Trim() : null;

                        var nopeg = get(iNopeg);
                        var nama = get(iNama);
                        var jenis = get(iJenis);
                        var sifat = get(iSifat);
                        var t1 = get(iStart);
                        var t2 = get(iEnd);
                        var lokasi = get(iLok);
                        var keg = get(iJudul);

                        if (await ProcessRowDynamic(row, nopeg, nama, jenis, sifat, t1, t2, keg, lokasi, errors))
                            imported++;

                        row++;
                    }
                }

                await trx.CommitAsync();

                if (imported == 0)
                {
                    var msg0 = errors.Count > 0
                        ? "Tidak ada baris yang berhasil diimpor."
                        : "Tidak ada data yang dapat diimpor.";
                    return Json(new { success = false, message = msg0 });
                }

                var msg = $"Berhasil mengimpor {imported} baris.";
                if (errors.Count > 0)
                    msg += $" {errors.Count} baris dilewati karena tidak valid.";

                return Json(new { success = true, message = msg });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        // ===================== Helper Import =====================
        private static bool TryParseSifatNullable(string? s, out SifatDinas? result, out string? error)
        {
            result = null; error = null;
            if (string.IsNullOrWhiteSpace(s)) return true;
            var t = s.Trim().Replace(".", "").Replace(" ", "").ToLowerInvariant();
            if (t == "pcu")
            {
                result = SifatDinas.PCU; return true;
            }
            if (t.StartsWith("util"))
            {
                result = SifatDinas.Utilities; return true;
            }
            error = $"Sifat tidak valid: '{s}'. Gunakan PCU / Utilities atau kosongkan.";
            return false;
        }

        private async Task<bool> ProcessRowDynamic(
            int row,
            string? nopeg,
            string? nama,
            string? jenis,
            string? sifat,
            string? startText,
            string? endText,
            string? kegiatan,
            string? lokasi,
            List<string> errors)
        {
            if (string.IsNullOrWhiteSpace(nama)) return false;
            if (string.IsNullOrWhiteSpace(jenis)) return false;

            if (!TryParseSifatNullable(sifat, out var sVal, out _))
                return false;

            if (!TryParseDateFlex(startText, out var tBrgkt))
                return false;
            if (!TryParseDateFlex(endText, out var tPlg))
                return false;

            var normNama = string.IsNullOrWhiteSpace(nama) ? null : nama.Trim();
            var normNopeg = string.IsNullOrWhiteSpace(nopeg) ? null : nopeg.Trim();

            if (string.IsNullOrEmpty(normNama) || string.IsNullOrEmpty(normNopeg))
            {
                errors.Add($"Baris {row}: Nama dan Person Id (Nopeg) wajib diisi dan harus cocok dengan master karyawan.");
                return false;
            }

            var nopegLower = normNopeg.ToLowerInvariant();

            var emp = await _db.Employees.FirstOrDefaultAsync(e =>
                e.NopegPersero != null &&
                e.NopegPersero.Trim().ToLower() == nopegLower);

            if (emp == null)
            {
                errors.Add($"Baris {row}: Tidak ditemukan karyawan dengan Nopeg '{normNopeg}' di master karyawan.");
                return false;
            }

            var empNamaNorm = (emp.NamaLengkap ?? string.Empty).Trim();
            if (!empNamaNorm.Equals(normNama, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"Baris {row}: Nama di file ('{normNama}') tidak cocok dengan master karyawan ('{empNamaNorm}') untuk Nopeg '{normNopeg}'.");
                return false;
            }

            var jenisNorm = jenis!.Trim();
            var jenisId = await EnsureJenisDinasFromName(jenisNorm, "Auto from import");
            if (jenisId == null)
            {
                errors.Add($"Baris {row}: Gagal menyimpan jenis dinas '{jenisNorm}'.");
                return false;
            }

            var dinas = new Dinas
            {
                EmployeeId = emp.Id,
                JenisDinasId = jenisId.Value,
                Sifat = sVal,
                Lokasi = Trunc(lokasi, 200),
                TanggalBerangkat = tBrgkt,
                TanggalPulang = tPlg,
                Kegiatan = Trunc(string.IsNullOrWhiteSpace(kegiatan) ? jenisNorm : kegiatan!, 500)
            };

            _db.Dinas.Add(dinas);
            try
            {
                await _db.SaveChangesAsync();
                return true;
            }
            catch (DbUpdateException)
            {
                return false;
            }
        }

        // ===================== EXPORT CSV =====================
        [HttpGet]
        public async Task<IActionResult> ExportCsv(int? employeeId, int? year, string? q)
        {
            var query = _db.Dinas
                .Include(d => d.Employee)
                .Include(d => d.JenisDinas)
                .AsQueryable();

            if (employeeId.HasValue) query = query.Where(d => d.EmployeeId == employeeId.Value);
            if (year.HasValue) query = query.Where(d => d.TanggalBerangkat.Year == year.Value);
            if (!string.IsNullOrWhiteSpace(q))
            {
                var s = q.Trim();
                query = query.Where(d =>
                    d.Kegiatan.Contains(s) ||
                    d.Lokasi!.Contains(s) ||
                    (d.Employee != null &&
                     (d.Employee.NamaLengkap.Contains(s) || d.Employee.NopegPersero.Contains(s))));
            }

            var list = await query.OrderByDescending(d => d.TanggalBerangkat).ToListAsync();

            var sb = new StringBuilder();
            sb.Append('\uFEFF');
            sb.AppendLine("Nama,Nopeg,Judul,Lokasi,Jenis,Sifat,Berangkat,Pulang,Hari,Kegiatan");
            foreach (var x in list)
            {
                string esc(string? v) => string.IsNullOrEmpty(v)
                    ? ""
                    : (v.Contains(',') || v.Contains('"') || v.Contains('\n'))
                        ? $"\"{v.Replace("\"", "\"\"")}\""
                        : v;

                sb.AppendLine(string.Join(",",
                    esc(x.Employee?.NamaLengkap),
                    esc(x.Employee?.NopegPersero),
                    esc(x.Kegiatan),
                    esc(x.Lokasi),
                    esc(x.JenisDinas?.Nama),
                    esc(x.Sifat?.ToString() ?? ""),
                    x.TanggalBerangkat.ToString("yyyy-MM-dd"),
                    x.TanggalPulang.ToString("yyyy-MM-dd"),
                    x.LamaHari.ToString(),
                    esc(x.Kegiatan)
                ));
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            return File(bytes, "text/csv", $"Dinas-{DateTime.Now:yyyyMMddHHmmss}.csv");
        }

        // ===================== EXPORT EXCEL =====================
        [HttpGet]
        public async Task<IActionResult> ExportExcel(int? employeeId, int? year, string? q)
        {
            EnsureEpplus();

            var query = _db.Dinas
                .Include(d => d.Employee)
                .Include(d => d.JenisDinas)
                .AsQueryable();

            if (employeeId.HasValue) query = query.Where(d => d.EmployeeId == employeeId.Value);
            if (year.HasValue) query = query.Where(d => d.TanggalBerangkat.Year == year.Value);
            if (!string.IsNullOrWhiteSpace(q))
            {
                var s = q.Trim();
                query = query.Where(d =>
                    d.Kegiatan.Contains(s) ||
                    d.Lokasi!.Contains(s) ||
                    (d.Employee != null &&
                     (d.Employee.NamaLengkap.Contains(s) || d.Employee.NopegPersero.Contains(s))));
            }

            var list = await query.OrderByDescending(d => d.TanggalBerangkat).ToListAsync();

            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Dinas");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Nopeg";
            ws.Cells[1, 3].Value = "Judul Dinas (Kegiatan)";
            ws.Cells[1, 4].Value = "Lokasi";
            ws.Cells[1, 5].Value = "Jenis Dinas";
            ws.Cells[1, 6].Value = "Sifat";
            ws.Cells[1, 7].Value = "Start Date";
            ws.Cells[1, 8].Value = "End Date";
            ws.Cells[1, 9].Value = "Hari";

            ws.Row(1).Style.Font.Bold = true;

            int r = 2;
            foreach (var x in list)
            {
                ws.Cells[r, 1].Value = x.Employee?.NamaLengkap;
                ws.Cells[r, 2].Value = x.Employee?.NopegPersero;
                ws.Cells[r, 3].Value = x.Kegiatan;
                ws.Cells[r, 4].Value = x.Lokasi;
                ws.Cells[r, 5].Value = x.JenisDinas?.Nama;
                ws.Cells[r, 6].Value = x.Sifat?.ToString() ?? "-";
                ws.Cells[r, 7].Value = x.TanggalBerangkat.ToString("yyyy-MM-dd");
                ws.Cells[r, 8].Value = x.TanggalPulang.ToString("yyyy-MM-dd");
                ws.Cells[r, 9].Value = x.LamaHari;
                r++;
            }
            ws.Cells.AutoFitColumns();

            var bytes = p.GetAsByteArray();
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Dinas-{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }

        // ===================== TEMPLATE IMPORT =====================
        public IActionResult DownloadImportTemplate()
        {
            EnsureEpplus();

            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("TemplateDinas");

            ws.Cells[1, 1].Value = "Nama";
            ws.Cells[1, 2].Value = "Person Id";
            ws.Cells[1, 3].Value = "Judul Dinas (Kegiatan)";
            ws.Cells[1, 4].Value = "Lokasi (opsional)";
            ws.Cells[1, 5].Value = "Jenis Dinas";
            ws.Cells[1, 6].Value = "Start Date (contoh: 19/05/2025)";
            ws.Cells[1, 7].Value = "End Date (contoh: 20/05/2025)";
            ws.Cells[1, 8].Value = "Sifat (PCU/Utilities, opsional)";

            ws.Row(1).Style.Font.Bold = true;
            ws.Cells.AutoFitColumns();

            var bytes = p.GetAsByteArray();
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Template_Import_Dinas.xlsx");
        }
    }
}
