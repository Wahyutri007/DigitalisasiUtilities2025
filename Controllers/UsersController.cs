using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    // Hanya Admin & SuperAdmin yang boleh kelola akun
    [Authorize(Roles = "Admin,SuperAdmin")]
    public class UsersController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly RoleManager<IdentityRole> _roleManager;

        private const string DefaultPassword = "Pertamina@2025";

        static UsersController()
        {
            ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR - PT KPI RU II");
        }

        public UsersController(
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager,
            RoleManager<IdentityRole> roleManager)
        {
            _db = db;
            _userManager = userManager;
            _roleManager = roleManager;
        }

        // ====================== HELPER SYNC ======================
        /// <summary>
        /// Kalau data User berubah, paksa Employee yang ter-link ikut update
        /// (NamaLengkap & EmailPertamina).
        /// </summary>
        private async Task SyncEmployeeFromUserAsync(ApplicationUser user)
        {
            if (user.EmployeeId is null) return;

            var emp = await _db.Employees.FirstOrDefaultAsync(e => e.Id == user.EmployeeId);
            if (emp == null) return;

            if (!string.IsNullOrWhiteSpace(user.FullName))
                emp.NamaLengkap = user.FullName;

            if (!string.IsNullOrWhiteSpace(user.Email))
                emp.EmailPertamina = user.Email;

            await _db.SaveChangesAsync();
        }

        private static List<string> CollectModelErrors(ModelStateDictionary modelState)
        {
            return modelState.Values
                .SelectMany(v => v.Errors)
                .Select(e => e.ErrorMessage)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }

        private static bool IsAjaxRequest(HttpRequest request)
        {
            return string.Equals(
                request.Headers["X-Requested-With"].ToString(),
                "XMLHttpRequest",
                StringComparison.OrdinalIgnoreCase);
        }

        // =========================================================
        // LIST USER
        // =========================================================
        public async Task<IActionResult> Index(string? q, int? employeeId)
        {
            ViewBag.Employees = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", employeeId
            );

            var query = _userManager.Users
                .Include(u => u.Employee)
                .AsQueryable();

            if (employeeId is not null)
                query = query.Where(u => u.EmployeeId == employeeId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                query = query.Where(u =>
                    (u.Email ?? "").Contains(q) ||
                    (u.FullName ?? "").Contains(q) ||
                    (u.Employee!.NopegPersero ?? "").Contains(q));
            }

            var data = await query
                .OrderBy(u => u.FullName ?? u.Email)
                .ToListAsync();

            return View(data);
        }

        // =========================================================
        // REGISTER MANUAL
        // =========================================================
        [HttpGet]
        public async Task<IActionResult> Create()
        {
            await LoadEmployeesForCreate();
            return View(new CreateUserVM());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(CreateUserVM model)
        {
            if (!ModelState.IsValid)
            {
                await LoadEmployeesForCreate(model.EmployeeId);
                return View(model);
            }

            var existing = await _userManager.FindByEmailAsync(model.Email);
            if (existing != null)
            {
                ModelState.AddModelError("Email", "Email sudah digunakan oleh akun lain.");
                await LoadEmployeesForCreate(model.EmployeeId);
                return View(model);
            }

            var user = new ApplicationUser
            {
                Email = model.Email,
                UserName = model.Email,
                FullName = model.FullName,
                EmployeeId = model.EmployeeId,
                MustChangePassword = true,   // << WAJIB GANTI PASSWORD
                LockoutEnabled = true        // << lockout aktif
            };

            var result = await _userManager.CreateAsync(user, DefaultPassword);
            if (!result.Succeeded)
            {
                foreach (var e in result.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);

                await LoadEmployeesForCreate(model.EmployeeId);
                return View(model);
            }

            // role default = User
            if (!await _roleManager.RoleExistsAsync("User"))
                await _roleManager.CreateAsync(new IdentityRole("User"));

            await _userManager.AddToRoleAsync(user, "User");

            // === sinkron ke Employee kalau ada relasi ===
            await SyncEmployeeFromUserAsync(user);

            TempData["UserMessage"] = "Akun berhasil dibuat. Password default: " + DefaultPassword;
            return RedirectToAction(nameof(Index));
        }

        private async Task LoadEmployeesForCreate(int? selectedId = null)
        {
            ViewBag.Employees = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(
                await _db.Employees.OrderBy(e => e.NamaLengkap).ToListAsync(),
                "Id", "NamaLengkap", selectedId
            );
        }

        public class CreateUserVM
        {
            public int? EmployeeId { get; set; }
            public string Email { get; set; } = "";
            public string FullName { get; set; } = "";
        }

        // =========================================================
        // EDIT USER + GANTI ROLE
        // =========================================================
        [HttpGet]
        public async Task<IActionResult> Edit(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return NotFound();

            var user = await _userManager.Users
                .Include(u => u.Employee)
                .FirstOrDefaultAsync(u => u.Id == id);

            if (user == null) return NotFound();

            await LoadEmployeesForCreate(user.EmployeeId);

            // list role yang tersedia
            var allRoles = await _roleManager.Roles
                .Select(r => r.Name!)
                .OrderBy(n => n)
                .ToListAsync();

            // role aktif user (ambil satu saja, kalau lebih dari satu)
            var userRoles = await _userManager.GetRolesAsync(user);
            var selectedRole = userRoles.FirstOrDefault() ?? "User";

            ViewBag.Roles = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(
                allRoles,
                selectedRole
            );

            var vm = new EditUserVM
            {
                Id = user.Id,
                Email = user.Email ?? "",
                FullName = user.FullName ?? "",
                EmployeeId = user.EmployeeId,
                SelectedRole = selectedRole
            };

            // Dipakai di modal, Layout = null di view
            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(EditUserVM model)
        {
            bool isAjax = IsAjaxRequest(Request);

            async Task<IActionResult> ReturnValidationError(string message)
            {
                var errors = CollectModelErrors(ModelState);

                if (isAjax)
                {
                    return Json(new
                    {
                        success = false,
                        message,
                        errors
                    });
                }

                await LoadEmployeesForCreate(model.EmployeeId);
                var allRolesLocal = _roleManager.Roles.Select(r => r.Name!).OrderBy(n => n).ToList();
                ViewBag.Roles = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(
                    allRolesLocal, model.SelectedRole
                );
                return View(model);
            }

            if (!ModelState.IsValid)
            {
                return await ReturnValidationError("Validasi gagal. Periksa kembali input.");
            }

            var user = await _userManager.FindByIdAsync(model.Id);
            if (user == null)
            {
                if (isAjax)
                    return Json(new { success = false, message = "User tidak ditemukan." });

                return NotFound();
            }

            // Cek email bentrok dengan user lain
            var existing = await _userManager.FindByEmailAsync(model.Email);
            if (existing != null && existing.Id != user.Id)
            {
                ModelState.AddModelError("Email", "Email sudah digunakan oleh akun lain.");
                return await ReturnValidationError("Email sudah digunakan oleh akun lain.");
            }

            // ==== update basic data ====
            user.FullName = model.FullName;
            user.Email = model.Email;
            user.UserName = model.Email;
            user.EmployeeId = model.EmployeeId;

            var updateResult = await _userManager.UpdateAsync(user);
            if (!updateResult.Succeeded)
            {
                foreach (var e in updateResult.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);

                return await ReturnValidationError("Gagal menyimpan perubahan akun.");
            }

            // ====== GANTI PASSWORD (OPSIONAL) ======
            if (!string.IsNullOrWhiteSpace(model.NewPassword))
            {
                if (model.NewPassword != model.ConfirmPassword)
                {
                    ModelState.AddModelError("NewPassword", "Password baru dan konfirmasi tidak sama.");
                    return await ReturnValidationError("Password baru dan konfirmasi tidak sama.");
                }

                // ❌ tidak boleh sama dengan password default
                if (model.NewPassword == DefaultPassword)
                {
                    ModelState.AddModelError("NewPassword", "Password baru tidak boleh sama dengan password default sistem.");
                    return await ReturnValidationError("Password baru tidak boleh sama dengan password default sistem.");
                }

                var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
                var passResult = await _userManager.ResetPasswordAsync(user, resetToken, model.NewPassword);

                if (!passResult.Succeeded)
                {
                    foreach (var e in passResult.Errors)
                        ModelState.AddModelError(string.Empty, e.Description);

                    return await ReturnValidationError("Gagal mengganti password.");
                }

                // Admin ganti password manual → anggap password sudah "baru"
                user.MustChangePassword = false;
                user.LastPasswordChangedAt = DateTimeOffset.UtcNow;
                await _userManager.UpdateAsync(user);
            }

            // ================== UPDATE ROLE ==================
            if (!string.IsNullOrWhiteSpace(model.SelectedRole))
            {
                // pastikan role ada
                if (!await _roleManager.RoleExistsAsync(model.SelectedRole))
                {
                    await _roleManager.CreateAsync(new IdentityRole(model.SelectedRole));
                }

                var currentRoles = await _userManager.GetRolesAsync(user);
                // hapus semua dulu supaya 1 user 1 role utama
                if (currentRoles.Any())
                    await _userManager.RemoveFromRolesAsync(user, currentRoles);

                await _userManager.AddToRoleAsync(user, model.SelectedRole);
            }

            // === sinkron balik ke Employee ===
            await SyncEmployeeFromUserAsync(user);

            if (isAjax)
            {
                return Json(new
                {
                    success = true,
                    message = "Data akun berhasil diupdate."
                });
            }

            TempData["UserMessage"] = "Data akun berhasil diupdate.";
            return RedirectToAction(nameof(Index));
        }

        public class EditUserVM
        {
            public string Id { get; set; } = "";
            public int? EmployeeId { get; set; }
            public string Email { get; set; } = "";
            public string FullName { get; set; } = "";
            public string SelectedRole { get; set; } = "User";

            // field untuk ganti password
            public string? NewPassword { get; set; }
            public string? ConfirmPassword { get; set; }
        }

        // =========================================================
        // RESET PASSWORD
        // =========================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ResetPassword(string id)
        {
            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
                return Json(new { success = false, message = "User tidak ditemukan." });

            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var result = await _userManager.ResetPasswordAsync(user, token, DefaultPassword);

            if (!result.Succeeded)
            {
                var msg = string.Join("; ", result.Errors.Select(e => e.Description));
                return Json(new { success = false, message = msg });
            }

            // Reset ke default → paksa user ganti lagi setelah login
            user.MustChangePassword = true;
            user.LastPasswordChangedAt = null; // akan diisi saat user benar2 ubah dari default
            await _userManager.UpdateAsync(user);

            return Json(new
            {
                success = true,
                message = $"Password berhasil direset ke {DefaultPassword}. User akan diminta mengganti password setelah login."
            });

        }

        // =========================================================
        // DELETE USER
        // =========================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Delete(string id)
        {
            // ambil user + employee terkait
            var user = await _userManager.Users
                .Include(u => u.Employee)
                .FirstOrDefaultAsync(u => u.Id == id);

            if (user == null)
                return Json(new { success = false, message = "User tidak ditemukan." });

            // ❗ KEBIJAKAN BARU:
            // tidak boleh hapus akun kalau masih terhubung ke Employee
            if (user.EmployeeId is int empId)
            {
                var empName = user.Employee?.NamaLengkap ?? "(tanpa nama)";
                var empNopeg = user.Employee?.NopegPersero ?? "-";

                return Json(new
                {
                    success = false,
                    requireDeleteEmployee = true,
                    message = $"Akun ini terhubung dengan data pekerja \"{empName}\" ({empNopeg}). " +
                              "Silakan hapus data pekerja tersebut dulu di menu Pekerja. " +
                              "Setelah data pekerja dihapus, akun ini akan hilang otomatis."
                });
            }

            var result = await _userManager.DeleteAsync(user);
            if (!result.Succeeded)
            {
                var msg = string.Join("; ", result.Errors.Select(e => e.Description));
                return Json(new { success = false, message = msg });
            }

            return Json(new { success = true, message = "Akun berhasil dihapus." });
        }

        // =========================================================
        // DOWNLOAD TEMPLATE IMPORT USER (KOSONG)
        // =========================================================
        [HttpGet]
        public IActionResult DownloadTemplate()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("TemplateUsers");

            ws.Cells[1, 1].Value = "Nopeg";  // Kolom A
            ws.Cells[1, 2].Value = "Nama";   // Kolom B
            ws.Cells[1, 3].Value = "Email";  // Kolom C

            using (var range = ws.Cells[1, 1, 1, 3])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                     .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            const string fileName = "Template_Import_Akun.xlsx";
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(bytes, contentType, fileName);
        }

        // =========================================================
        // IMPORT AKUN DARI EXCEL (Nopeg, Nama, Email)
        // =========================================================
        [HttpGet]
        public IActionResult Import() => PartialView("Import");

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return Json(new { success = false, message = "File tidak ditemukan." });

            var employees = await _db.Employees
                .Select(e => new
                {
                    e.Id,
                    e.NamaLengkap,
                    e.NopegPersero,
                    e.EmailPertamina
                })
                .ToListAsync();

            var created = 0;
            var skipped = new List<string>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using var package = new ExcelPackage(stream);
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null)
                    return Json(new { success = false, message = "Worksheet tidak ditemukan." });

                var row = 2;
                while (true)
                {
                    var nopeg = ws.Cells[row, 1].Text?.Trim();
                    var nama = ws.Cells[row, 2].Text?.Trim();
                    var email = ws.Cells[row, 3].Text?.Trim();

                    if (string.IsNullOrEmpty(nopeg) &&
                        string.IsNullOrEmpty(nama) &&
                        string.IsNullOrEmpty(email))
                        break;

                    var emp = employees.FirstOrDefault(e =>
                        (!string.IsNullOrEmpty(nopeg) &&
                         string.Equals(e.NopegPersero ?? "", nopeg, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(nama) &&
                         string.Equals(e.NamaLengkap ?? "", nama, StringComparison.OrdinalIgnoreCase)));

                    if (emp == null)
                    {
                        skipped.Add($"Baris {row}: Pegawai tidak ditemukan (Nopeg '{nopeg}', Nama '{nama}').");
                        row++;
                        continue;
                    }

                    var finalEmail = !string.IsNullOrWhiteSpace(email)
                        ? email
                        : emp.EmailPertamina;

                    if (string.IsNullOrWhiteSpace(finalEmail))
                    {
                        skipped.Add($"Baris {row}: Email kosong (tidak di file dan tidak di data pegawai).");
                        row++;
                        continue;
                    }

                    var existingUser = await _userManager.FindByEmailAsync(finalEmail);
                    if (existingUser != null)
                    {
                        skipped.Add($"Baris {row}: Email {finalEmail} sudah punya akun.");
                        row++;
                        continue;
                    }

                    var user = new ApplicationUser
                    {
                        Email = finalEmail,
                        UserName = finalEmail,
                        FullName = emp.NamaLengkap,
                        EmployeeId = emp.Id,
                        EmailConfirmed = true,
                        MustChangePassword = true,   // << WAJIB GANTI PASSWORD
                        LockoutEnabled = true        // << aktifkan lockout
                    };

                    var result = await _userManager.CreateAsync(user, DefaultPassword);
                    if (result.Succeeded)
                    {
                        created++;
                        if (!await _roleManager.RoleExistsAsync("User"))
                            await _roleManager.CreateAsync(new IdentityRole("User"));

                        await _userManager.AddToRoleAsync(user, "User");

                        // sync ke employee juga (nama & email)
                        await SyncEmployeeFromUserAsync(user);
                    }
                    else
                    {
                        skipped.Add($"Baris {row}: {string.Join(", ", result.Errors.Select(e => e.Description))}");
                    }

                    row++;
                }
            }

            var msg = $"Berhasil membuat {created} akun baru.";
            if (skipped.Count > 0)
                msg += $" {skipped.Count} baris dilewati.";

            return Json(new { success = true, message = msg, errors = skipped });
        }

        // =========================================================
        // EXPORT USERS (EXCEL & CSV)
        // =========================================================
        [HttpGet]
        public async Task<IActionResult> ExportExcel(string? q, int? employeeId)
        {
            var usersQuery = _userManager.Users.Include(u => u.Employee).AsQueryable();

            if (employeeId is not null)
                usersQuery = usersQuery.Where(u => u.EmployeeId == employeeId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                usersQuery = usersQuery.Where(u =>
                    (u.Email ?? "").Contains(q) ||
                    (u.FullName ?? "").Contains(q) ||
                    (u.Employee!.NopegPersero ?? "").Contains(q));
            }

            var users = await usersQuery.OrderBy(u => u.FullName ?? u.Email).ToListAsync();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Users");

            ws.Cells[1, 1].Value = "FullName";
            ws.Cells[1, 2].Value = "Email";
            ws.Cells[1, 3].Value = "Nopeg";
            ws.Cells[1, 4].Value = "EmployeeId";

            using (var range = ws.Cells[1, 1, 1, 4])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor
                    .SetColor(System.Drawing.Color.FromArgb(234, 243, 255));
            }

            var row = 2;
            foreach (var u in users)
            {
                ws.Cells[row, 1].Value = u.FullName;
                ws.Cells[row, 2].Value = u.Email;
                ws.Cells[row, 3].Value = u.Employee?.NopegPersero;
                ws.Cells[row, 4].Value = u.EmployeeId;
                row++;
            }

            ws.Cells.AutoFitColumns();

            var bytes = package.GetAsByteArray();
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Users.xlsx");
        }

        [HttpGet]
        public async Task<IActionResult> ExportCsv(string? q, int? employeeId)
        {
            var usersQuery = _userManager.Users.Include(u => u.Employee).AsQueryable();

            if (employeeId is not null)
                usersQuery = usersQuery.Where(u => u.EmployeeId == employeeId);

            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                usersQuery = usersQuery.Where(u =>
                    (u.Email ?? "").Contains(q) ||
                    (u.FullName ?? "").Contains(q) ||
                    (u.Employee!.NopegPersero ?? "").Contains(q));
            }

            var users = await usersQuery.OrderBy(u => u.FullName ?? u.Email).ToListAsync();

            var sb = new StringBuilder();
            sb.AppendLine("FullName,Email,Nopeg,EmployeeId");

            foreach (var u in users)
            {
                string line = string.Join(",",
                    Quote(u.FullName),
                    Quote(u.Email),
                    Quote(u.Employee?.NopegPersero),
                    u.EmployeeId?.ToString() ?? "");
                sb.AppendLine(line);
            }

            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            return File(bytes, "text/csv", "Users.csv");

            static string Quote(string? value)
            {
                value ??= "";
                if (value.Contains('"') || value.Contains(','))
                    value = "\"" + value.Replace("\"", "\"\"") + "\"";
                return value;
            }
        }
    }
}
