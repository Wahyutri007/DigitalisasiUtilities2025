using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    public class AccountController : Controller
    {
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly UserManager<ApplicationUser> _userManager;

        // password default global
        private const string DefaultPassword = "Pertamina@2025";

        public AccountController(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager)
        {
            _userManager = userManager;
            _signInManager = signInManager;
        }

        // ==================== LOGIN ====================

        [HttpGet]
        [AllowAnonymous]
        public IActionResult Login(string? returnUrl = null)
        {
            if (User.Identity?.IsAuthenticated == true)
                return RedirectToAction("Index", "Home");

            ViewData["ReturnUrl"] = returnUrl;
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(
            string email,
            string password,
            bool rememberMe = false,
            string? returnUrl = null)
        {
            ViewData["ReturnUrl"] = returnUrl;

            // validasi basic
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(password))
            {
                ModelState.AddModelError(string.Empty, "Email dan password wajib diisi.");
                return View();
            }

            var user = await _userManager.FindByEmailAsync(email);
            if (user == null)
            {
                // Jangan bocorkan info: tetap bilang email/password salah
                ModelState.AddModelError(string.Empty, "Email atau password salah.");
                return View();
            }

            // 🌟 CASE KHUSUS default password:
            // blok hanya kalau:
            // - MustChangePassword == false (sudah pernah ganti password),
            // - password input = default,
            // - LastPasswordChangedAt sudah ada
            if (!user.MustChangePassword &&
                password == DefaultPassword &&
                user.LastPasswordChangedAt.HasValue)
            {
                var info = BuildPasswordChangedMessage(user.LastPasswordChangedAt.Value);
                var msg = $"Password default tidak lagi berlaku. {info} Gunakan password terbaru Anda. Jika lupa, hubungi admin untuk reset.";

                ModelState.AddModelError(string.Empty, msg);
                return View();
            }

            // proses login Identity + AKTIFKAN LOCKOUT
            var result = await _signInManager.PasswordSignInAsync(
                user.UserName!,
                password,
                rememberMe,
                lockoutOnFailure: true); // penting untuk limit percobaan

            // kalau akun terkunci karena gagal berulang
            if (result.IsLockedOut)
            {
                var lockoutEnd = await _userManager.GetLockoutEndDateAsync(user);

                string msg;
                if (lockoutEnd.HasValue)
                {
                    msg = BuildLockoutMessage(lockoutEnd.Value);
                }
                else
                {
                    msg = "Akun Anda terkunci karena terlalu banyak percobaan login gagal. Silakan coba lagi beberapa menit lagi.";
                }

                ModelState.AddModelError(string.Empty, msg);
                return View();
            }

            if (!result.Succeeded)
            {
                ModelState.AddModelError(string.Empty, "Email atau password salah.");
                return View();
            }

            // ====== LOGIC WAJIB GANTI PASSWORD PERTAMA KALI ======
            if (user.MustChangePassword || password == DefaultPassword)
            {
                if (!user.MustChangePassword)
                {
                    user.MustChangePassword = true;
                    await _userManager.UpdateAsync(user);
                }

                return RedirectToAction(nameof(ChangePasswordFirstTime));
            }

            // (opsional) Info kapan password terakhir diubah
            if (user.LastPasswordChangedAt.HasValue)
            {
                TempData["PasswordInfo"] = BuildPasswordChangedMessage(user.LastPasswordChangedAt.Value);
            }

            // kalau ada returnUrl valid, pakai
            if (!string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl))
                return Redirect(returnUrl);

            // default setelah login
            return RedirectToAction("Index", "Home");
        }

        // ==================== LOGOUT ====================

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await _signInManager.SignOutAsync();
            return RedirectToAction("Login", "Account");
        }

        // ==================== ACCESS DENIED ====================

        [HttpGet]
        public IActionResult AccessDenied()
        {
            return View();
        }

        // ==================== PROFILE (DATA DIRI) ====================

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> Profile()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return RedirectToAction("Login");

            var vm = new ProfileVM
            {
                Email = user.Email ?? "",
                FullName = user.FullName ?? "",
                PhoneNumber = user.PhoneNumber ?? ""
            };

            ViewBag.PasswordInfo = user.LastPasswordChangedAt.HasValue
                ? BuildPasswordChangedMessage(user.LastPasswordChangedAt.Value)
                : "Password belum pernah tercatat diubah.";

            ViewBag.ChangePasswordError = false;

            return View(vm);
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Profile(ProfileVM model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var user = await _userManager.GetUserAsync(User);
            if (user == null) return RedirectToAction("Login");

            user.FullName = model.FullName;
            user.PhoneNumber = model.PhoneNumber;

            // boleh ubah email → sekalian jadikan username
            if (!string.IsNullOrWhiteSpace(model.Email) &&
                !string.Equals(user.Email, model.Email, StringComparison.OrdinalIgnoreCase))
            {
                var existing = await _userManager.FindByEmailAsync(model.Email);
                if (existing != null && existing.Id != user.Id)
                {
                    ModelState.AddModelError("Email", "Email sudah digunakan akun lain.");
                    return View(model);
                }

                user.Email = model.Email;
                user.UserName = model.Email;
            }

            var result = await _userManager.UpdateAsync(user);
            if (result.Succeeded)
            {
                TempData["ProfileMessage"] = "Profil berhasil diperbarui.";
                return RedirectToAction("Profile");
            }

            foreach (var e in result.Errors)
                ModelState.AddModelError(string.Empty, e.Description);

            return View(model);
        }

        // ==================== GANTI PASSWORD BIASA (DARI PROFILE) ====================

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ChangePassword(ChangePasswordVM model)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null)
            {
                return RedirectToAction("Login");
            }

            async Task<IActionResult> ReturnProfileWithError()
            {
                var profileVm = new ProfileVM
                {
                    Email = user.Email ?? "",
                    FullName = user.FullName ?? "",
                    PhoneNumber = user.PhoneNumber ?? ""
                };

                ViewBag.PasswordInfo = user.LastPasswordChangedAt.HasValue
                    ? BuildPasswordChangedMessage(user.LastPasswordChangedAt.Value)
                    : "Password belum pernah tercatat diubah.";

                ViewBag.ChangePasswordError = true;
                return View("Profile", profileVm);
            }

            if (!ModelState.IsValid)
            {
                return await ReturnProfileWithError();
            }

            if (model.NewPassword != model.ConfirmPassword)
            {
                ModelState.AddModelError(string.Empty, "Password baru dan konfirmasi tidak sama.");
                return await ReturnProfileWithError();
            }

            if (model.NewPassword == DefaultPassword)
            {
                ModelState.AddModelError(string.Empty,
                    "Password baru tidak boleh sama dengan password default sistem.");
                return await ReturnProfileWithError();
            }

            if (model.NewPassword == model.OldPassword)
            {
                ModelState.AddModelError(string.Empty,
                    "Password baru tidak boleh sama dengan password lama Anda.");
                return await ReturnProfileWithError();
            }

            var result = await _userManager.ChangePasswordAsync(user, model.OldPassword, model.NewPassword);
            if (!result.Succeeded)
            {
                foreach (var e in result.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);

                return await ReturnProfileWithError();
            }

            user.MustChangePassword = false;
            user.LastPasswordChangedAt = DateTimeOffset.UtcNow;
            await _userManager.UpdateAsync(user);
            await _signInManager.RefreshSignInAsync(user);

            TempData["SuccessMessage"] = "Password berhasil diubah.";
            return RedirectToAction("Profile");
        }

        // ==================== GANTI PASSWORD PERTAMA KALI (WAJIB) ====================

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> ChangePasswordFirstTime()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null)
                return RedirectToAction("Login");

            // Kalau user TIDAK lagi wajib ganti password,
            // jangan boleh buka halaman ini
            if (!user.MustChangePassword)
            {
                return RedirectToAction("Index", "Home");
            }

            ModelState.Clear();
            return View(new ChangePasswordFirstTimeVM());
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ChangePasswordFirstTime(ChangePasswordFirstTimeVM model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var user = await _userManager.GetUserAsync(User);
            if (user == null)
            {
                ModelState.AddModelError(string.Empty, "User tidak ditemukan.");
                return View(model);
            }

            if (model.NewPassword != model.ConfirmPassword)
            {
                ModelState.AddModelError(string.Empty, "Password baru dan konfirmasi tidak sama.");
                return View(model);
            }

            // Tidak boleh pakai password default lagi
            if (model.NewPassword == DefaultPassword)
            {
                ModelState.AddModelError(string.Empty,
                    "Password baru tidak boleh sama dengan password default sistem.");
                return View(model);
            }

            // Tidak boleh sama dengan password lama (default)
            if (model.NewPassword == model.OldPassword)
            {
                ModelState.AddModelError(string.Empty,
                    "Password baru tidak boleh sama dengan password lama Anda.");
                return View(model);
            }

            var result = await _userManager.ChangePasswordAsync(user, model.OldPassword, model.NewPassword);
            if (!result.Succeeded)
            {
                foreach (var e in result.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);
                return View(model);
            }

            // Setelah password diganti, clear flag wajib ganti password
            user.MustChangePassword = false;
            user.LastPasswordChangedAt = DateTimeOffset.UtcNow;
            await _userManager.UpdateAsync(user);
            await _signInManager.RefreshSignInAsync(user);

            TempData["SuccessMessage"] = "Password berhasil diubah.";

            // 🔁 SETELAH GANTI PASSWORD PERTAMA KALI:
            // arahkan ke HOME + kirim id pekerja untuk popup edit
            if (user.EmployeeId is int empId && empId > 0)
            {
                // nanti di layout akan baca query firstEditId dan buka popup di Home
                return RedirectToAction("Index", "Home", new { firstEditId = empId });
            }

            // fallback kalau tidak punya EmployeeId
            return RedirectToAction("Index", "Home");
        }

        // ==================== HELPER FORMAT PASSWORD AGE ====================

        private static string BuildPasswordChangedMessage(DateTimeOffset lastChanged)
        {
            var span = DateTimeOffset.UtcNow - lastChanged;

            if (span.TotalDays < 1)
            {
                var hours = (int)Math.Floor(span.TotalHours);

                if (hours <= 0)
                    return "Password Anda terakhir diubah kurang dari 1 jam yang lalu.";

                if (hours == 1)
                    return "Password Anda terakhir diubah 1 jam yang lalu.";

                return $"Password Anda terakhir diubah {hours} jam yang lalu.";
            }

            var days = (int)Math.Floor(span.TotalDays);

            if (days == 1)
                return "Password Anda terakhir diubah 1 hari yang lalu.";

            return $"Password Anda terakhir diubah {days} hari yang lalu.";
        }

        // ==================== HELPER FORMAT LOCKOUT ====================

        private static string BuildLockoutMessage(DateTimeOffset lockoutEnd)
        {
            var remaining = lockoutEnd - DateTimeOffset.UtcNow;

            if (remaining.TotalSeconds <= 0)
            {
                return "Akun Anda sementara terkunci. Silakan coba login kembali.";
            }

            var totalMinutes = (int)Math.Ceiling(remaining.TotalMinutes);

            if (totalMinutes <= 1)
            {
                return "Akun Anda terkunci karena terlalu banyak percobaan login gagal. Silakan coba lagi dalam sekitar 1 menit.";
            }

            return $"Akun Anda terkunci karena terlalu banyak percobaan login gagal. Silakan coba lagi dalam sekitar {totalMinutes} menit.";
        }

        // ==================== VIEWMODELS ====================

        public class ProfileVM
        {
            public string Email { get; set; } = "";
            public string FullName { get; set; } = "";
            public string PhoneNumber { get; set; } = "";
        }

        public class ChangePasswordVM
        {
            public string OldPassword { get; set; } = "";
            public string NewPassword { get; set; } = "";
            public string ConfirmPassword { get; set; } = "";
        }

        public class ChangePasswordFirstTimeVM
        {
            public string OldPassword { get; set; } = "";
            public string NewPassword { get; set; } = "";
            public string ConfirmPassword { get; set; } = "";
        }
    }
}
