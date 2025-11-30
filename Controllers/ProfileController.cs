using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;
using UtilitiesHR.Models;

namespace UtilitiesHR.Controllers
{
    [Authorize]
    public class ProfileController : Controller
    {
        private readonly UserManager<ApplicationUser> _userManager;

        public ProfileController(UserManager<ApplicationUser> userManager)
        {
            _userManager = userManager;
        }

        public class ProfileVM
        {
            public string Email { get; set; } = "";
            public string FullName { get; set; } = "";
            public string PhoneNumber { get; set; } = "";

            // info employee (readonly di view)
            public string? EmployeeName { get; set; }
            public string? Nopeg { get; set; }

            // untuk ubah password di halaman profil
            public string OldPassword { get; set; } = "";
            public string NewPassword { get; set; } = "";
            public string ConfirmPassword { get; set; } = "";
        }

        // ================== GET PROFIL ==================
        [HttpGet]
        public async Task<IActionResult> Index()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null)
                return RedirectToAction("Login", "Account");

            var vm = new ProfileVM
            {
                Email = user.Email ?? "",
                FullName = user.FullName ?? "",
                PhoneNumber = user.PhoneNumber ?? "",
                EmployeeName = user.Employee?.NamaLengkap,
                Nopeg = user.Employee?.NopegPersero
            };

            return View(vm);
        }

        // ================== POST PROFIL ==================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Index(ProfileVM model)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null)
                return RedirectToAction("Login", "Account");

            if (!ModelState.IsValid)
                return View(model);

            // --- Update data profil dasar ---
            user.FullName = model.FullName;
            user.PhoneNumber = model.PhoneNumber;

            var updateResult = await _userManager.UpdateAsync(user);
            if (!updateResult.Succeeded)
            {
                foreach (var e in updateResult.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);

                // pastikan field readonly tetap terisi kalau balik ke view
                model.EmployeeName = user.Employee?.NamaLengkap;
                model.Nopeg = user.Employee?.NopegPersero;
                return View(model);
            }

            // --- Cek apakah user ingin ganti password ---
            bool wantsPasswordChange =
                !string.IsNullOrWhiteSpace(model.OldPassword) ||
                !string.IsNullOrWhiteSpace(model.NewPassword) ||
                !string.IsNullOrWhiteSpace(model.ConfirmPassword);

            if (wantsPasswordChange)
            {
                // semua field password wajib diisi
                if (string.IsNullOrWhiteSpace(model.OldPassword) ||
                    string.IsNullOrWhiteSpace(model.NewPassword) ||
                    string.IsNullOrWhiteSpace(model.ConfirmPassword))
                {
                    ModelState.AddModelError(string.Empty,
                        "Untuk mengubah password, isi password lama, password baru, dan konfirmasi password.");
                    model.EmployeeName = user.Employee?.NamaLengkap;
                    model.Nopeg = user.Employee?.NopegPersero;
                    return View(model);
                }

                if (model.NewPassword != model.ConfirmPassword)
                {
                    ModelState.AddModelError("ConfirmPassword", "Konfirmasi password baru tidak sama.");
                    model.EmployeeName = user.Employee?.NamaLengkap;
                    model.Nopeg = user.Employee?.NopegPersero;
                    return View(model);
                }

                if (model.NewPassword == model.OldPassword)
                {
                    ModelState.AddModelError("NewPassword", "Password baru tidak boleh sama dengan password lama.");
                    model.EmployeeName = user.Employee?.NamaLengkap;
                    model.Nopeg = user.Employee?.NopegPersero;
                    return View(model);
                }

                // pakai mekanisme ChangePassword (wajib old password benar)
                var pwdResult = await _userManager.ChangePasswordAsync(
                    user,
                    model.OldPassword,
                    model.NewPassword
                );

                if (!pwdResult.Succeeded)
                {
                    foreach (var e in pwdResult.Errors)
                        ModelState.AddModelError(string.Empty, e.Description);

                    model.EmployeeName = user.Employee?.NamaLengkap;
                    model.Nopeg = user.Employee?.NopegPersero;
                    return View(model);
                }

                // catat kapan password diubah (kalau property ini ada di ApplicationUser)
                user.LastPasswordChangedAt = DateTimeOffset.UtcNow;
                await _userManager.UpdateAsync(user);

                TempData["ProfileMessage"] = "Profil dan password berhasil diperbarui.";
            }
            else
            {
                TempData["ProfileMessage"] = "Profil berhasil diperbarui.";
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
