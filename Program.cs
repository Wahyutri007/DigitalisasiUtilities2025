using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using UtilitiesHR.Data;
using UtilitiesHR.Models;

var builder = WebApplication.CreateBuilder(args);

// =====================================================
// EPPlus License (EPPlus 8 -> WAJIB PAKAI INI)
// =====================================================
ExcelPackage.License.SetNonCommercialOrganization("UtilitiesHR");

// =====================================================
// Services
// =====================================================

// DbContext
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(
        builder.Configuration.GetConnectionString("DefaultConnection")
    )
);

// Identity + ApplicationUser
builder.Services.AddIdentity<ApplicationUser, IdentityRole>(options =>
{
    options.Password.RequiredLength = 8;
    options.Password.RequireDigit = true;
    options.Password.RequireUppercase = true;
    options.Password.RequireLowercase = true;
    options.Password.RequireNonAlphanumeric = true;

    options.SignIn.RequireConfirmedAccount = false;

    options.Lockout.MaxFailedAccessAttempts = 5;
    options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(10);
    options.Lockout.AllowedForNewUsers = true;
})
.AddEntityFrameworkStores<ApplicationDbContext>()
.AddDefaultTokenProviders();

// Cookie auth
builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Account/Login";          // default halaman login
    options.LogoutPath = "/Account/Logout";
    options.AccessDeniedPath = "/Account/AccessDenied";
    options.SlidingExpiration = true;
});

builder.Services.AddControllersWithViews();

var app = builder.Build();

// =====================================================
// SEED ROLE + SUPERADMIN USER
// =====================================================
using (var scope = app.Services.CreateScope())
{
    var roleManager = scope.ServiceProvider.GetRequiredService<RoleManager<IdentityRole>>();
    var userManager = scope.ServiceProvider.GetRequiredService<UserManager<ApplicationUser>>();

    // ---- Seed Roles ----
    string[] roles = new[] { "User", "Admin", "SuperAdmin", "Supervisor" };

    foreach (var roleName in roles)
    {
        if (!await roleManager.RoleExistsAsync(roleName))
        {
            await roleManager.CreateAsync(new IdentityRole(roleName));
        }
    }

    // ---- Seed SuperAdmin User ----
    // NOTE: kalau mau ganti email/password default, ubah di sini
    const string superAdminEmail = "superadmin@utilitieshr.local";
    const string superAdminPassword = "Pertamina@2025";

    var existingSuperAdmin = await userManager.FindByEmailAsync(superAdminEmail);
    if (existingSuperAdmin == null)
    {
        var superAdmin = new ApplicationUser
        {
            UserName = superAdminEmail,
            Email = superAdminEmail,
            EmailConfirmed = true,
            // Kalau kamu punya properti lain di ApplicationUser, boleh diisi di sini.
            MustChangePassword = true  // supaya dipaksa ganti password saat pertama login
        };

        var createResult = await userManager.CreateAsync(superAdmin, superAdminPassword);

        if (createResult.Succeeded)
        {
            // pastikan dia ada di role SuperAdmin
            if (!await userManager.IsInRoleAsync(superAdmin, "SuperAdmin"))
            {
                await userManager.AddToRoleAsync(superAdmin, "SuperAdmin");
            }
        }
        else
        {
            // Optional: log error kalau mau
            // var logger = scope.ServiceProvider.GetRequiredService<ILogger<Program>>();
            // logger.LogError("Gagal membuat SuperAdmin: {Errors}", string.Join("; ", createResult.Errors.Select(e => e.Description)));
        }
    }
}

// =====================================================
// Middleware Pipeline
// =====================================================

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Error/ServerError");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();

// =====================================================
// Middleware: Force user change password on login
// =====================================================
app.Use(async (context, next) =>
{
    var path = context.Request.Path.Value?.ToLowerInvariant() ?? "";

    if (path.Contains("/account/") || path.StartsWith("/error"))
    {
        await next();
        return;
    }

    if (path.StartsWith("/css") ||
        path.StartsWith("/js") ||
        path.StartsWith("/lib") ||
        path.StartsWith("/images") ||
        path.StartsWith("/img") ||
        path.StartsWith("/favicon"))
    {
        await next();
        return;
    }

    if (context.User?.Identity?.IsAuthenticated == true)
    {
        var userManager = context.RequestServices.GetRequiredService<UserManager<ApplicationUser>>();
        var user = await userManager.GetUserAsync(context.User);

        if (user != null && user.MustChangePassword)
        {
            context.Response.Redirect("/Account/ChangePasswordFirstTime");
            return;
        }
    }

    await next();
});

app.UseAuthorization();

// =====================================================
// Status codes redirect
// =====================================================
app.UseStatusCodePagesWithReExecute("/Error/StatusCode", "?code={0}");

// =====================================================
// Custom Root Redirect
// =====================================================
// Jika user membuka "/", arahkan ke Dashboard (Employees/Index)
app.MapGet("/", context =>
{
    if (!context.User.Identity.IsAuthenticated)
    {
        context.Response.Redirect("/Account/Login");
    }
    else
    {
        context.Response.Redirect("/Employees/Index");
    }

    return Task.CompletedTask;
});

// =====================================================
// Controller Routing
// =====================================================
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Employees}/{action=Index}/{id?}"
);

app.Run();
