using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;

namespace UtilitiesHR.Data
{
    // Dipakai EF Tools saat Add-Migration / Update-Database
    public class DesignTimeDbContextFactory : IDesignTimeDbContextFactory<ApplicationDbContext>
    {
        public ApplicationDbContext CreateDbContext(string[] args)
        {
            // Base path = folder proyek web UtilitiesHR (tempat file ini berada)
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false)
                .AddJsonFile("appsettings.Development.json", optional: true)
                .AddEnvironmentVariables()
                .Build();

            var cs = config.GetConnectionString("DefaultConnection")
                ?? "Server=DESKTOP-V4I4VVA\\MSSQLSERVER2022;Database=UtilitiesHR;Trusted_Connection=True;MultipleActiveResultSets=true;TrustServerCertificate=True";

            var options = new DbContextOptionsBuilder<ApplicationDbContext>()
                .UseSqlServer(cs)
                .Options;

            return new ApplicationDbContext(options);
        }
    }
}
