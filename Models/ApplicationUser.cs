using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations.Schema;

namespace UtilitiesHR.Models
{
    public class ApplicationUser : IdentityUser
    {
        public string? FullName { get; set; }

        public int? EmployeeId { get; set; }

        [ForeignKey(nameof(EmployeeId))]
        public Employee? Employee { get; set; }
        // Menyimpan kapan password terakhir berhasil diubah.
        /// </summary>
        public DateTimeOffset? LastPasswordChangedAt { get; set; }

        /// <summary>
        /// Flag untuk paksa user ganti password setelah login.
        /// </summary>
        public bool MustChangePassword { get; set; } = false;
    }
}