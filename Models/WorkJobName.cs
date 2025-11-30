using System;
using System.ComponentModel.DataAnnotations;

namespace UtilitiesHR.Models
{
    public class WorkJobName
    {
        public int Id { get; set; }

        [Required]
        [StringLength(200)]
        [Display(Name = "Nama Pekerjaan")]
        public string NamaPekerjaan { get; set; } = string.Empty;

        [Display(Name = "Tanggal Dibuat")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
