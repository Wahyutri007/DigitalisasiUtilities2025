using System;
using System.ComponentModel.DataAnnotations;

namespace UtilitiesHR.Models
{
    public enum WorkTaskStatus
    {
        [Display(Name = "Pending")]
        Pending = 0,

        [Display(Name = "In Progress")]
        InProgress = 1,

        [Display(Name = "Done")]
        Done = 2
    }

    public enum WorkTaskPriority
    {
        [Display(Name = "Low")]
        Low = 0,

        [Display(Name = "Medium")]
        Medium = 1,

        [Display(Name = "High")]
        High = 2,

        [Display(Name = "Emergency")]
        Emergency = 3
    }

    public class WorkTask
    {
        public int Id { get; set; }

        [Required]
        [Display(Name = "Nama Pekerjaan")]
        [StringLength(200)]
        public string NamaPekerjaan { get; set; } = string.Empty;

        [Display(Name = "Start Date")]
        [DataType(DataType.Date)]
        public DateTime? StartDate { get; set; }

        [Display(Name = "Due Date")]
        [DataType(DataType.Date)]
        public DateTime? DueDate { get; set; }

        [Display(Name = "Tanggal Penyelesaian")]
        [DataType(DataType.Date)]
        public DateTime? CompletedDate { get; set; }

        [Display(Name = "Nama Request")]
        [StringLength(200)]
        public string? NamaRequest { get; set; }

        [Display(Name = "Progres (%)")]
        [Range(0, 100)]
        public int ProgressPercent { get; set; }

        [Display(Name = "Status")]
        public WorkTaskStatus Status { get; set; } = WorkTaskStatus.Pending;

        [Display(Name = "Keterangan Tindak Lanjut")]
        [DataType(DataType.MultilineText)]
        [StringLength(1000)]
        public string? KeteranganTindakLanjut { get; set; }

        [Display(Name = "Prioritas")]
        public WorkTaskPriority Prioritas { get; set; } = WorkTaskPriority.Medium;

        [Display(Name = "Nama PIC")]
        [StringLength(150)]
        public string? NamaPIC { get; set; }
    }
}
