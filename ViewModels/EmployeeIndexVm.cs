using System.Collections.Generic;
using UtilitiesHR.Models;

namespace UtilitiesHR.ViewModels
{
    public class EmployeeIndexVm
    {
        public int TotalEmployee { get; set; }
        public int TotalEmail { get; set; }
        public int TotalPhone { get; set; }
        public string? Search { get; set; }
        public List<Employee> Items { get; set; } = new();
    }
}
