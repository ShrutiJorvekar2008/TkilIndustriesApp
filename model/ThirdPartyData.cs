using System.ComponentModel.DataAnnotations;

namespace TkilIndustriesApp.Models
{
    public class ThirdPartyData
    {

        [MaxLength(8)]
        public string EmployeeCode { get; set; } = string.Empty;

        [MaxLength(8)]
        public string Employee8ID { get; set; } = string.Empty;

        public string ROLL { get; set; } = string.Empty;
        public string LOCATION_Sit_HO_Branch { get; set; } = string.Empty;
        public string DEPT { get; set; } = string.Empty;
        public string DEPT_NAME { get; set; } = string.Empty;
        public string Cost_Center { get; set; } = string.Empty;
        public string Salutation { get; set; } = string.Empty;
        public string Employee_Name { get; set; } = string.Empty;
        public string NEW_DESIGNATION { get; set; } = string.Empty;
        public string GRADE { get; set; } = string.Empty;
        public DateTime? Date_of_Joining { get; set; }
        public string Reporting_Manager { get; set; } = string.Empty;
        public string Functional_Manager { get; set; } = string.Empty;
        public string CAT { get; set; } = string.Empty;
        public string Category_For_Leave_Valuation { get; set; } = string.Empty;
        public string Gender { get; set; } = string.Empty;
        public string Business_Unit { get; set; } = string.Empty;

        [EmailAddress]
        public string Email { get; set; } = string.Empty;

        public string Active { get; set; } = "Active";
    }
}
