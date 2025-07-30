using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TkilIndustriesApp.Models
{
    [Table("IT_LicenseMaster")] 
    public class LicenseMaster
    {
        [Key]
        [Required]
        [StringLength(100)]
        public string LicenseName { get; set; } = string.Empty;

        [Required]
        [Column(TypeName = "decimal(18,2)")]
        public decimal CostPerUser { get; set; }

    }
}
