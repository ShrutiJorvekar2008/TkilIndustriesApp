using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace TkilIndustriesApp.Models
{
    [Table("IT_ITRecords")]
    public class IT
    {
        [Key]
        [EmailAddress]
        public string EmailAddress { get; set; } = string.Empty;

        public string EmpID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Dept { get; set; } = string.Empty;
        public string BU { get; set; } = string.Empty;
        public string? CostCenter { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? O365 { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Internet { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Servers { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? VisualStudio { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? TeklaStructure { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Tally { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? StadProConnect { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SQLServerStdCore { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SimuTherm { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SAPCrystalReport { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? ProE { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PrimaVera { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? VaultProfessional { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? OracleStandard { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? NeiNastran { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? NavisworkManage { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? NavisworkFreedom { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? MSVisioProf { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? MSProjectProf { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Femap_NXNastran { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Ceaser { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? ChemCAD { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? DSignPdfSigner { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PDMC { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? AEC { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? MasterCAM219 { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Mill3D { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? STARCCM { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? EEC { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? ThinkCell { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? EfficientElement { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PDFExchangePro { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Dameware { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PowerBIPremium { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PowerBIPro { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? BeltAnalyst { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? PDMC_Inventor_AutoCAD { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Damware { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SDP { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SentinelOne { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? SAP_7_7 { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? GoTo { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? AutodeskVault { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Pipenet { get; set; }
        [Column(TypeName = "decimal(18,2)")]
        public decimal? Mecastack { get; set; }
    }
}
