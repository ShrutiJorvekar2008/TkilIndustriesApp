using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.IO;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Net;
using TkilIndustriesApp.Data;
using TkilIndustriesApp.Models;


public class UploadController : Controller
{
    private readonly TkilContext _context;


    public UploadController(TkilContext context)
    {
        _context = context;
    }

    public IActionResult Index(string message = "")
    {
        ViewBag.Message = message;
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> UploadITFile(IFormFile file)
    {
        if (file != null && file.Length > 0)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(file.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;
                    var records = new List<IT>();

                    for (int row = 2; row <= rowCount; row++) // Assuming first row is header
                    {
                        var IsEmailCheck = worksheet.Cells[row, 1].Text?.Trim();
                        if (IsEmailCheck != "" && IsEmailCheck != null && IsEmailCheck != string.Empty)
                        {
                            var record = new IT
                            {
                                EmailAddress = worksheet.Cells[row, 1].Text?.Trim() ?? string.Empty,
                                EmpID = worksheet.Cells[row, 2].Text?.Trim() ?? string.Empty,
                                Name = worksheet.Cells[row, 3].Text?.Trim() ?? string.Empty,
                                Dept = worksheet.Cells[row, 4].Text?.Trim() ?? string.Empty,
                                BU = worksheet.Cells[row, 5].Text?.Trim() ?? string.Empty,
                                CostCenter = worksheet.Cells[row, 6].Text?.Trim(),
                                O365 = ParseDecimal(worksheet.Cells[row, 7].Value),
                                Internet = ParseDecimal(worksheet.Cells[row, 8].Value),
                                Servers = ParseDecimal(worksheet.Cells[row, 9].Value),
                                VisualStudio = ParseDecimal(worksheet.Cells[row, 10].Value),
                                TeklaStructure = ParseDecimal(worksheet.Cells[row, 11].Value),
                                Tally = ParseDecimal(worksheet.Cells[row, 12].Value),
                                StadProConnect = ParseDecimal(worksheet.Cells[row, 13].Value),
                                SQLServerStdCore = ParseDecimal(worksheet.Cells[row, 14].Value),
                                SimuTherm = ParseDecimal(worksheet.Cells[row, 15].Value),
                                SAPCrystalReport = ParseDecimal(worksheet.Cells[row, 16].Value),
                                ProE = ParseDecimal(worksheet.Cells[row, 17].Value),
                                PrimaVera = ParseDecimal(worksheet.Cells[row, 18].Value),
                                VaultProfessional = ParseDecimal(worksheet.Cells[row, 19].Text),
                                OracleStandard = ParseDecimal(worksheet.Cells[row, 20].Text),
                                NeiNastran = ParseDecimal(worksheet.Cells[row, 21].Text),
                                NavisworkManage = ParseDecimal(worksheet.Cells[row, 22].Text),
                                NavisworkFreedom = ParseDecimal(worksheet.Cells[row, 23].Text),
                                MSVisioProf = ParseDecimal(worksheet.Cells[row, 24].Text),
                                MSProjectProf = ParseDecimal(worksheet.Cells[row, 25].Text),
                                Femap_NXNastran = ParseDecimal(worksheet.Cells[row, 26].Text),
                                Ceaser = ParseDecimal(worksheet.Cells[row, 27].Text),
                                ChemCAD = ParseDecimal(worksheet.Cells[row, 28].Text),
                                DSignPdfSigner = ParseDecimal(worksheet.Cells[row, 29].Text),
                                PDMC = ParseDecimal(worksheet.Cells[row, 30].Text),
                                AEC = ParseDecimal(worksheet.Cells[row, 31].Text),
                                MasterCAM219 = ParseDecimal(worksheet.Cells[row, 32].Text),
                                Mill3D = ParseDecimal(worksheet.Cells[row, 33].Text),
                                STARCCM = ParseDecimal(worksheet.Cells[row, 34].Text),
                                EEC = ParseDecimal(worksheet.Cells[row, 35].Text),
                                ThinkCell = ParseDecimal(worksheet.Cells[row, 36].Text),
                                EfficientElement = ParseDecimal(worksheet.Cells[row, 37].Text),
                                PDFExchangePro = ParseDecimal(worksheet.Cells[row, 38].Text),
                                Dameware = ParseDecimal(worksheet.Cells[row, 39].Text),
                                PowerBIPremium = ParseDecimal(worksheet.Cells[row, 40].Text),
                                PowerBIPro = ParseDecimal(worksheet.Cells[row, 41].Text),
                                BeltAnalyst = ParseDecimal(worksheet.Cells[row, 42].Text),
                                PDMC_Inventor_AutoCAD = ParseDecimal(worksheet.Cells[row, 43].Text),
                                Damware = ParseDecimal(worksheet.Cells[row, 44].Text),
                                SDP = ParseDecimal(worksheet.Cells[row, 45].Text),
                                SentinelOne = ParseDecimal(worksheet.Cells[row, 46].Text),
                                SAP_7_7 = ParseDecimal(worksheet.Cells[row, 47].Text),
                                GoTo = ParseDecimal(worksheet.Cells[row, 48].Text),
                                AutodeskVault = ParseDecimal(worksheet.Cells[row, 49].Text),
                                Pipenet = ParseDecimal(worksheet.Cells[row, 50].Text),
                                Mecastack = ParseDecimal(worksheet.Cells[row, 51].Text),
                            };
                            records.Add(record);
                        }
                    }
                    _context.IT_ITRecords.RemoveRange(_context.IT_ITRecords);
                    await _context.SaveChangesAsync();
                    _context.ChangeTracker.Clear(); // Important

                    // Ensure no duplicate EmailAddress in 'records' list
                    records = records
                        .GroupBy(static x => x.EmailAddress?.Trim().ToLower())
                        .Select(static g => g.First())
                        .ToList();

                    _context.IT_ITRecords.AddRange(records);
                    await _context.SaveChangesAsync();
                    _context.ChangeTracker.Clear();

                    int count = records.Count;
                    int recordCount = count;

                    return RedirectToAction("Index", new { message = $"IT Excel file uploaded successfully. {recordCount} records updated." });
                }
            }
            catch (Exception ex)
            {
                return RedirectToAction("Index", new { message = $"Error uploading IT Excel file: {ex.Message}" });
            }
        }

        return RedirectToAction("Index", new { message = "No IT file uploaded." });
    }

    private static decimal? ParseDecimal(object value)
    {
        if (value == null) return null;

        if (value is double doubleVal)
            return Math.Round(Convert.ToDecimal(doubleVal), 2);

        if (value is decimal decimalVal)
            return Math.Round(decimalVal, 2);

        if (decimal.TryParse(value.ToString(), out var result))
            return Math.Round(result, 2);

        return null;
    }

    [HttpPost]
    public async Task<IActionResult> UploadHRFile(IFormFile file)
    {
        if (file != null && file.Length > 0)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var newRecords = new List<ThirdPartyData>();

                for (int row = 2; row <= rowCount; row++)
                {
                    newRecords.Add(new ThirdPartyData
                    {
                        EmployeeCode = worksheet.Cells[row, 1].Text?.Trim() ?? string.Empty,
                        Employee8ID = worksheet.Cells[row, 2].Text?.Trim() ?? string.Empty,
                        ROLL = worksheet.Cells[row, 3].Text?.Trim() ?? string.Empty,
                        LOCATION_Sit_HO_Branch = worksheet.Cells[row, 4].Text?.Trim() ?? string.Empty,
                        DEPT = worksheet.Cells[row, 5].Text?.Trim() ?? string.Empty,
                        DEPT_NAME = worksheet.Cells[row, 6].Text?.Trim() ?? string.Empty,
                        Cost_Center = worksheet.Cells[row, 7].Text?.Trim() ?? string.Empty,
                        Salutation = worksheet.Cells[row, 8].Text?.Trim() ?? string.Empty,
                        Employee_Name = worksheet.Cells[row, 9].Text?.Trim() ?? string.Empty,
                        NEW_DESIGNATION = worksheet.Cells[row, 10].Text?.Trim() ?? string.Empty,
                        GRADE = worksheet.Cells[row, 11].Text?.Trim() ?? string.Empty,
                        Date_of_Joining = DateTime.TryParse(worksheet.Cells[row, 12].Text?.Trim(), out var doj) ? doj : (DateTime?)null,
                        Reporting_Manager = worksheet.Cells[row, 13].Text?.Trim() ?? string.Empty,
                        Functional_Manager = worksheet.Cells[row, 14].Text?.Trim() ?? string.Empty,
                        CAT = worksheet.Cells[row, 15].Text?.Trim() ?? string.Empty,
                        Category_For_Leave_Valuation = worksheet.Cells[row, 16].Text?.Trim() ?? string.Empty,
                        Gender = worksheet.Cells[row, 17].Text?.Trim() ?? string.Empty,
                        Business_Unit = worksheet.Cells[row, 18].Text?.Trim() ?? string.Empty,
                        Email = worksheet.Cells[row, 19].Text?.Trim() ?? string.Empty,
                        Active = "Active"
                    });
                }
                newRecords = newRecords
                                .GroupBy(x => (x.Employee8ID?.Trim() + x.Email?.Trim().ToLower()))
                                .Select(g => g.First())
                                .ToList();

                var existingRecords = _context.IT_ThirdPartyRecords.AsNoTracking().ToList();
                var newKeys = newRecords.Select(x => x.Employee8ID?.Trim() + x.Email?.Trim().ToLower()).ToHashSet();

                int updatedOrAddedCount = 0;

                foreach (var record in newRecords)
                {
                    var exists = _context.IT_ThirdPartyRecords.Any(x => x.Employee8ID == record.Employee8ID && x.Email == record.Email);
                    if (!exists)
                    {
                        _context.IT_ThirdPartyRecords.Add(record);
                        updatedOrAddedCount++;
                    }
                    else
                    {
                        _context.Entry(record).State = EntityState.Modified;
                        updatedOrAddedCount++;
                    }
                }

                foreach (var existing in existingRecords)
                {
                    if (!newKeys.Contains(existing.Employee8ID?.Trim() + existing.Email?.Trim().ToLower()))
                    {
                        existing.Active = "Inactive";
                        _context.IT_ThirdPartyRecords.Update(existing);
                    }
                }

                await _context.SaveChangesAsync();

                return RedirectToAction("Index", new { message = $"Third Party File uploaded successfully. {updatedOrAddedCount} records updated." });
            }
        }

        return RedirectToAction("Index", new { message = "No file uploaded." });
    }
    [HttpPost]
    public async Task<IActionResult> UploadLicenseMasterFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return RedirectToAction("Index", new { message = "No file uploaded." });
        }

        var newRecords = new List<LicenseMaster>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    return RedirectToAction("Index", new { message = "Invalid Excel format." });
                }

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++) // Assuming header at row 1
                {
                    var licenseName = worksheet.Cells[row, 1]?.Text?.Trim();
                    var costText = worksheet.Cells[row, 2]?.Text?.Trim();

                    if (string.IsNullOrEmpty(licenseName) || !decimal.TryParse(costText, out decimal cost))
                        continue;

                    newRecords.Add(new LicenseMaster
                    {
                        LicenseName = licenseName,
                        CostPerUser = cost
                    });
                }
            }
        }

        newRecords = newRecords
                        .GroupBy(x => x.LicenseName.ToLower())
                        .Select(g => g.First())
                        .ToList();

        var existingRecords = _context.IT_LicenseMaster.AsNoTracking().ToList();
        var newKeys = newRecords.Select(x => x.LicenseName.ToLower()).ToHashSet();

        int updatedOrAddedCount = 0;

        foreach (var record in newRecords)
        {
            var exists = _context.IT_LicenseMaster.Any(x => x.LicenseName.ToLower() == record.LicenseName.ToLower());

            if (!exists)
            {
                _context.IT_LicenseMaster.Add(record);
                updatedOrAddedCount++;
            }
            else
            {
                _context.IT_LicenseMaster.Update(record); // Treat as replace (or you can merge fields if needed)
                updatedOrAddedCount++;
            }
        }

        foreach (var existing in existingRecords)
        {
            if (!newKeys.Contains(existing.LicenseName.ToLower()))
            {
                _context.IT_LicenseMaster.Remove(existing); // Optional: remove missing ones
            }
        }

        await _context.SaveChangesAsync();

        return RedirectToAction("Index", new { message = $"License Master file uploaded successfully. {updatedOrAddedCount} records processed." });
    }

    [HttpGet]
    public async Task<IActionResult> DownloadReport()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            // Sheet 1: INNER JOIN
            var worksheet1 = package.Workbook.Worksheets.Add("Matched Records");

            var joinedData = await (from hr in _context.IT_ThirdPartyRecords
                                    join it in _context.IT_ITRecords
                                    on hr.Email equals it.EmailAddress
                                    select new { hr, it }).ToListAsync();

            worksheet1.Cells["A1"].LoadFromCollection(joinedData.Select(static r => new
            {
                r.hr.Employee8ID,
                r.hr.Email,
                r.hr.EmployeeCode,
                r.hr.ROLL,
                r.hr.LOCATION_Sit_HO_Branch,
                r.hr.DEPT,
                r.hr.DEPT_NAME,
                r.hr.Cost_Center,
                r.hr.Salutation,
                r.hr.Employee_Name,
                r.hr.NEW_DESIGNATION,
                r.hr.GRADE,
                r.hr.Date_of_Joining,
                r.hr.Reporting_Manager,
                r.hr.Functional_Manager,
                r.hr.CAT,
                r.hr.Category_For_Leave_Valuation,
                r.hr.Gender,
                r.hr.Business_Unit,
                r.hr.Active,
                r.it.EmailAddress,
                r.it.EmpID,
                r.it.Name,
                r.it.Dept,
                r.it.BU,
                r.it.CostCenter,
                r.it.O365,
                r.it.Internet,
                r.it.Servers,
                r.it.VisualStudio,
                r.it.TeklaStructure,
                r.it.Tally,
                r.it.StadProConnect,
                r.it.SQLServerStdCore,
                r.it.SimuTherm,
                r.it.SAPCrystalReport,
                r.it.ProE,
                r.it.PrimaVera,
                r.it.VaultProfessional,
                r.it.OracleStandard,
                r.it.NeiNastran,
                r.it.NavisworkManage,
                r.it.NavisworkFreedom,
                r.it.MSVisioProf,
                r.it.MSProjectProf,
                r.it.Femap_NXNastran,
                r.it.Ceaser,
                r.it.ChemCAD,
                r.it.DSignPdfSigner,
                r.it.PDMC,
                r.it.AEC,
                r.it.MasterCAM219,
                r.it.Mill3D,
                r.it.STARCCM,
                r.it.EEC,
                r.it.ThinkCell,
                r.it.EfficientElement,
                r.it.PDFExchangePro,
                r.it.Dameware,
                r.it.PowerBIPremium,
                r.it.PowerBIPro,
                r.it.BeltAnalyst,
                r.it.PDMC_Inventor_AutoCAD,
                r.it.Damware,
                r.it.SDP,
                r.it.SentinelOne,
                r.it.SAP_7_7,
                r.it.GoTo,
                r.it.AutodeskVault,
                r.it.Pipenet,
                r.it.Mecastack
            }), true);

            // Sheet 2: IT LEFT JOIN HR WHERE HR.Email IS NULL
            var worksheet2 = package.Workbook.Worksheets.Add("Unmatched IT Records");

            var itOnlyRecords = await (from it in _context.IT_ITRecords
                                       join hr in _context.IT_ThirdPartyRecords
                                       on it.EmailAddress equals hr.Email into gj
                                       from subhr in gj.DefaultIfEmpty()
                                       where subhr == null
                                       select it).ToListAsync();

            worksheet2.Cells["A1"].LoadFromCollection(itOnlyRecords, true);

            // Sheet 3: IT Usage Summary by BU
            var worksheet3 = package.Workbook.Worksheets.Add("BU-Wise Report");

            var usageSummary = await _context.IT_ITRecords
                .GroupBy(static it => it.BU)
                .Select(static g => new
                {
                    BU = g.Key ?? "Not Available",

                    O365_UserCount = g.Count(static x => x.O365 > 0),
                    O365_TotalCost = g.Sum(static x => x.O365) ?? 0,

                    Internet_UserCount = g.Count(static x => x.Internet > 0),
                    Internet_TotalCost = g.Sum(static x => (decimal?)x.Internet) ?? 0,

                    Servers_UserCount = g.Count(static x => x.Servers > 0),
                    Servers_TotalCost = g.Sum(static x => (decimal?)x.Servers) ?? 0,

                    VisualStudio_UserCount = g.Count(static x => x.VisualStudio > 0),
                    VisualStudio_TotalCost = g.Sum(static x => (decimal?)x.VisualStudio) ?? 0,

                    TeklaStructure_UserCount = g.Count(static x => x.TeklaStructure > 0),
                    TeklaStructure_TotalCost = g.Sum(static x => (decimal?)x.TeklaStructure) ?? 0,

                    Tally_UserCount = g.Count(static x => x.Tally > 0),
                    Tally_TotalCost = g.Sum(static x => (decimal?)x.Tally) ?? 0,

                    StadProConnect_UserCount = g.Count(static x => x.StadProConnect > 0),
                    StadProConnect_TotalCost = g.Sum(static x => (decimal?)x.StadProConnect) ?? 0,

                    SQLServerStdCore_UserCount = g.Count(static x => x.SQLServerStdCore > 0),
                    SQLServerStdCore_TotalCost = g.Sum(static x => (decimal?)x.SQLServerStdCore) ?? 0,

                    SimuTherm_UserCount = g.Count(static x => x.SimuTherm > 0),
                    SimuTherm_TotalCost = g.Sum(static x => (decimal?)x.SimuTherm) ?? 0,

                    SAPCrystalReport_UserCount = g.Count(static x => x.SAPCrystalReport > 0),
                    SAPCrystalReport_TotalCost = g.Sum(static x => (decimal?)x.SAPCrystalReport) ?? 0,

                    ProE_UserCount = g.Count(static x => x.ProE > 0),
                    ProE_TotalCost = g.Sum(static x => (decimal?)x.ProE) ?? 0,

                    PrimaVera_UserCount = g.Count(static x => x.PrimaVera > 0),
                    PrimaVera_TotalCost = g.Sum(static x => (decimal?)x.PrimaVera) ?? 0,

                    VaultProfessional_UserCount = g.Count(static x => x.VaultProfessional > 0),
                    VaultProfessional_TotalCost = g.Sum(static x => (decimal?)x.VaultProfessional) ?? 0,

                    OracleStandard_UserCount = g.Count(static x => x.OracleStandard > 0),
                    OracleStandard_TotalCost = g.Sum(static x => (decimal?)x.OracleStandard) ?? 0,

                    NeiNastran_UserCount = g.Count(static x => x.NeiNastran > 0),
                    NeiNastran_TotalCost = g.Sum(static x => (decimal?)x.NeiNastran) ?? 0,

                    NavisworkManage_UserCount = g.Count(static x => x.NavisworkManage > 0),
                    NavisworkManage_TotalCost = g.Sum(static x => (decimal?)x.NavisworkManage) ?? 0,

                    NavisworkFreedom_UserCount = g.Count(static x => x.NavisworkFreedom > 0),
                    NavisworkFreedom_TotalCost = g.Sum(static x => (decimal?)x.NavisworkFreedom) ?? 0,

                    MSVisioProf_UserCount = g.Count(static x => x.MSVisioProf > 0),
                    MSVisioProf_TotalCost = g.Sum(static x => (decimal?)x.MSVisioProf) ?? 0,

                    MSProjectProf_UserCount = g.Count(static x => x.MSProjectProf > 0),
                    MSProjectProf_TotalCost = g.Sum(static x => (decimal?)x.MSProjectProf) ?? 0,

                    Femap_NXNastran_UserCount = g.Count(static x => x.Femap_NXNastran > 0),
                    Femap_NXNastran_TotalCost = g.Sum(static x => (decimal?)x.Femap_NXNastran) ?? 0,

                    Ceaser_UserCount = g.Count(static x => x.Ceaser > 0),
                    Ceaser_TotalCost = g.Sum(static x => (decimal?)x.Ceaser) ?? 0,

                    ChemCAD_UserCount = g.Count(static x => x.ChemCAD > 0),
                    ChemCAD_TotalCost = g.Sum(static x => (decimal?)x.ChemCAD) ?? 0,

                    DSignPdfSigner_UserCount = g.Count(static x => x.DSignPdfSigner > 0),
                    DSignPdfSigner_TotalCost = g.Sum(static x => (decimal?)x.DSignPdfSigner) ?? 0,

                    PDMC_UserCount = g.Count(static x => x.PDMC > 0),
                    PDMC_TotalCost = g.Sum(static x => (decimal?)x.PDMC) ?? 0,

                    AEC_UserCount = g.Count(static x => x.AEC > 0),
                    AEC_TotalCost = g.Sum(static x => (decimal?)x.AEC) ?? 0,

                    MasterCAM219_UserCount = g.Count(static x => x.MasterCAM219 > 0),
                    MasterCAM219_TotalCost = g.Sum(static x => (decimal?)x.MasterCAM219) ?? 0,

                    Mill3D_UserCount = g.Count(static x => x.Mill3D > 0),
                    Mill3D_TotalCost = g.Sum(static x => (decimal?)x.Mill3D) ?? 0,

                    STARCCM_UserCount = g.Count(static x => x.STARCCM > 0),
                    STARCCM_TotalCost = g.Sum(static x => (decimal?)x.STARCCM) ?? 0,

                    EEC_UserCount = g.Count(static x => x.EEC > 0),
                    EEC_TotalCost = g.Sum(static x => (decimal?)x.EEC) ?? 0,

                    ThinkCell_UserCount = g.Count(static x => x.ThinkCell > 0),
                    ThinkCell_TotalCost = g.Sum(static x => (decimal?)x.ThinkCell) ?? 0,

                    EfficientElement_UserCount = g.Count(static x => x.EfficientElement > 0),
                    EfficientElement_TotalCost = g.Sum(static x => (decimal?)x.EfficientElement) ?? 0,

                    PDFExchangePro_UserCount = g.Count(static x => x.PDFExchangePro > 0),
                    PDFExchangePro_TotalCost = g.Sum(static x => (decimal?)x.PDFExchangePro) ?? 0,

                    Dameware_UserCount = g.Count(static x => x.Dameware > 0),
                    Dameware_TotalCost = g.Sum(static x => (decimal?)x.Dameware) ?? 0,

                    PowerBIPremium_UserCount = g.Count(static x => x.PowerBIPremium > 0),
                    PowerBIPremium_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPremium) ?? 0,

                    PowerBIPro_UserCount = g.Count(static x => x.PowerBIPro > 0),
                    PowerBIPro_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPro) ?? 0,

                    BeltAnalyst_UserCount = g.Count(static x => x.BeltAnalyst > 0),
                    BeltAnalyst_TotalCost = g.Sum(static x => (decimal?)x.BeltAnalyst) ?? 0,

                    PDMC_Inventor_AutoCAD_UserCount = g.Count(static x => x.PDMC_Inventor_AutoCAD > 0),
                    PDMC_Inventor_AutoCAD_TotalCost = g.Sum(static x => (decimal?)x.PDMC_Inventor_AutoCAD) ?? 0,

                    Damware_UserCount = g.Count(static x => x.Damware > 0),
                    Damware_TotalCost = g.Sum(static x => (decimal?)x.Damware) ?? 0,

                    SDP_UserCount = g.Count(static x => x.SDP > 0),
                    SDP_TotalCost = g.Sum(static x => (decimal?)x.SDP) ?? 0,

                    SentinelOne_UserCount = g.Count(static x => x.SentinelOne > 0),
                    SentinelOne_TotalCost = g.Sum(static x => (decimal?)x.SentinelOne) ?? 0,

                    SAP_7_7_UserCount = g.Count(static x => x.SAP_7_7 > 0),
                    SAP_7_7_TotalCost = g.Sum(static x => (decimal?)x.SAP_7_7) ?? 0,

                    GoTo_UserCount = g.Count(static x => x.GoTo > 0),
                    GoTo_TotalCost = g.Sum(static x => (decimal?)x.GoTo) ?? 0,

                    AutodeskVault_UserCount = g.Count(static x => x.AutodeskVault > 0),
                    AutodeskVault_TotalCost = g.Sum(static x => (decimal?)x.AutodeskVault) ?? 0,

                    Pipenet_UserCount = g.Count(static x => x.Pipenet > 0),
                    Pipenet_TotalCost = g.Sum(static x => (decimal?)x.Pipenet) ?? 0,

                    Mecastack_UserCount = g.Count(static x => x.Mecastack > 0),
                    Mecastack_TotalCost = g.Sum(static x => (decimal?)x.Mecastack) ?? 0
                })
                .OrderBy(static x => x.BU)
                .ToListAsync();

            worksheet3.Cells["A1"].LoadFromCollection(usageSummary, true);

            // Sheet 4: IT Usage Summary by CostCenter
            var worksheet4 = package.Workbook.Worksheets.Add("Cost center wise report");

            var costCenterSummary = await _context.IT_ITRecords
                .GroupBy(static it => it.CostCenter)
                .Select(static g => new
                {
                    CostCenter = g.Key ?? "Not Available",

                    O365_UserCount = g.Count(static x => x.O365 > 0),
                    O365_TotalCost = g.Sum(static x => (decimal?)x.O365) ?? 0,

                    Internet_UserCount = g.Count(static x => x.Internet > 0),
                    Internet_TotalCost = g.Sum(static x => (decimal?)x.Internet) ?? 0,

                    Servers_UserCount = g.Count(static x => x.Servers > 0),
                    Servers_TotalCost = g.Sum(static x => (decimal?)x.Servers) ?? 0,

                    VisualStudio_UserCount = g.Count(static x => x.VisualStudio > 0),
                    VisualStudio_TotalCost = g.Sum(static x => (decimal?)x.VisualStudio) ?? 0,

                    TeklaStructure_UserCount = g.Count(static x => x.TeklaStructure > 0),
                    TeklaStructure_TotalCost = g.Sum(static x => (decimal?)x.TeklaStructure) ?? 0,

                    Tally_UserCount = g.Count(static x => x.Tally > 0),
                    Tally_TotalCost = g.Sum(static x => (decimal?)x.Tally) ?? 0,

                    StadProConnect_UserCount = g.Count(static x => x.StadProConnect > 0),
                    StadProConnect_TotalCost = g.Sum(static x => (decimal?)x.StadProConnect) ?? 0,

                    SQLServerStdCore_UserCount = g.Count(static x => x.SQLServerStdCore > 0),
                    SQLServerStdCore_TotalCost = g.Sum(static x => (decimal?)x.SQLServerStdCore) ?? 0,

                    SimuTherm_UserCount = g.Count(static x => x.SimuTherm > 0),
                    SimuTherm_TotalCost = g.Sum(static x => (decimal?)x.SimuTherm) ?? 0,

                    SAPCrystalReport_UserCount = g.Count(static x => x.SAPCrystalReport > 0),
                    SAPCrystalReport_TotalCost = g.Sum(static x => (decimal?)x.SAPCrystalReport) ?? 0,

                    ProE_UserCount = g.Count(static x => x.ProE > 0),
                    ProE_TotalCost = g.Sum(static x => (decimal?)x.ProE) ?? 0,

                    PrimaVera_UserCount = g.Count(static x => x.PrimaVera > 0),
                    PrimaVera_TotalCost = g.Sum(static x => (decimal?)x.PrimaVera) ?? 0,

                    VaultProfessional_UserCount = g.Count(static x => x.VaultProfessional > 0),
                    VaultProfessional_TotalCost = g.Sum(static x => (decimal?)x.VaultProfessional) ?? 0,

                    OracleStandard_UserCount = g.Count(static x => x.OracleStandard > 0),
                    OracleStandard_TotalCost = g.Sum(static x => (decimal?)x.OracleStandard) ?? 0,

                    NeiNastran_UserCount = g.Count(static x => x.NeiNastran > 0),
                    NeiNastran_TotalCost = g.Sum(static x => (decimal?)x.NeiNastran) ?? 0,

                    NavisworkManage_UserCount = g.Count(static x => x.NavisworkManage > 0),
                    NavisworkManage_TotalCost = g.Sum(static x => (decimal?)x.NavisworkManage) ?? 0,

                    NavisworkFreedom_UserCount = g.Count(static x => x.NavisworkFreedom > 0),
                    NavisworkFreedom_TotalCost = g.Sum(static x => (decimal?)x.NavisworkFreedom) ?? 0,

                    MSVisioProf_UserCount = g.Count(static x => x.MSVisioProf > 0),
                    MSVisioProf_TotalCost = g.Sum(static x => (decimal?)x.MSVisioProf) ?? 0,

                    MSProjectProf_UserCount = g.Count(static x => x.MSProjectProf > 0),
                    MSProjectProf_TotalCost = g.Sum(static x => (decimal?)x.MSProjectProf) ?? 0,

                    Femap_NXNastran_UserCount = g.Count(static x => x.Femap_NXNastran > 0),
                    Femap_NXNastran_TotalCost = g.Sum(static x => (decimal?)x.Femap_NXNastran) ?? 0,

                    Ceaser_UserCount = g.Count(static x => x.Ceaser > 0),
                    Ceaser_TotalCost = g.Sum(static x => (decimal?)x.Ceaser) ?? 0,

                    ChemCAD_UserCount = g.Count(static x => x.ChemCAD > 0),
                    ChemCAD_TotalCost = g.Sum(static x => (decimal?)x.ChemCAD) ?? 0,

                    DSignPdfSigner_UserCount = g.Count(static x => x.DSignPdfSigner > 0),
                    DSignPdfSigner_TotalCost = g.Sum(static x => (decimal?)x.DSignPdfSigner) ?? 0,

                    PDMC_UserCount = g.Count(static x => x.PDMC > 0),
                    PDMC_TotalCost = g.Sum(static x => (decimal?)x.PDMC) ?? 0,

                    AEC_UserCount = g.Count(static x => x.AEC > 0),
                    AEC_TotalCost = g.Sum(static x => (decimal?)x.AEC) ?? 0,

                    MasterCAM219_UserCount = g.Count(static x => x.MasterCAM219 > 0),
                    MasterCAM219_TotalCost = g.Sum(static x => (decimal?)x.MasterCAM219) ?? 0,

                    Mill3D_UserCount = g.Count(static x => x.Mill3D > 0),
                    Mill3D_TotalCost = g.Sum(static x => (decimal?)x.Mill3D) ?? 0,

                    STARCCM_UserCount = g.Count(static x => x.STARCCM > 0),
                    STARCCM_TotalCost = g.Sum(static x => (decimal?)x.STARCCM) ?? 0,

                    EEC_UserCount = g.Count(static x => x.EEC > 0),
                    EEC_TotalCost = g.Sum(static x => (decimal?)x.EEC) ?? 0,

                    ThinkCell_UserCount = g.Count(static x => x.ThinkCell > 0),
                    ThinkCell_TotalCost = g.Sum(static x => (decimal?)x.ThinkCell) ?? 0,

                    EfficientElement_UserCount = g.Count(static x => x.EfficientElement > 0),
                    EfficientElement_TotalCost = g.Sum(static x => (decimal?)x.EfficientElement) ?? 0,

                    PDFExchangePro_UserCount = g.Count(static x => x.PDFExchangePro > 0),
                    PDFExchangePro_TotalCost = g.Sum(static x => (decimal?)x.PDFExchangePro) ?? 0,

                    Dameware_UserCount = g.Count(static x => x.Dameware > 0),
                    Dameware_TotalCost = g.Sum(static x => (decimal?)x.Dameware) ?? 0,

                    PowerBIPremium_UserCount = g.Count(static x => x.PowerBIPremium > 0),
                    PowerBIPremium_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPremium) ?? 0,

                    PowerBIPro_UserCount = g.Count(static x => x.PowerBIPro > 0),
                    PowerBIPro_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPro) ?? 0,

                    BeltAnalyst_UserCount = g.Count(static x => x.BeltAnalyst > 0),
                    BeltAnalyst_TotalCost = g.Sum(static x => (decimal?)x.BeltAnalyst) ?? 0,

                    PDMC_Inventor_AutoCAD_UserCount = g.Count(static x => x.PDMC_Inventor_AutoCAD > 0),
                    PDMC_Inventor_AutoCAD_TotalCost = g.Sum(static x => (decimal?)x.PDMC_Inventor_AutoCAD) ?? 0,

                    Damware_UserCount = g.Count(static x => x.Damware > 0),
                    Damware_TotalCost = g.Sum(static x => (decimal?)x.Damware) ?? 0,

                    SDP_UserCount = g.Count(static x => x.SDP > 0),
                    SDP_TotalCost = g.Sum(static x => (decimal?)x.SDP) ?? 0,

                    SentinelOne_UserCount = g.Count(static x => x.SentinelOne > 0),
                    SentinelOne_TotalCost = g.Sum(static x => (decimal?)x.SentinelOne) ?? 0,

                    SAP_7_7_UserCount = g.Count(static x => x.SAP_7_7 > 0),
                    SAP_7_7_TotalCost = g.Sum(static x => (decimal?)x.SAP_7_7) ?? 0,

                    GoTo_UserCount = g.Count(static x => x.GoTo > 0),
                    GoTo_TotalCost = g.Sum(static x => (decimal?)x.GoTo) ?? 0,

                    AutodeskVault_UserCount = g.Count(static x => x.AutodeskVault > 0),
                    AutodeskVault_TotalCost = g.Sum(static x => (decimal?)x.AutodeskVault) ?? 0,

                    Pipenet_UserCount = g.Count(static x => x.Pipenet > 0),
                    Pipenet_TotalCost = g.Sum(static x => (decimal?)x.Pipenet) ?? 0,

                    Mecastack_UserCount = g.Count(static x => x.Mecastack > 0),
                    Mecastack_TotalCost = g.Sum(static x => (decimal?)x.Mecastack) ?? 0
                })
    .OrderBy(static x => x.CostCenter)
    .ToListAsync();

            worksheet4.Cells["A1"].LoadFromCollection(costCenterSummary, true);

            // Sheet 5: IT Usage Summary by Dept
            var worksheet5 = package.Workbook.Worksheets.Add("Dept wise report");

            var deptSummary = await _context.IT_ITRecords
                .GroupBy(static it => it.Dept)
                .Select(static g => new
                {
                    Dept = g.Key ?? "Not Available",

                    O365_UserCount = g.Count(static x => x.O365 > 0),
                    O365_TotalCost = g.Sum(static x => (decimal?)x.O365) ?? 0,

                    Internet_UserCount = g.Count(static x => x.Internet > 0),
                    Internet_TotalCost = g.Sum(static x => (decimal?)x.Internet) ?? 0,

                    Servers_UserCount = g.Count(static x => x.Servers > 0),
                    Servers_TotalCost = g.Sum(static x => (decimal?)x.Servers) ?? 0,

                    VisualStudio_UserCount = g.Count(static x => x.VisualStudio > 0),
                    VisualStudio_TotalCost = g.Sum(static x => (decimal?)x.VisualStudio) ?? 0,

                    TeklaStructure_UserCount = g.Count(static x => x.TeklaStructure > 0),
                    TeklaStructure_TotalCost = g.Sum(static x => (decimal?)x.TeklaStructure) ?? 0,

                    Tally_UserCount = g.Count(static x => x.Tally > 0),
                    Tally_TotalCost = g.Sum(static x => (decimal?)x.Tally) ?? 0,

                    StadProConnect_UserCount = g.Count(static x => x.StadProConnect > 0),
                    StadProConnect_TotalCost = g.Sum(static x => (decimal?)x.StadProConnect) ?? 0,

                    SQLServerStdCore_UserCount = g.Count(static x => x.SQLServerStdCore > 0),
                    SQLServerStdCore_TotalCost = g.Sum(static x => (decimal?)x.SQLServerStdCore) ?? 0,

                    SimuTherm_UserCount = g.Count(static x => x.SimuTherm > 0),
                    SimuTherm_TotalCost = g.Sum(static x => (decimal?)x.SimuTherm) ?? 0,

                    SAPCrystalReport_UserCount = g.Count(static x => x.SAPCrystalReport > 0),
                    SAPCrystalReport_TotalCost = g.Sum(static x => (decimal?)x.SAPCrystalReport) ?? 0,

                    ProE_UserCount = g.Count(static x => x.ProE > 0),
                    ProE_TotalCost = g.Sum(static x => (decimal?)x.ProE) ?? 0,

                    PrimaVera_UserCount = g.Count(static x => x.PrimaVera > 0),
                    PrimaVera_TotalCost = g.Sum(static x => (decimal?)x.PrimaVera) ?? 0,

                    VaultProfessional_UserCount = g.Count(static x => x.VaultProfessional > 0),
                    VaultProfessional_TotalCost = g.Sum(static x => (decimal?)x.VaultProfessional) ?? 0,

                    OracleStandard_UserCount = g.Count(static x => x.OracleStandard > 0),
                    OracleStandard_TotalCost = g.Sum(static x => (decimal?)x.OracleStandard) ?? 0,

                    NeiNastran_UserCount = g.Count(static x => x.NeiNastran > 0),
                    NeiNastran_TotalCost = g.Sum(static x => (decimal?)x.NeiNastran) ?? 0,

                    NavisworkManage_UserCount = g.Count(static x => x.NavisworkManage > 0),
                    NavisworkManage_TotalCost = g.Sum(static x => (decimal?)x.NavisworkManage) ?? 0,

                    NavisworkFreedom_UserCount = g.Count(static x => x.NavisworkFreedom > 0),
                    NavisworkFreedom_TotalCost = g.Sum(static x => (decimal?)x.NavisworkFreedom) ?? 0,

                    MSVisioProf_UserCount = g.Count(static x => x.MSVisioProf > 0),
                    MSVisioProf_TotalCost = g.Sum(static x => (decimal?)x.MSVisioProf) ?? 0,

                    MSProjectProf_UserCount = g.Count(static x => x.MSProjectProf > 0),
                    MSProjectProf_TotalCost = g.Sum(static x => (decimal?)x.MSProjectProf) ?? 0,

                    Femap_NXNastran_UserCount = g.Count(static x => x.Femap_NXNastran > 0),
                    Femap_NXNastran_TotalCost = g.Sum(static x => (decimal?)x.Femap_NXNastran) ?? 0,

                    Ceaser_UserCount = g.Count(static x => x.Ceaser > 0),
                    Ceaser_TotalCost = g.Sum(static x => (decimal?)x.Ceaser) ?? 0,

                    ChemCAD_UserCount = g.Count(static x => x.ChemCAD > 0),
                    ChemCAD_TotalCost = g.Sum(static x => (decimal?)x.ChemCAD) ?? 0,

                    DSignPdfSigner_UserCount = g.Count(static x => x.DSignPdfSigner > 0),
                    DSignPdfSigner_TotalCost = g.Sum(static x => (decimal?)x.DSignPdfSigner) ?? 0,

                    PDMC_UserCount = g.Count(static x => x.PDMC > 0),
                    PDMC_TotalCost = g.Sum(static x => (decimal?)x.PDMC) ?? 0,

                    AEC_UserCount = g.Count(static x => x.AEC > 0),
                    AEC_TotalCost = g.Sum(static x => (decimal?)x.AEC) ?? 0,

                    MasterCAM219_UserCount = g.Count(static x => x.MasterCAM219 > 0),
                    MasterCAM219_TotalCost = g.Sum(static x => (decimal?)x.MasterCAM219) ?? 0,

                    Mill3D_UserCount = g.Count(static x => x.Mill3D > 0),
                    Mill3D_TotalCost = g.Sum(static x => (decimal?)x.Mill3D) ?? 0,

                    STARCCM_UserCount = g.Count(static x => x.STARCCM > 0),
                    STARCCM_TotalCost = g.Sum(static x => (decimal?)x.STARCCM) ?? 0,

                    EEC_UserCount = g.Count(static x => x.EEC > 0),
                    EEC_TotalCost = g.Sum(static x => (decimal?)x.EEC) ?? 0,

                    ThinkCell_UserCount = g.Count(static x => x.ThinkCell > 0),
                    ThinkCell_TotalCost = g.Sum(static x => (decimal?)x.ThinkCell) ?? 0,

                    EfficientElement_UserCount = g.Count(static x => x.EfficientElement > 0),
                    EfficientElement_TotalCost = g.Sum(static x => (decimal?)x.EfficientElement) ?? 0,

                    PDFExchangePro_UserCount = g.Count(static x => x.PDFExchangePro > 0),
                    PDFExchangePro_TotalCost = g.Sum(static x => (decimal?)x.PDFExchangePro) ?? 0,

                    Dameware_UserCount = g.Count(static x => x.Dameware > 0),
                    Dameware_TotalCost = g.Sum(static x => (decimal?)x.Dameware) ?? 0,

                    PowerBIPremium_UserCount = g.Count(static x => x.PowerBIPremium > 0),
                    PowerBIPremium_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPremium) ?? 0,

                    PowerBIPro_UserCount = g.Count(static x => x.PowerBIPro > 0),
                    PowerBIPro_TotalCost = g.Sum(static x => (decimal?)x.PowerBIPro) ?? 0,

                    BeltAnalyst_UserCount = g.Count(static x => x.BeltAnalyst > 0),
                    BeltAnalyst_TotalCost = g.Sum(static x => (decimal?)x.BeltAnalyst) ?? 0,

                    PDMC_Inventor_AutoCAD_UserCount = g.Count(static x => x.PDMC_Inventor_AutoCAD > 0),
                    PDMC_Inventor_AutoCAD_TotalCost = g.Sum(static x => (decimal?)x.PDMC_Inventor_AutoCAD) ?? 0,

                    Damware_UserCount = g.Count(static x => x.Damware > 0),
                    Damware_TotalCost = g.Sum(static x => (decimal?)x.Damware) ?? 0,

                    SDP_UserCount = g.Count(static x => x.SDP > 0),
                    SDP_TotalCost = g.Sum(static x => (decimal?)x.SDP) ?? 0,

                    SentinelOne_UserCount = g.Count(static x => x.SentinelOne > 0),
                    SentinelOne_TotalCost = g.Sum(static x => (decimal?)x.SentinelOne) ?? 0,

                    SAP_7_7_UserCount = g.Count(static x => x.SAP_7_7 > 0),
                    SAP_7_7_TotalCost = g.Sum(static x => (decimal?)x.SAP_7_7) ?? 0,

                    GoTo_UserCount = g.Count(static x => x.GoTo > 0),
                    GoTo_TotalCost = g.Sum(static x => (decimal?)x.GoTo) ?? 0,

                    AutodeskVault_UserCount = g.Count(static x => x.AutodeskVault > 0),
                    AutodeskVault_TotalCost = g.Sum(static x => (decimal?)x.AutodeskVault) ?? 0,

                    Pipenet_UserCount = g.Count(static x => x.Pipenet > 0),
                    Pipenet_TotalCost = g.Sum(static x => (decimal?)x.Pipenet) ?? 0,

                    Mecastack_UserCount = g.Count(static x => x.Mecastack > 0),
                    Mecastack_TotalCost = g.Sum(static x => (decimal?)x.Mecastack) ?? 0
                })
                .OrderBy(static x => x.Dept)
                .ToListAsync();

            worksheet5.Cells["A1"].LoadFromCollection(deptSummary, true);
            // Worksheet 6 - Consolidated Report
            var worksheet6 = package.Workbook.Worksheets.Add("Consolidated Report");

            var itData = _context.IT_ITRecords
                .AsEnumerable() // We switch to LINQ-to-Objects for complex aggregations
                .Select(r => new
                {
                    r.BU,
                    r.CostCenter,
                    EmployeeName = r.Name,
                    O365 = r.O365 ?? 0,
                    Internet = r.Internet ?? 0,
                    Servers = r.Servers ?? 0,
                    VisualStudio = r.VisualStudio ?? 0,
                    TeklaStructure = r.TeklaStructure ?? 0,
                    Tally = r.Tally ?? 0,
                    StadProConnect = r.StadProConnect ?? 0,
                    SQLServerStdCore = r.SQLServerStdCore ?? 0,
                    SimuTherm = r.SimuTherm ?? 0,
                    SAPCrystalReport = r.SAPCrystalReport ?? 0,
                    ProE = r.ProE ?? 0,
                    PrimaVera = r.PrimaVera ?? 0,
                    VaultProfessional = r.VaultProfessional ?? 0,
                    OracleStandard = r.OracleStandard ?? 0,
                    NeiNastran = r.NeiNastran ?? 0,
                    NavisworkManage = r.NavisworkManage ?? 0,
                    NavisworkFreedom = r.NavisworkFreedom ?? 0,
                    MSVisioProf = r.MSVisioProf ?? 0,
                    MSProjectProf = r.MSProjectProf ?? 0,
                    Femap_NXNastran = r.Femap_NXNastran ?? 0,
                    Ceaser = r.Ceaser ?? 0,
                    ChemCAD = r.ChemCAD ?? 0,
                    DSignPdfSigner = r.DSignPdfSigner ?? 0,
                    PDMC = r.PDMC ?? 0,
                    AEC = r.AEC ?? 0,
                    MasterCAM219 = r.MasterCAM219 ?? 0,
                    Mill3D = r.Mill3D ?? 0,
                    STARCCM = r.STARCCM ?? 0,
                    EEC = r.EEC ?? 0,
                    ThinkCell = r.ThinkCell ?? 0,
                    EfficientElement = r.EfficientElement ?? 0,
                    PDFExchangePro = r.PDFExchangePro ?? 0,
                    Dameware = r.Dameware ?? 0,
                    PowerBIPremium = r.PowerBIPremium ?? 0,
                    PowerBIPro = r.PowerBIPro ?? 0,
                    BeltAnalyst = r.BeltAnalyst ?? 0,
                    PDMC_Inventor_AutoCAD = r.PDMC_Inventor_AutoCAD ?? 0,
                    Damware = r.Damware ?? 0,
                    SDP = r.SDP ?? 0,
                    SentinelOne = r.SentinelOne ?? 0,
                    SAP_7_7 = r.SAP_7_7 ?? 0,
                    GoTo = r.GoTo ?? 0,
                    AutodeskVault = r.AutodeskVault ?? 0,
                    Pipenet = r.Pipenet ?? 0,
                    Mecastack = r.Mecastack ?? 0
                })
                .Select(r => new
                {
                    r.BU,
                    r.CostCenter,
                    r.EmployeeName,
                    r.O365,
                    r.Internet,
                    r.Servers,
                    r.VisualStudio,
                    r.TeklaStructure,
                    r.Tally,
                    r.StadProConnect,
                    r.SQLServerStdCore,
                    r.SimuTherm,
                    r.SAPCrystalReport,
                    r.ProE,
                    r.PrimaVera,
                    r.VaultProfessional,
                    r.OracleStandard,
                    r.NeiNastran,
                    r.NavisworkManage,
                    r.NavisworkFreedom,
                    r.MSVisioProf,
                    r.MSProjectProf,
                    r.Femap_NXNastran,
                    r.Ceaser,
                    r.ChemCAD,
                    r.DSignPdfSigner,
                    r.PDMC,
                    r.AEC,
                    r.MasterCAM219,
                    r.Mill3D,
                    r.STARCCM,
                    r.EEC,
                    r.ThinkCell,
                    r.EfficientElement,
                    r.PDFExchangePro,
                    r.Dameware,
                    r.PowerBIPremium,
                    r.PowerBIPro,
                    r.BeltAnalyst,
                    r.PDMC_Inventor_AutoCAD,
                    r.Damware,
                    r.SDP,
                    r.SentinelOne,
                    r.SAP_7_7,
                    r.GoTo,
                    r.AutodeskVault,
                    r.Pipenet,
                    r.Mecastack,
                    TotalEmployeeCost = r.O365 + r.Internet + r.Servers + r.VisualStudio + r.TeklaStructure + r.Tally +
                        r.StadProConnect + r.SQLServerStdCore + r.SimuTherm + r.SAPCrystalReport + r.ProE + r.PrimaVera +
                        r.VaultProfessional + r.OracleStandard + r.NeiNastran + r.NavisworkManage + r.NavisworkFreedom +
                        r.MSVisioProf + r.MSProjectProf + r.Femap_NXNastran + r.Ceaser + r.ChemCAD + r.DSignPdfSigner +
                        r.PDMC + r.AEC + r.MasterCAM219 + r.Mill3D + r.STARCCM + r.EEC + r.ThinkCell + r.EfficientElement +
                        r.PDFExchangePro + r.Dameware + r.PowerBIPremium + r.PowerBIPro + r.BeltAnalyst + r.PDMC_Inventor_AutoCAD +
                        r.Damware + r.SDP + r.SentinelOne + r.SAP_7_7 + r.GoTo + r.AutodeskVault + r.Pipenet + r.Mecastack
                })
                .ToList();

            // Add TotalCostCenterCost and TotalBUCost
            var grouped = itData
                .GroupBy(x => new { x.BU, x.CostCenter })
                .ToDictionary(g => g.Key, g => g.Sum(x => x.TotalEmployeeCost));

            var buGrouped = itData
                .GroupBy(x => x.BU)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.TotalEmployeeCost));

            var dataWithTotals = itData
                .Select(x => new
                {
                    x.BU,
                    TotalBUCost = buGrouped[x.BU],
                    x.CostCenter,
                    TotalCostCenterCost = grouped[new { x.BU, x.CostCenter }],
                    x.EmployeeName,
                    x.TotalEmployeeCost,
                    x.O365,
                    x.Internet,
                    x.Servers,
                    x.VisualStudio,
                    x.TeklaStructure,
                    x.Tally,
                    x.StadProConnect,
                    x.SQLServerStdCore,
                    x.SimuTherm,
                    x.SAPCrystalReport,
                    x.ProE,
                    x.PrimaVera,
                    x.VaultProfessional,
                    x.OracleStandard,
                    x.NeiNastran,
                    x.NavisworkManage,
                    x.NavisworkFreedom,
                    x.MSVisioProf,
                    x.MSProjectProf,
                    x.Femap_NXNastran,
                    x.Ceaser,
                    x.ChemCAD,
                    x.DSignPdfSigner,
                    x.PDMC,
                    x.AEC,
                    x.MasterCAM219,
                    x.Mill3D,
                    x.STARCCM,
                    x.EEC,
                    x.ThinkCell,
                    x.EfficientElement,
                    x.PDFExchangePro,
                    x.Dameware,
                    x.PowerBIPremium,
                    x.PowerBIPro,
                    x.BeltAnalyst,
                    x.PDMC_Inventor_AutoCAD,
                    x.Damware,
                    x.SDP,
                    x.SentinelOne,
                    x.SAP_7_7,
                    x.GoTo,
                    x.AutodeskVault,
                    x.Pipenet,
                    x.Mecastack
                })
                .OrderBy(x => x.BU).ThenBy(x => x.CostCenter).ThenBy(x => x.EmployeeName)
                .ToList();

            // Add data to worksheet
            worksheet6.Cells[1, 1].LoadFromCollection(dataWithTotals, true);
            worksheet6.Cells.AutoFitColumns();


            // Return file
            using var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
        }
    }
}
