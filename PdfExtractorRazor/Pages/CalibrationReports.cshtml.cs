using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Newtonsoft.Json;
using PdfExtractorRazor.Models;
using System.Text;
using ClosedXML.Excel;

namespace PdfExtractorRazor.Pages
{
    public class CalibrationReportsModel : PageModel
    {
        private readonly ILogger<CalibrationReportsModel> _logger;
        private readonly IWebHostEnvironment _environment;

        public CalibrationReportsModel(ILogger<CalibrationReportsModel> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        public List<CalibrationCertificate> AllCertificates { get; set; } = new();
        public List<JsonFileInfo> JsonFiles { get; set; } = new();
        public CalibrationStatistics Statistics { get; set; } = new();
        public string? ErrorMessage { get; set; }
        public string? SuccessMessage { get; set; }

        public async Task OnGetAsync()
        {
            await LoadAllCertificates();
        }

        public async Task<IActionResult> OnPostLoadAllAsync()
        {
            await LoadAllCertificates();
            if (AllCertificates.Any())
            {
                SuccessMessage = $"Successfully loaded {AllCertificates.Count} certificates";
            }
            return Page();
        }

        public async Task<IActionResult> OnGetDownloadExcelAsync()
        {
            try
            {
                await LoadAllCertificates();

                if (!AllCertificates.Any())
                {
                    TempData["ErrorMessage"] = "No data available to export";
                    return RedirectToPage();
                }

                var excelData = GenerateExcelFile(AllCertificates);
                var fileName = $"Calibration_Certificates_{DateTime.Now:yyyy-MM-dd}.xlsx";

                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating Excel file");
                TempData["ErrorMessage"] = "Error generating Excel file: " + ex.Message;
                return RedirectToPage();
            }
        }

        public async Task<IActionResult> OnGetCertificatesDataAsync()
        {
            try
            {
                await LoadAllCertificates();
                return new JsonResult(AllCertificates);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting certificates data");
                return new JsonResult(new { error = ex.Message });
            }
        }

        private async Task LoadAllCertificates()
        {
            try
            {
                var exportsPath = Path.Combine(_environment.ContentRootPath, "Exports");

                if (Directory.Exists(exportsPath))
                {
                    var jsonFiles = Directory.GetFiles(exportsPath, "*.json")
                                             .OrderByDescending(f => System.IO.File.GetCreationTime(f))
                                             .ToList();

                    JsonFiles = jsonFiles.Select(f => new JsonFileInfo
                    {
                        FileName = Path.GetFileName(f),
                        FilePath = f,
                        CreatedDate = System.IO.File.GetCreationTime(f),
                        FileSize = new FileInfo(f).Length
                    }).ToList();

                    var allCertificates = new List<CalibrationCertificate>();

                    foreach (var jsonFile in jsonFiles)
                    {
                        try
                        {
                            var jsonContent = await System.IO.File.ReadAllTextAsync(jsonFile);
                            var certificates = JsonConvert.DeserializeObject<List<CalibrationCertificate>>(jsonContent);
                            if (certificates != null)
                            {
                                allCertificates.AddRange(certificates);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Error reading JSON file {FileName}", Path.GetFileName(jsonFile));
                        }
                    }

                    AllCertificates = allCertificates;
                    Statistics = CalculateStatistics(allCertificates);
                }
                else
                {
                    _logger.LogWarning("Exports directory not found at {Path}", exportsPath);
                    ErrorMessage = $"Exports directory not found at {exportsPath}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading calibration reports");
                ErrorMessage = "Error loading calibration reports: " + ex.Message;
            }
        }

        private byte[] GenerateExcelFile(List<CalibrationCertificate> certificates)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Calibration Certificates");

            // Add headers
            var headers = new[]
            {
                "Certificate No", "Equipment Type", "Serial No", "Manufacturer", "Model No",
                "Range", "Units", "Accuracy Grade", "Calibration Date", "Next Cal Date",
                "Max Deviation", "Status", "Location","Acceptance Criteria"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // Add data
            for (int i = 0; i < certificates.Count; i++)
            {
                var cert = certificates[i];
                var row = i + 2;

                worksheet.Cell(row, 1).Value = cert.CertificateNo ?? "";
                worksheet.Cell(row, 2).Value = cert.EquipmentType ?? "";
                worksheet.Cell(row, 3).Value = cert.SerialNo ?? "";
                worksheet.Cell(row, 4).Value = cert.Manufacturer ?? "";
                worksheet.Cell(row, 5).Value = cert.ModelNo ?? "";
                worksheet.Cell(row, 6).Value = cert.Range ?? "";
                worksheet.Cell(row, 7).Value = cert.Units ?? "";
                worksheet.Cell(row, 8).Value = cert.AccuracyGrade ?? "";
                worksheet.Cell(row, 9).Value = cert.CalibrationDate ?? "";
                worksheet.Cell(row, 10).Value = cert.NextCalDate ?? "";
                worksheet.Cell(row, 11).Value = cert.MaxDeviation ?? "";

                var statusCell = worksheet.Cell(row, 12);
                statusCell.Value = cert.Status ?? "";
                if (cert.Status == "PASS")
                {
                    statusCell.Style.Font.FontColor = XLColor.Green;
                }
                else if (cert.Status == "FAIL")
                {
                    statusCell.Style.Font.FontColor = XLColor.Red;
                }

                worksheet.Cell(row, 13).Value = cert.Location ?? "";
                worksheet.Cell(row, 14).Value = cert.AcceptanceCriteria ?? "";

                // Add borders to all cells in the row
                for (int j = 1; j <= headers.Length; j++)
                {
                    worksheet.Cell(row, j).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            // Add summary sheet
            var summarySheet = workbook.Worksheets.Add("Summary");
            summarySheet.Cell(1, 1).Value = "Calibration Certificates Summary";
            summarySheet.Cell(1, 1).Style.Font.Bold = true;
            summarySheet.Cell(1, 1).Style.Font.FontSize = 16;

            summarySheet.Cell(3, 1).Value = "Total Certificates:";
            summarySheet.Cell(3, 2).Value = Statistics.TotalCertificates;
            summarySheet.Cell(4, 1).Value = "Pressure Gauges:";
            summarySheet.Cell(4, 2).Value = Statistics.PressureGauges;
            summarySheet.Cell(5, 1).Value = "Relief Valves:";
            summarySheet.Cell(5, 2).Value = Statistics.ReliefValves;
            summarySheet.Cell(6, 1).Value = "Pass Rate:";
            summarySheet.Cell(6, 2).Value = $"{Statistics.PassRate}%";

            summarySheet.Column(1).Width = 20;
            summarySheet.Column(2).Width = 15;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        public async Task<IActionResult> OnGetFileDataAsync(string fileName)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    return new JsonResult(new { error = "File name is required" });
                }

                var exportsPath = Path.Combine(_environment.ContentRootPath, "Exports");
                var filePath = Path.Combine(exportsPath, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return new JsonResult(new { error = "File not found" });
                }

                var jsonContent = await System.IO.File.ReadAllTextAsync(filePath);
                var certificates = JsonConvert.DeserializeObject<List<CalibrationCertificate>>(jsonContent);

                return new JsonResult((certificates ?? new List<CalibrationCertificate>()).Distinct());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading file {FileName}", fileName);
                return new JsonResult(new { error = ex.Message });
            }
        }

        public async Task<IActionResult> OnGetFilesListAsync()
        {
            try
            {
                var exportsPath = Path.Combine(_environment.ContentRootPath, "Exports");
                var filesList = new List<JsonFileInfo>();

                if (Directory.Exists(exportsPath))
                {
                    var jsonFiles = Directory.GetFiles(exportsPath, "*.json")
                                             .OrderByDescending(f => System.IO.File.GetCreationTime(f));

                    filesList = jsonFiles.Select(f => new JsonFileInfo
                    {
                        FileName = Path.GetFileName(f),
                        FilePath = f,
                        CreatedDate = System.IO.File.GetCreationTime(f),
                        FileSize = new FileInfo(f).Length
                    }).ToList();
                }

                return new JsonResult(filesList);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting files list");
                return new JsonResult(new { error = ex.Message });
            }
        }

        private CalibrationStatistics CalculateStatistics(List<CalibrationCertificate> certificates)
        {
            if (!certificates.Any())
            {
                return new CalibrationStatistics();
            }

            var manufacturers = certificates.Where(c => !string.IsNullOrEmpty(c.Manufacturer))
                                          .GroupBy(c => c.Manufacturer)
                                          .ToDictionary(g => g.Key, g => g.Count());

            var equipmentTypes = certificates.Where(c => !string.IsNullOrEmpty(c.EquipmentType))
                                           .GroupBy(c => c.EquipmentType)
                                           .ToDictionary(g => g.Key, g => g.Count());

            return new CalibrationStatistics
            {
                TotalCertificates = certificates.Count,
                PressureGauges = certificates.Count(c => c.EquipmentType?.Contains("GAUGE") == true),
                ReliefValves = certificates.Count(c => c.EquipmentType?.Contains("RELIEF") == true),
                PassCount = certificates.Count(c => c.Status == "PASS"),
                FailCount = certificates.Count(c => c.Status == "FAIL"),
                PassRate = certificates.Count > 0 ? Math.Round((double)certificates.Count(c => c.Status == "PASS") / certificates.Count * 100, 1) : 0,
                Manufacturers = manufacturers,
                EquipmentTypes = equipmentTypes
            };
        }
    }

    public class JsonFileInfo
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; }
        public long FileSize { get; set; }

        public string FileSizeFormatted => FormatFileSize(FileSize);

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }

    public class CalibrationStatistics
    {
        public int TotalCertificates { get; set; }
        public int PressureGauges { get; set; }
        public int ReliefValves { get; set; }
        public int PassCount { get; set; }
        public int FailCount { get; set; }
        public double PassRate { get; set; }
        public Dictionary<string, int> Manufacturers { get; set; } = new();
        public Dictionary<string, int> EquipmentTypes { get; set; } = new();
    }
}