


// Services/PdfExtractionService.cs
using PdfExtractorRazor.Models;
using PdfExtractorRazor.Services;
using System.Text.Json;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace PdfExtractorRazor.Services
{
    public class PdfExtractionService : IPdfExtractionService
    {
        private   Dictionary<string, Regex> _patterns;
        private readonly ILogger<PdfExtractionService> _logger;

        public PdfExtractionService(ILogger<PdfExtractionService> logger)
        {
            _logger = logger;
            InitializePatterns();
        }

        private void InitializePatterns()
        {
            _patterns = new Dictionary<string, Regex>
            {
                ["CertificateNo"] = new Regex(@"CERTIFICATE\s+NO\s*:?\s*([A-Z0-9\-]+)", RegexOptions.IgnoreCase),
                ["EquipmentType"] = new Regex(@"EQUIPMENT\s*:?\s*([A-Z\s]+?)(?=\s*MANUFACTURER|\s*Project)", RegexOptions.IgnoreCase),
                ["SerialNo"] = new Regex(@"SERIAL\s+NO\s*:?\s*([A-Z0-9\-]+)", RegexOptions.IgnoreCase),
                ["Manufacturer"] = new Regex(@"MANUFACTURER\s*:?\s*([A-Z\s/]+?)(?=\s*MODEL|\s*SERIAL)", RegexOptions.IgnoreCase),
                ["ModelNo"] = new Regex(@"MODEL\s+NO\s*:?\s*([A-Z0-9\s/]+?)(?=\s*SERIAL|\s*CERTIFICATE)", RegexOptions.IgnoreCase),
                ["CalibratedRange"] = new Regex(@"CALIBRATED\s+RANGE\s*:?\s*([0-9\-\s]+(?:psi|bar|inHg|kPa)?)", RegexOptions.IgnoreCase),
                ["AccuracyGrade"] = new Regex(@"ACCURACY\s+GRADE\s*:?\s*([A-Z0-9\s\(\)½""″]+)", RegexOptions.IgnoreCase),
                ["CalibrationDate"] = new Regex(@"DATE\s+OF\s+CALIBRATION\s*:?\s*([0-9\-\/]+)", RegexOptions.IgnoreCase),
                ["NextCalDate"] = new Regex(@"RECOMMENDED\s+CALIBRATION\s+DATE\s*:?\s*([0-9\-\/]+)", RegexOptions.IgnoreCase),
                ["Location"] = new Regex(@"LOCATION\s*:?\s*([A-Z\s\(\)]+?)(?=\s*ACCURACY|\s*STANDARD)", RegexOptions.IgnoreCase),
                ["AcceptanceCriteria"] = new Regex(@"(\+\/\-[0-9\.]+\s*%[^.]*(?:OEM\s+Instructions)?)", RegexOptions.IgnoreCase)
            };
        }

        public async Task<ExtractionResult> ExtractFromFilesAsync(IFormFileCollection files)
        {
            var result = new ExtractionResult();

            foreach (var file in files)
            {
                if (file.Length > 0 && Path.GetExtension(file.FileName).ToLower() == ".pdf")
                {
                    try
                    {
                        var certificate = await ExtractFromSingleFileAsync(file);
                        if (certificate != null)
                        {
                            result.Certificates.Add(certificate);
                        }
                        result.ProcessedFiles++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error processing file {FileName}", file.FileName);
                        result.Errors.Add($"Error processing {file.FileName}: {ex.Message}");
                    }
                }
                else
                {
                    result.Errors.Add($"Invalid file: {file.FileName} (must be PDF)");
                }
            }

            if (result.HasResults)
            {
                result.JsonOutput = ConvertToJson(result.Certificates);
            }

            return result;
        }

        public async Task<CalibrationCertificate?> ExtractFromSingleFileAsync(IFormFile file)
        {
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            stream.Position = 0;

            using var document = PdfDocument.Open(stream);
            if (document.NumberOfPages == 0)
                return null;

            var page = document.GetPage(1);
            var text = page.Text;

            return ExtractDataFromText(text);
        }

        private CalibrationCertificate ExtractDataFromText(string text)
        {
            var certificate = new CalibrationCertificate();

            // Extract basic information
            certificate.CertificateNo = ExtractValue(text, "CertificateNo") ?? "";
            certificate.EquipmentType = ExtractValue(text, "EquipmentType")?.Trim() ?? "";
            certificate.SerialNo = ExtractValue(text, "SerialNo") ?? "";
            certificate.Manufacturer = ExtractValue(text, "Manufacturer")?.Trim() ?? "";
            certificate.ModelNo = ExtractValue(text, "ModelNo")?.Trim() ?? "";
            certificate.AccuracyGrade = ExtractValue(text, "AccuracyGrade")?.Trim() ?? "";
            certificate.CalibrationDate = ExtractValue(text, "CalibrationDate") ?? "";
            certificate.NextCalDate = ExtractValue(text, "NextCalDate") ?? "";
            certificate.Location = ExtractValue(text, "Location")?.Trim() ?? "";
            certificate.AcceptanceCriteria = ExtractValue(text, "AcceptanceCriteria")?.Trim() ?? "";

            // Extract range and units
            ExtractRangeAndUnits(text, certificate);

            // Extract maximum deviation from calibration results
            certificate.MaxDeviation = ExtractMaxDeviation(text);

            // Determine status
            certificate.Status = DetermineStatus(certificate);

            // Clean up extracted data
            CleanupCertificateData(certificate);

            return certificate;
        }

        private string? ExtractValue(string text, string patternKey)
        {
            if (_patterns.TryGetValue(patternKey, out var pattern))
            {
                var match = pattern.Match(text);
                if (match.Success && match.Groups.Count > 1)
                {
                    return match.Groups[1].Value.Trim();
                }
            }
            return null;
        }

        private void ExtractRangeAndUnits(string text, CalibrationCertificate certificate)
        {
            var rangeMatch = _patterns["CalibratedRange"].Match(text);
            if (rangeMatch.Success)
            {
                var rangeText = rangeMatch.Groups[1].Value.Trim();

                var unitsPattern = new Regex(@"(psi|bar|inHg|kPa|Pa)", RegexOptions.IgnoreCase);
                var unitsMatch = unitsPattern.Match(rangeText);
                if (unitsMatch.Success)
                {
                    certificate.Units = unitsMatch.Groups[1].Value;
                    certificate.Range = rangeText.Replace(certificate.Units, "").Trim();
                }
                else
                {
                    var textUnitsMatch = unitsPattern.Match(text);
                    if (textUnitsMatch.Success)
                    {
                        certificate.Units = textUnitsMatch.Groups[1].Value;
                    }
                    certificate.Range = rangeText;
                }
            }

            if (string.IsNullOrEmpty(certificate.Range))
            {
                var altRangePattern = new Regex(@"([0-9]+\-[0-9]+)\s*(psi|bar|inHg|kPa)", RegexOptions.IgnoreCase);
                var altMatch = altRangePattern.Match(text);
                if (altMatch.Success)
                {
                    certificate.Range = altMatch.Groups[1].Value;
                    certificate.Units = altMatch.Groups[2].Value;
                }
            }
        }

        private string ExtractMaxDeviation(string text)
        {
            var deviationPattern = new Regex(@"Deviation[^0-9]*([0-9]+\.?[0-9]*)", RegexOptions.IgnoreCase);
            var matches = deviationPattern.Matches(text);

            if (matches.Count > 0)
            {
                double maxDev = 0;
                foreach (Match match in matches)
                {
                    if (double.TryParse(match.Groups[1].Value, out double deviation))
                    {
                        maxDev = Math.Max(maxDev, Math.Abs(deviation));
                    }
                }
                return maxDev.ToString("F2");
            }

            return "0.00";
        }

        private string DetermineStatus(CalibrationCertificate certificate)
        {
            if (!string.IsNullOrEmpty(certificate.AcceptanceCriteria) &&
                !string.IsNullOrEmpty(certificate.MaxDeviation))
            {
                var percentMatch = Regex.Match(certificate.AcceptanceCriteria, @"([0-9\.]+)\s*%");
                if (percentMatch.Success && double.TryParse(percentMatch.Groups[1].Value, out double allowedPercent))
                {
                    if (double.TryParse(certificate.MaxDeviation, out double maxDev))
                    {
                        return maxDev <= allowedPercent ? "PASS" : "FAIL";
                    }
                }
            }

            return "PASS";
        }

        private void CleanupCertificateData(CalibrationCertificate certificate)
        {
            certificate.Manufacturer = certificate.Manufacturer.Replace(":", "").Trim();
            certificate.ModelNo = certificate.ModelNo.Replace(":", "").Trim();
            certificate.EquipmentType = certificate.EquipmentType.Replace(":", "").Trim().ToUpper();
            certificate.Location = certificate.Location.Replace(":", "").Trim().ToUpper();

            if (certificate.ModelNo.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                certificate.ModelNo = "N/A";
            }
        }

        public string ConvertToJson(List<CalibrationCertificate> certificates)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            return JsonSerializer.Serialize(certificates, options);
        }
    }
}