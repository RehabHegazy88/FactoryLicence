using Newtonsoft.Json;
using PdfExtractorRazor.Models;
using PdfExtractorRazor.Services;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using UglyToad.PdfPig;
using Formatting = Newtonsoft.Json.Formatting;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace PdfExtractorRazor.Services
{
    public class PdfExtractionService : IPdfExtractionService
    {
        private Dictionary<string, Regex> _patterns;
        private readonly ILogger<PdfExtractionService> _logger;
        private readonly IWebHostEnvironment _environment;

        public PdfExtractionService(ILogger<PdfExtractionService> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
            InitializePatterns();
        }

        private void InitializePatterns()
        {
            _patterns = new Dictionary<string, Regex>
            {
                // More flexible patterns to handle various formats
                ["CertificateNo"] = new Regex(@"CERTIFICATE\s+NO[:\s]*([A-Z0-9\-]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["EquipmentType"] = new Regex(@"EQUIPMENT[:\s]*([A-Z\s]+?)(?=\s*SERIAL\s+NO|\s*MANUFACTURER|\s*PROJECT)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["SerialNo"] = new Regex(@"SERIAL\s+NO[:\s]*([A-Z0-9\-]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["Manufacturer"] = new Regex(@"MANUFACTURER[:\s]*([A-Z\s&/]+?)(?=\s*MODEL|\s*SERIAL|\s*CERTIFICATE)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["ModelNo"] = new Regex(@"MODEL\s+NO[:\s]*([A-Z0-9\s/\-\(\)½""″]+?)(?=\s*CERTIFICATE|\s*MANUFACTURER|\s*SERIAL)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["CalibratedRange"] = new Regex(@"CALIBRATED\s+RANGE[:\s]*([0-9\-\s]+(?:mpa|psi|bar|inHg|kPa)?)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["AccuracyGrade"] = new Regex(@"ACCURACY\s+GRADE[:\s]*([A-Z0-9\s\(\)½""″]+?)(?=\s*RIG\s+NUMBER|\s*CERTIFICATE|\s*MANUFACTURER)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["CalibrationDate"] = new Regex(@"DATE\s+OF\s+CALIBRATION[:\s]*([0-9\-\/]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["NextCalDate"] = new Regex(@"RECOMMENDED\s+CALIBRATION\s+DATE[:\s]*([0-9\-\/]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["Location"] = new Regex(@"LOCATION[:\s]*([A-Z\s\(\),\-]+?)(?=\s*The\s+certificate|\s*ACCURACY|\s*STANDARD|\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["AcceptanceCriteria"] = new Regex(@"(\+\/\-[0-9\.]+\s*%[^.]*(?:OEM\s+Instructions)?)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
                ["RigNumber"] = new Regex(@"RIG\s+NUMBER[:\s]*([A-Z0-9\-]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline)
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

                // Automatically save to JSON file
                try
                {
                    var savedFilePath = await SaveToJsonFileAsync(result.Certificates);
                    result.SavedFilePath = savedFilePath;
                    _logger.LogInformation("JSON file saved to: {FilePath}", savedFilePath);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error saving JSON file");
                    result.Errors.Add($"Error saving JSON file: {ex.Message}");
                }
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

            // Log the extracted text for debugging
            _logger.LogDebug("Extracted PDF text: {Text}", text);

            return ExtractDataFromText(text);
        }

        private CalibrationCertificate ExtractDataFromText(string text)
        {
            var certificate = new CalibrationCertificate();

            // Clean up the text first
            text = CleanupExtractedText(text);

            // Extract basic information with improved logic
            certificate.CertificateNo = ExtractCertificateNumber(text);
            certificate.EquipmentType = ExtractEquipmentType(text);
            certificate.SerialNo = ExtractSerialNumber(text);
            certificate.Manufacturer = ExtractManufacturer(text);
            certificate.ModelNo = ExtractModelNumber(text);
            certificate.AccuracyGrade = ExtractAccuracyGrade(text);
            certificate.CalibrationDate = ExtractValue(text, "CalibrationDate") ?? "";
            certificate.NextCalDate = ExtractValue(text, "NextCalDate") ?? "";
            certificate.Location = ExtractLocation(text);
            certificate.AcceptanceCriteria = ExtractAcceptanceCriteria(text);

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

        private string CleanupExtractedText(string text)
        {
            // First, try to add spaces before known keywords to separate jumbled text
            var keywords = new[] { "CUSTOMER", "ADDRESS", "EQUIPMENT", "PHO-CC", "PRESSURE", "GAUGE",
                                   "SERIAL", "MANUFACTURER", "MODEL", "CERTIFICATE", "CALIBRATED", "RANGE" };

            foreach (var keyword in keywords)
            {
                // Add space before keyword if it's not already there
                text = Regex.Replace(text, $@"(?<![:\s])({keyword})", " $1", RegexOptions.IgnoreCase);
            }

            // Remove excessive whitespace and normalize line breaks
            text = Regex.Replace(text, @"\s+", " ");
            text = Regex.Replace(text, @":\s*:", ":");

            // Fix common concatenations
            text = text.Replace("CUSTOMERADDRESS", "CUSTOMER ADDRESS ");
            text = text.Replace("ADDRESSEQUIPMENT", "ADDRESS EQUIPMENT ");
            text = text.Replace("EQUIPMENTPHO-CC", "EQUIPMENT PHO-CC");
            text = text.Replace("PHO-CCPRESSURE", "PHO-CC PRESSURE");

            return text.Trim();
        }

        private string ExtractCertificateNumber(string text)
        {
            // First, try to find the exact PHO-CC-##### pattern anywhere in text
            var phoPattern = Regex.Match(text, @"(PHO-CC-\d+)", RegexOptions.IgnoreCase);
            if (phoPattern.Success)
            {
                return phoPattern.Groups[1].Value;
            }

            // Try to find the pattern with possible text interference
            var patterns = new[]
            {
                @"PHO-CC-(\d+)",  // Direct pattern
                @"(?:EQUIPMENT|ADDRESS)?.*?(PHO-CC-\d+)", // Pattern with possible prefix text
                @"CERTIFICATE\s+NO[:\s]*(?:.*?)(PHO-CC-\d+)", // After certificate no label
                @"CERTIFICATE.*?(PHO-CC-\d+)", // Anywhere after certificate word
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    // Get the last group which should contain PHO-CC-#####
                    var lastGroup = match.Groups[match.Groups.Count - 1];
                    if (lastGroup.Success && lastGroup.Value.StartsWith("PHO-CC-", StringComparison.OrdinalIgnoreCase))
                    {
                        return lastGroup.Value.ToUpper();
                    }
                }
            }

            // As a final fallback, try to extract just the number after finding PHO-CC
            var numberPattern = Regex.Match(text, @"PHO-CC.*?(\d{5,})", RegexOptions.IgnoreCase);
            if (numberPattern.Success)
            {
                return $"PHO-CC-{numberPattern.Groups[1].Value}";
            }

            return "";
        }

        private string ExtractEquipmentType(string text)
        {
            // Look for equipment type in specific contexts
            var patterns = new[]
            {
                @"EQUIPMENT[:\s]*([A-Z\s]+?)(?=\s*SERIAL\s+NO|\s*MANUFACTURER|\s*PROJECT|PHO-CC)",
                @"(?:EQUIPMENT|INSTRUMENT)[:\s]*([A-Z\s]+GAUGE[A-Z\s]*)",
                @"(PRESSURE\s+GAUGE)",
                @"EQUIPMENT[:\s]*([^:]+?)(?=\s*PROJECT|\s*SERIAL)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[1].Value.Trim();
                    if (value.Length > 3 && !value.Contains("PHO-CC") && !Regex.IsMatch(value, @"^\d+"))
                    {
                        return value;
                    }
                }
            }

            return "";
        }

        private string ExtractSerialNumber(string text)
        {
            var patterns = new[]
            {
                @"SERIAL\s+NO[:\s]*([A-Z0-9\-]+?)(?=\s*CALIBRATED|\s*RANGE|\s*MANUFACTURER)",
                @"SERIAL\s+NO[:\s]*(\d+)",
                @"S/N[:\s]*([A-Z0-9\-]+)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[1].Value.Trim();
                    if (!value.Contains("CALIBRATED") && value.Length < 20)
                    {
                        return value;
                    }
                }
            }

            return "";
        }

        private string ExtractManufacturer(string text)
        {
            // Look for NAGMAN specifically (common in your certificates)
            var nagmanMatch = Regex.Match(text, @"\b(NAGMAN)\b", RegexOptions.IgnoreCase);
            if (nagmanMatch.Success)
            {
                return nagmanMatch.Groups[1].Value.ToUpper();
            }

            // Try other manufacturer patterns
            var patterns = new[]
            {
                @"MANUFACTURER[:\s]*([A-Z\s&/]+?)(?=\s*MODEL|\s*SERIAL|\s*CERTIFICATE)",
                @"MFG[:\s]*([A-Z\s&/]+?)(?=\s*MODEL|\s*SERIAL)",
                @"MAKE[:\s]*([A-Z\s&/]+?)(?=\s*MODEL|\s*SERIAL)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[1].Value.Trim();
                    if (value.Length > 2 && !value.Contains("CERTIFICATE") && !value.Contains("MODEL") && !value.Contains("NO"))
                    {
                        return value;
                    }
                }
            }

            return "";
        }

        private string ExtractModelNumber(string text)
        {
            // Look for S10 specifically in this PDF
            var s10Match = Regex.Match(text, @"\b(S10)\b", RegexOptions.IgnoreCase);
            if (s10Match.Success)
            {
                return s10Match.Groups[1].Value.ToUpper();
            }

            // Look for other common model patterns, but exclude technical references
            var modelPatterns = new[]
            {
                @"\b([A-Z]\d{1,4})\b", // Pattern like S10, A15, B123, etc.
                @"\b(\d{2,4}[A-Z]{1,2})\b", // Pattern like 314A, 12BC, etc.
                @"\b([A-Z]{2}\d{1,3})\b", // Pattern like AB12, XY123, etc.
            };

            foreach (var pattern in modelPatterns)
            {
                var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    var value = match.Groups[1].Value.Trim().ToUpper();
                    // Exclude technical references and standards
                    if (!IsModelNumberExclusion(value))
                    {
                        return value;
                    }
                }
            }

            return "N/A";
        }

        private bool IsModelNumberExclusion(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            // Exclude technical standards and common false positives
            var exclusions = new[] {
                "R101", "R102", // OIML standards
                "EN837", "B401", // Technical standards  
                "QF48", "C230", // Certificate references
                "CC56", // Part of certificate number
                "PC13", "BOX1", // Address components
                "NO", "CC", "PC", "BOX", "PHO", "RIG", "HPU"
            };

            return exclusions.Any(ex => value.Contains(ex, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsValidModelNumber(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            // Valid model numbers should be short and contain both letters and numbers typically
            if (value.Length > 10) return false;

            // Must contain at least one letter or number
            if (!Regex.IsMatch(value, @"[A-Z0-9]", RegexOptions.IgnoreCase)) return false;

            // Exclude common words that are not model numbers
            var excludeWords = new[] { "REFERENCE", "STANDARD", "EQUIPMENT", "MANUFACTURER", "CERTIFICATE",
                                     "SERIAL", "CALIBRATED", "RANGE", "ACCURACY", "GRADE", "DATE", "LOCATION" };

            return !excludeWords.Any(word => value.Contains(word, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsCommonFalsePositive(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;

            // Common false positives to exclude
            var falsePositives = new[] { "NO", "CC", "PC", "BOX", "PHO", "RIG", "HPU", "REFERENCE", "STANDARD" };
            return falsePositives.Any(fp => value.Equals(fp, StringComparison.OrdinalIgnoreCase));
        }

        private string ExtractAccuracyGrade(string text)
        {
            var patterns = new[]
            {
                @"ACCURACY\s+GRADE[:\s]*([A-Z0-9\s\(\)½""″\./]+?)(?=\s*RIG\s+NUMBER|\s*CERTIFICATE|\s*$|\s*0324)",
                @"ACCURACY[:\s]*([A-Z0-9\s\(\)½""″\./]+?)(?=\s*RIG|\s*CERTIFICATE)",
                @"CLASS[:\s]*([A-Z0-9\s\(\)½""″\./]+?)(?=\s*RIG|\s*CERTIFICATE)",
                @"1A\s*\([^)]*\)" // Specific pattern for "1A (2 ½")" type
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[0].Value.Trim(); // Use the whole match for specific patterns
                    if (pattern.Contains("1A"))
                    {
                        return value;
                    }

                    value = match.Groups[1].Value.Trim();
                    if (value.Length > 0 && value.Length < 50 && !value.Equals("GRADE", StringComparison.OrdinalIgnoreCase))
                    {
                        return value;
                    }
                }
            }

            return "";
        }

        private string ExtractLocation(string text)
        {
            // Look for specific location patterns from your PDFs
            var locationPatterns = new[]
            {
                @"(AIR\s+TANK-?\d*\s+ENGINE\s+ROOM)", // AIR TANK-3 ENGINE ROOM
                @"(BRAKE\s+HPU[^:]*)", // BRAKE HPU variations
                @"(ENGINE\s+ROOM)", // ENGINE ROOM
                @"(RIG\s+FLOOR)", // RIG FLOOR
                @"(PUMP\s+WORK\s+PRESSURE[^:]*)", // PUMP WORK PRESSURE variations
            };

            foreach (var pattern in locationPatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim().ToUpper();
                }
            }

            // Try standard LOCATION field extraction
            var standardPattern = Regex.Match(text, @"LOCATION[:\s]*([A-Z\s\(\),\-\d]+?)(?=\s*The\s+certificate|\s*ACCURACY|\s*$|\d+\.?\d*$)", RegexOptions.IgnoreCase);
            if (standardPattern.Success)
            {
                var value = standardPattern.Groups[1].Value.Trim();
                if (value.Length > 3 && value.Length < 100)
                {
                    return value.ToUpper();
                }
            }

            return "";
        }

        private string ExtractAcceptanceCriteria(string text)
        {
            var patterns = new[]
            {
                @"(\+\/\-\s*[0-9\.]+\s*%\s*of\s*FS\s*and\s*as\s*per\s*OEM\s*Instructions)", // Full pattern
                @"(\+\/\-\s*[0-9\.]+\s*%[^.]*?OEM\s*Instructions)", // Flexible version
                @"Acceptance\s*Criteria[:\s]*([^\n\r]+)", // After acceptance criteria label
                @"(\+\/-[0-9\.]+\s*%\s*of\s*FS)", // Basic percentage pattern
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var value = match.Groups[1].Value.Trim();
                    // Clean up the extracted value
                    value = Regex.Replace(value, @"\s+", " "); // Normalize whitespace
                    if (value.Length > 5 && value.Contains("%"))
                    {
                        return value;
                    }
                }
            }

            return "";
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
            // Look for range patterns - prioritize CALIBRATED RANGE
            var rangePatterns = new[]
            {
                @"CALIBRATED\s+RANGE[:\s]*(\d+\-\d+)\s*(mpa|psi|bar|inHg|kPa)", // Standard pattern
                @"(\d+\-\d+)\s*(psi|mpa|bar|inHg|kPa)", // Any range with units
                @"RANGE[:\s]*(\d+\-\d+)\s*(mpa|psi|bar|inHg|kPa)" // Alternative
            };

            foreach (var pattern in rangePatterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var range = match.Groups[1].Value.Trim();
                    var units = match.Groups[2].Value.ToLower();

                    // Validate the range makes sense
                    var rangeParts = range.Split('-');
                    if (rangeParts.Length == 2 &&
                        int.TryParse(rangeParts[0], out int start) &&
                        int.TryParse(rangeParts[1], out int end) &&
                        start < end && end <= 1000) // Reasonable range check
                    {
                        certificate.Range = range;
                        certificate.Units = units;
                        return;
                    }
                }
            }

            // Fallback: look for units first, then find nearby range
            var unitsMatch = Regex.Match(text, @"\b(mpa|psi|bar|inHg|kPa)\b", RegexOptions.IgnoreCase);
            if (unitsMatch.Success)
            {
                certificate.Units = unitsMatch.Groups[1].Value.ToLower();

                // Look for range pattern near the units
                var nearbyText = GetTextAroundMatch(text, unitsMatch, 50);
                var rangeMatch = Regex.Match(nearbyText, @"(\d+\-\d+)");
                if (rangeMatch.Success)
                {
                    certificate.Range = rangeMatch.Groups[1].Value;
                }
            }
        }

        private string GetTextAroundMatch(string text, Match match, int contextLength)
        {
            int start = Math.Max(0, match.Index - contextLength);
            int end = Math.Min(text.Length, match.Index + match.Length + contextLength);
            return text.Substring(start, end - start);
        }

        private string ExtractMaxDeviation(string text)
        {
            // Look for the calibration results table and find the actual maximum deviation
            var tableMatch = Regex.Match(text, @"UP\s+DOWN\s+UP(.*?)(?=REMARKS|Name:|$)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (tableMatch.Success)
            {
                var tableData = tableMatch.Groups[1].Value;

                // Extract deviation values (typically the last column in each row)
                var rows = tableData.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                double maxDev = 0;

                foreach (var row in rows)
                {
                    var numbers = Regex.Matches(row, @"\b(\d+\.?\d*)\b");
                    if (numbers.Count >= 3) // Should have Applied, Measured, Deviation
                    {
                        var lastNumber = numbers[numbers.Count - 1].Groups[1].Value;
                        if (double.TryParse(lastNumber, out double deviation))
                        {
                            // Deviation values should be small (< 10 typically)
                            if (deviation < 10)
                            {
                                maxDev = Math.Max(maxDev, Math.Abs(deviation));
                            }
                        }
                    }
                }

                if (maxDev > 0)
                    return maxDev.ToString("F2");
            }

            // Alternative: look for explicit deviation mentions
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

            // Look for small decimal numbers that could be deviations
            var smallNumbers = Regex.Matches(text, @"\b([0-9]+\.00)\b");
            double maxFound = 0;
            foreach (Match match in smallNumbers)
            {
                if (double.TryParse(match.Groups[1].Value, out double value) && value < 10 && value > maxFound)
                {
                    maxFound = value;
                }
            }

            return maxFound > 0 ? maxFound.ToString("F2") : "0.00";
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
            // Clean up manufacturer
            certificate.Manufacturer = CleanField(certificate.Manufacturer);
            certificate.ModelNo = CleanField(certificate.ModelNo);
            certificate.EquipmentType = CleanField(certificate.EquipmentType).ToUpper();
            certificate.Location = CleanField(certificate.Location).ToUpper();
            certificate.AccuracyGrade = CleanField(certificate.AccuracyGrade);

            if (certificate.ModelNo.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrEmpty(certificate.ModelNo))
            {
                certificate.ModelNo = "N/A";
            }
        }

        private string CleanField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";

            // Remove colons and excessive whitespace
            field = field.Replace(":", "").Trim();
            field = Regex.Replace(field, @"\s+", " ");

            // Remove common unwanted suffixes/prefixes
            field = Regex.Replace(field, @"(?:CERTIFICATE|MODEL|SERIAL|MANUFACTURER)\s*$", "", RegexOptions.IgnoreCase);

            return field.Trim();
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

        public async Task<string> SaveToJsonFileAsync(List<CalibrationCertificate> certificates, string? customFileName = null, bool append = false)
        {
            // Create output directory if it doesn't exist
            var outputDir = Path.Combine("Exports");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Generate filename
            var fileName = customFileName ?? $"calibration_certificates.json";
            if (!fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".json";
            }

            var filePath = Path.Combine(outputDir, fileName);

            // Handle appending or creating new file
            if (append && File.Exists(filePath))
            {
                // Read existing content
                var existingContent = await File.ReadAllTextAsync(filePath);
                var existingCertificates = JsonConvert.DeserializeObject<List<CalibrationCertificate>>(existingContent) ?? new List<CalibrationCertificate>();

                // Add new certificates
                existingCertificates.AddRange(certificates);

                // Convert to JSON and save
                var jsonContent = JsonConvert.SerializeObject(existingCertificates, Formatting.Indented);
                await File.WriteAllTextAsync(filePath, jsonContent);
            }
            else
            {
                // Convert to JSON and save as new file
                var jsonContent = JsonConvert.SerializeObject(certificates, Formatting.Indented);
                await File.WriteAllTextAsync(filePath, jsonContent);
            }

            return filePath;
        }

        public async Task<List<CalibrationCertificate>> GetFromJsonFileAsync(string fileName)
        {
            var filePath = Path.Combine("Exports", fileName);

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The requested JSON file was not found.", filePath);
            }

            var jsonContent = await File.ReadAllTextAsync(filePath);
            return JsonConvert.DeserializeObject<List<CalibrationCertificate>>(jsonContent) ?? new List<CalibrationCertificate>();
        }
    }
}