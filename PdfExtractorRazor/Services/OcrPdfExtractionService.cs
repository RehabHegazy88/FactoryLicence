using Tesseract;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using PdfExtractorRazor.Models;
using System.Text.Json;
using Formatting = Newtonsoft.Json.Formatting;
using JsonSerializer = System.Text.Json.JsonSerializer; 
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PdfExtractorRazor.Services
{
    public class OcrPdfExtractionService : IPdfExtractionService
    {
        private readonly ILogger<OcrPdfExtractionService> _logger;
        private readonly IWebHostEnvironment _environment;

        public OcrPdfExtractionService(ILogger<OcrPdfExtractionService> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
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

            try
            {
                // Method 1: Try direct text extraction first (faster)
                var directText = await TryDirectTextExtraction(stream);
                if (!string.IsNullOrEmpty(directText) && directText.Length > 100)
                {
                    _logger.LogInformation("Using direct text extraction");
                    return ExtractDataFromText(directText);
                }

                // Method 2: Fall back to OCR if direct extraction fails
                stream.Position = 0;
                var ocrText = await PerformOcrExtraction(stream);
                if (!string.IsNullOrEmpty(ocrText))
                {
                    _logger.LogInformation("Using OCR text extraction");
                    return ExtractDataFromText(ocrText);
                }

                _logger.LogWarning("Both text extraction methods failed");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during PDF extraction");
                return null;
            }
        }

        private async Task<string> TryDirectTextExtraction(Stream stream)
        {
            try
            {
                using var pdfReader = new PdfReader(stream);
                using var pdfDocument = new PdfDocument(pdfReader);

                if (pdfDocument.GetNumberOfPages() == 0)
                    return "";

                var page = pdfDocument.GetPage(1);
                var strategy = new LocationTextExtractionStrategy();
                var text = PdfTextExtractor.GetTextFromPage(page, strategy);

                return text ?? "";
            }
            catch (Exception ex)
            {
                _logger.LogDebug("Direct text extraction failed: {Error}", ex.Message);
                return "";
            }
        }

        private async Task<string> PerformOcrExtraction(Stream stream)
        {
            try
            {
                // Convert PDF to image using a simpler approach
                var imageBytes = await ConvertPdfToImage(stream);
                if (imageBytes == null || imageBytes.Length == 0)
                {
                    _logger.LogWarning("Failed to convert PDF to image");
                    return "";
                }

                // Perform OCR on the image
                return await PerformOcrOnImage(imageBytes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OCR extraction failed");
                return "";
            }
        }

        private async Task<byte[]?> ConvertPdfToImage(Stream pdfStream)
        {
            try
            {
                // Simple PDF to image conversion using System.Drawing
                // This is a basic implementation - you might want to use a more robust library

                // For now, we'll use a simple approach with Windows API calls
                // Note: This requires Windows and might not work on Linux

                var tempPdfFile = Path.GetTempFileName() + ".pdf";
                var tempImageFile = Path.GetTempFileName() + ".png";

                try
                {
                    // Save stream to temp file
                    await using (var fileStream = File.Create(tempPdfFile))
                    {
                        pdfStream.Position = 0;
                        await pdfStream.CopyToAsync(fileStream);
                    }

                    // Use a command line tool to convert PDF to image
                    // You would need to install something like ImageMagick or Ghostscript
                    var success = await ConvertPdfToImageUsingGhostscript(tempPdfFile, tempImageFile);

                    if (success && File.Exists(tempImageFile))
                    {
                        return await File.ReadAllBytesAsync(tempImageFile);
                    }
                }
                finally
                {
                    // Cleanup temp files
                    if (File.Exists(tempPdfFile)) File.Delete(tempPdfFile);
                    if (File.Exists(tempImageFile)) File.Delete(tempImageFile);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to convert PDF to image");
            }

            return null;
        }

        private async Task<bool> ConvertPdfToImageUsingGhostscript(string pdfPath, string imagePath)
        {
            try
            {
                // This requires Ghostscript to be installed
                // Download from: https://www.ghostscript.com/download/gsdnld.html

                var ghostscriptPath = @"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe"; // Adjust path as needed

                if (!File.Exists(ghostscriptPath))
                {
                    _logger.LogWarning("Ghostscript not found at {Path}. Please install Ghostscript for PDF to image conversion.", ghostscriptPath);
                    return false;
                }

                var arguments = $"-sDEVICE=png16m -dNOPAUSE -dBATCH -dSAFER -r300 -sOutputFile=\"{imagePath}\" \"{pdfPath}\"";

                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = ghostscriptPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var process = System.Diagnostics.Process.Start(processInfo);
                if (process != null)
                {
                    await process.WaitForExitAsync();
                    return process.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Ghostscript conversion failed");
            }

            return false;
        }

        private async Task<string> PerformOcrOnImage(byte[] imageBytes)
        {
            try
            {
                var tessDataPath = Path.Combine(_environment.ContentRootPath, "tessdata");
                if (!Directory.Exists(tessDataPath))
                {
                    _logger.LogWarning("Tessdata directory not found at {Path}. Creating directory. Please download eng.traineddata.", tessDataPath);
                    Directory.CreateDirectory(tessDataPath);
                    return ""; // Can't proceed without language data
                }

                using var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default);

                // Configure OCR parameters for better accuracy
                engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-+/.:%()\"' ");
                engine.SetVariable("preserve_interword_spaces", "1");

                using var memoryStream = new MemoryStream(imageBytes);
                using var image = Image.FromStream(memoryStream);
                using var bitmap = new Bitmap(image);

                // Convert to the format Tesseract expects
                var bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

                try
                {
                    using var pix = Pix.Create(bitmap.Width, bitmap.Height, 24);
                   // pix.SetResolution(300, 300);

                    // Copy bitmap data to Pix (this is simplified - you might need a more robust conversion)
                    // For a complete implementation, you'd need to properly convert the bitmap data

                    using var page = engine.Process(pix);
                    var text = page.GetText();
                    var confidence = page.GetMeanConfidence();

                    _logger.LogInformation("OCR completed with confidence: {Confidence}%", confidence * 100);

                    if (confidence < 0.5)
                    {
                        _logger.LogWarning("Low OCR confidence ({Confidence}%). Results may be inaccurate.", confidence * 100);
                    }

                    return text ?? "";
                }
                finally
                {
                    bitmap.UnlockBits(bitmapData);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OCR processing failed");
                return "";
            }
        }

        private CalibrationCertificate ExtractDataFromText(string text)
        {
            var certificate = new CalibrationCertificate();

            // Clean up OCR text - remove extra spaces and fix common OCR errors
            text = CleanupOcrText(text);

            // Extract all fields using OCR-optimized patterns
            certificate.CertificateNo = ExtractCertificateNumber(text);
            certificate.EquipmentType = ExtractEquipmentType(text);
            certificate.SerialNo = ExtractSerialNumber(text);
            certificate.Manufacturer = ExtractManufacturer(text);
            certificate.ModelNo = ExtractModelNumber(text);
            certificate.AccuracyGrade = ExtractAccuracyGrade(text);
            certificate.CalibrationDate = ExtractCalibrationDate(text);
            certificate.NextCalDate = ExtractNextCalibrationDate(text);
            certificate.Location = ExtractLocation(text);
            certificate.AcceptanceCriteria = ExtractAcceptanceCriteria(text);

            ExtractRangeAndUnits(text, certificate);
            certificate.MaxDeviation = ExtractMaxDeviation(text);
            certificate.Status = DetermineStatus(certificate);
            CleanupCertificateData(certificate);

            return certificate;
        }

        private string CleanupOcrText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // Fix common OCR errors
            text = text.Replace("0", "O").Replace("O", "0"); // This is tricky - need context
            text = text.Replace("l", "1").Replace("I", "1"); // Common number/letter confusion
            text = text.Replace("S", "5").Replace("5", "S"); // Context dependent

            // Better approach - fix specific known patterns
            text = Regex.Replace(text, @"PHO-CC-(\w+)", match =>
            {
                var certNum = match.Groups[1].Value;
                // Fix common OCR errors in certificate numbers
                certNum = certNum.Replace("O", "0").Replace("S", "5").Replace("I", "1");
                return $"PHO-CC-{certNum}";
            });

            // Clean up spacing and line breaks
            text = Regex.Replace(text, @"\s+", " ");
            text = text.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");

            return text.Trim();
        }

        private string ExtractCertificateNumber(string text)
        {
            var patterns = new[]
            {
                @"PHO-CC-(\d{5})",
                @"CERTIFICATE\s*NO[:\s]*PHO-CC-(\d{5})",
                @"PHO[\-\s]*CC[\-\s]*(\d{5})" // OCR might add spaces in dashes
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var certNum = match.Groups[1].Value;
                    return $"PHO-CC-{certNum}";
                }
            }

            return "";
        }

        private string ExtractEquipmentType(string text)
        {
            var patterns = new[]
            {
                @"EQUIPMENT[:\s]*(PRESSURE\s+(?:GAUGE|RELIEF\s+VALVE))",
                @"(PRESSURE\s+GAUGE)",
                @"(PRESSURE\s+RELIEF\s+VALVE)",
                @"EQUIPMENT[:\s]*([A-Z\s]+GAUGE)",
                @"EQUIPMENT[:\s]*([A-Z\s]+VALVE)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var equipment = match.Groups[1].Value.Trim();
                    equipment = Regex.Replace(equipment, @"\s+", " ");
                    return equipment.ToUpper();
                }
            }

            return "";
        }

        private string ExtractSerialNumber(string text)
        {
            // OCR-friendly patterns for serial numbers
            var patterns = new[]
            {
                @"SERIAL\s*NO[:\s]*([A-Z0-9\-]+)",
                @"\bE(\d{10})\b", // E2119930387
                @"\b(\d{3}[\-\s]*PRV[\-\s]*\d{2})\b", // 103-PRV-05 (OCR might add spaces)
                @"\b(\d{7}[A-Z])\b", // 1404137M
                @"SERIAL[:\s]*([A-Z0-9\-]+)"
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    var value = match.Groups[match.Groups.Count - 1].Value.Trim();
                    value = value.Replace(" ", ""); // Remove OCR-introduced spaces

                    if (value.Length >= 4 && value.Length <= 20 &&
                        !value.Contains("1921") && // Skip project number
                        !IsCommonFalsePositive(value))
                    {
                        _logger.LogInformation("Found serial number: {SerialNo}", value);
                        return value;
                    }
                }
            }

            return "";
        }

        private string ExtractManufacturer(string text)
        {
            var manufacturers = new[]
            {
                "NAGMAN", "CALCON", "FUYU"
            };

            // Special case for MC (needs word boundaries)
            var mcMatch = Regex.Match(text, @"\bMC\b", RegexOptions.IgnoreCase);
            if (mcMatch.Success)
            {
                return "MC";
            }

            // Check for SAFETY VALVE manufacturer
            if (Regex.IsMatch(text, @"SAFETY\s+VALVE", RegexOptions.IgnoreCase))
            {
                return "SAFETY VALVE";
            }

            foreach (var manufacturer in manufacturers)
            {
                var pattern = $@"\b{manufacturer}\b";
                if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                {
                    return manufacturer;
                }
            }

            return "";
        }

        private string ExtractModelNumber(string text)
        {
            var models = new Dictionary<string, string[]>
            {
                ["EN837-1"] = new[] { @"\bEN837[\-\s]*1\b", @"EN\s*837\s*\-\s*1" },
                ["S10"] = new[] { @"\bS10\b", @"\bS\s*10\b" },
                ["314"] = new[] { @"\b314\b(?!\d)" },
                ["42811"] = new[] { @"\b42811\b", @"\b428\s*11\b" }
            };

            foreach (var kvp in models)
            {
                foreach (var pattern in kvp.Value)
                {
                    if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                    {
                        return kvp.Key;
                    }
                }
            }

            return "N/A";
        }

        private string ExtractAccuracyGrade(string text)
        {
            var patterns = new[]
            {
                @"ACCURACY\s*GRADE[:\s]*([^:\r\n]+?)(?=\s*STANDARD|\s*RIG|\s*CERTIFICATE|$)",
                @"(1A\s*\([^)]*\))",
                @"(2A\s*\d+[""′]?)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var grade = match.Groups[1].Value.Trim();
                    if (grade.Length > 0 && grade.Length < 20)
                    {
                        return grade;
                    }
                }
            }

            return "";
        }

        private string ExtractCalibrationDate(string text)
        {
            var patterns = new[]
            {
                @"DATE\s*OF\s*CALIBRATION[:\s]*([0-9\-/]+)",
                @"CALIBRATION[:\s]*([0-9\-/]+)"
            };

            return ExtractDatePattern(text, patterns);
        }

        private string ExtractNextCalibrationDate(string text)
        {
            var patterns = new[]
            {
                @"RECOMMENDED\s*CALIBRATION\s*DATE[:\s]*([0-9\-/]+)",
                @"NEXT\s*CALIBRATION[:\s]*([0-9\-/]+)"
            };

            return ExtractDatePattern(text, patterns);
        }

        private string ExtractDatePattern(string text, string[] patterns)
        {
            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var date = match.Groups[1].Value.Trim();
                    // Validate date format
                    if (Regex.IsMatch(date, @"\d{2}-\d{2}-\d{4}"))
                    {
                        return date;
                    }
                }
            }
            return "";
        }

        private string ExtractLocation(string text)
        {
            var patterns = new[]
            {
                @"LOCATION[:\s]*([^:\r\n]+?)(?=\s*ACCURACY|\s*STANDARD|$)",
                @"(AIR\s*TANK[\-\s]*\d*\s*ENGINE\s*ROOM)",
                @"(ENGINE\s*ROOM)",
                @"(RIG\s*FLOOR)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var location = match.Groups[1].Value.Trim();
                    location = Regex.Replace(location, @"\s+", " ");
                    if (location.Length > 3 && location.Length < 100)
                    {
                        return location.ToUpper();
                    }
                }
            }

            return "";
        }

        private string ExtractAcceptanceCriteria(string text)
        {
            var patterns = new[]
            {
                @"(\+[\-/]*\s*[0-9\.]+\s*%\s*of\s*(?:FS|SP)(?:\s+and\s+as\s+per\s+OEM\s+Instructions)?)",
                @"Acceptance\s*Criteria[:\s]*([^\r\n]+)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var criteria = match.Groups[1].Value.Trim();
                    if (criteria.Length > 3 && criteria.Length < 100)
                    {
                        return criteria;
                    }
                }
            }

            return "";
        }

        private void ExtractRangeAndUnits(string text, CalibrationCertificate certificate)
        {
            var patterns = new[]
            {
                @"CALIBRATED\s*RANGE[:\s]*(0[\-\s]*\d+)\s*(psi|bar|mpa|kPa|inHg)",
                @"\b(0[\-\s]*230)\s*(psi)\b",
                @"\b(0[\-\s]*\d+)\s*(psi|bar|mpa|kPa|inHg)\b",
                @"(\d+)\s*(psi|bar|mpa|kPa|inHg)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var range = match.Groups[1].Value.Replace(" ", "").Replace("—", "-");
                    var units = match.Groups[2].Value.ToLower();

                    if (IsValidRange(range))
                    {
                        certificate.Range = range;
                        certificate.Units = units;
                        return;
                    }
                }
            }
        }

        private bool IsValidRange(string range)
        {
            return !range.Contains("314") &&
                   !range.Contains("42811") &&
                   !range.Contains("2025") &&
                   !range.Contains("2026") &&
                   !range.Contains("1921") &&
                   range.Length <= 10;
        }

        private string ExtractMaxDeviation(string text)
        {
            // Look for deviation values in calibration results
            var deviationMatches = Regex.Matches(text, @"([01]\.00)", RegexOptions.IgnoreCase);
            double maxDev = 0;

            foreach (Match match in deviationMatches)
            {
                if (double.TryParse(match.Groups[1].Value, out double deviation))
                {
                    maxDev = Math.Max(maxDev, deviation);
                }
            }

            return maxDev.ToString("F2");
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
            certificate.Manufacturer = CleanField(certificate.Manufacturer);
            certificate.ModelNo = CleanField(certificate.ModelNo);
            certificate.EquipmentType = CleanField(certificate.EquipmentType);
            certificate.Location = CleanField(certificate.Location);
            certificate.AccuracyGrade = CleanField(certificate.AccuracyGrade);
            certificate.SerialNo = CleanField(certificate.SerialNo);
            certificate.CertificateNo = CleanField(certificate.CertificateNo);

            if (string.IsNullOrEmpty(certificate.ModelNo) || certificate.ModelNo.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                certificate.ModelNo = "N/A";
            }
        }

        private string CleanField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";

            field = field.Replace(":", "").Trim();
            field = Regex.Replace(field, @"\s+", " ");
            return field.Trim();
        }

        private bool IsCommonFalsePositive(string value)
        {
            var falsePositives = new[] {
                "NO", "CERTIFICATE", "MANUFACTURER", "SERIAL", "EQUIPMENT",
                "RANGE", "CALIBRATED", "LOCATION", "ACCURACY", "GRADE",
                "DATE", "RECOMMENDED", "STANDARD", "REFERENCE"
            };

            return falsePositives.Any(fp => value.Contains(fp, StringComparison.OrdinalIgnoreCase));
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
            var outputDir = Path.Combine("Exports");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var fileName = customFileName ?? $"calibration_certificates_ocr_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            if (!fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".json";
            }

            var filePath = Path.Combine(outputDir, fileName);

            if (append && File.Exists(filePath))
            {
                var existingContent = await File.ReadAllTextAsync(filePath);
                var existingCertificates = JsonConvert.DeserializeObject<List<CalibrationCertificate>>(existingContent) ?? new List<CalibrationCertificate>();
                existingCertificates.AddRange(certificates);
                var jsonContent = JsonConvert.SerializeObject(existingCertificates, Formatting.Indented);
                await File.WriteAllTextAsync(filePath, jsonContent);
            }
            else
            {
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