
using PdfExtractorRazor.Models;

namespace PdfExtractorRazor.Services
{
    public interface IPdfExtractionService
    {
        Task<ExtractionResult> ExtractFromFilesAsync(IFormFileCollection files);
        Task<CalibrationCertificate?> ExtractFromSingleFileAsync(IFormFile file);
        string ConvertToJson(List<CalibrationCertificate> certificates);
        // Task<string> SaveToJsonFileAsync(List<CalibrationCertificate> certificates, string? customFileName = null);
         Task<string> SaveToJsonFileAsync(List<CalibrationCertificate> certificates, string? customFileName = null, bool append = false);


    }
}