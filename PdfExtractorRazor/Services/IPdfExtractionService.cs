
using PdfExtractorRazor.Models;

namespace PdfExtractorRazor.Services
{
    public interface IPdfExtractionService
    {
        Task<ExtractionResult> ExtractFromFilesAsync(IFormFileCollection files);
        Task<CalibrationCertificate?> ExtractFromSingleFileAsync(IFormFile file);
        string ConvertToJson(List<CalibrationCertificate> certificates);
    }
}