using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages; 
using static System.Collections.Specialized.BitVector32;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.ComponentModel.DataAnnotations;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Runtime.ConstrainedExecution;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using System;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig;

namespace PdfExtractorRazor.Pages; 
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Hosting;
using PdfExtractorRazor.Models;
using PdfExtractorRazor.Services;

public class IndexModel : PageModel
{
    private readonly IPdfExtractionService _extractionService;
    private readonly ILogger<IndexModel> _logger;

    public IndexModel(IPdfExtractionService extractionService, ILogger<IndexModel> logger)
    {
        _extractionService = extractionService;
        _logger = logger;
    }

    [BindProperty]
    public IFormFileCollection UploadedFiles { get; set; } = default!;

    public ExtractionResult? ExtractionResult { get; set; }

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (UploadedFiles == null || UploadedFiles.Count == 0)
        {
            ModelState.AddModelError("UploadedFiles", "Please select at least one PDF file.");
            return Page();
        }

        try
        {
            ExtractionResult = await _extractionService.ExtractFromFilesAsync(UploadedFiles);

            if (ExtractionResult.HasResults)
            {
                _logger.LogInformation("Successfully processed {Count} certificates and saved to {FilePath}",
                    ExtractionResult.Certificates.Count, ExtractionResult.SavedFilePath);

                TempData["SuccessMessage"] = $"Successfully extracted {ExtractionResult.Certificates.Count} certificates and saved to {ExtractionResult.SavedFileName}";
            }

            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing uploaded files");
            ModelState.AddModelError("", "An error occurred while processing the files.");
            return Page();
        }
    }

    public IActionResult OnPostDownloadJson()
    {
        var jsonData = Request.Form["jsonData"];
        if (string.IsNullOrEmpty(jsonData))
        {
            return BadRequest("No JSON data available");
        }

        var bytes = System.Text.Encoding.UTF8.GetBytes(jsonData);
        var fileName = $"calibration_certificates_{DateTime.Now:yyyyMMdd_HHmmss}.json";

        return File(bytes, "application/json", fileName);
    }
}


