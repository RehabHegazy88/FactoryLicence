using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using UglyToad.PdfPig;

namespace PdfExtractorRazor.Pages;

public class UploadPdfModel : PageModel
{
	[BindProperty]
	public IFormFile? Pdf { get; set; }

	public string? ExtractedText { get; private set; }
	public string? ErrorMessage { get; private set; }

	public void OnGet()
	{
	}

	public async Task<IActionResult> OnPostAsync()
	{
		if (Pdf == null || Pdf.Length == 0)
		{
			ErrorMessage = "Please choose a PDF file.";
			return Page();
		}

		if (!string.Equals(Pdf.ContentType, "application/pdf", StringComparison.OrdinalIgnoreCase)
			&& !Pdf.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
		{
			ErrorMessage = "File must be a PDF.";
			return Page();
		}

		// Optional: size guard (50 MB)
		const long maxBytes = 50L * 1024 * 1024;
		if (Pdf.Length > maxBytes)
		{
			ErrorMessage = "File exceeds 50 MB limit.";
			return Page();
		}

		try
		{
			await using var memory = new MemoryStream();
			await Pdf.CopyToAsync(memory);
			memory.Position = 0;

			var builder = new StringBuilder();
			using (var document = PdfDocument.Open(memory))
			{
				foreach (var page in document.GetPages())
				{
					builder.AppendLine(page.Text);
				}
			}

			ExtractedText = builder.ToString();
		}
		catch (Exception ex)
		{
			ErrorMessage = $"Failed to read PDF: {ex.Message}";
		}

		return Page();
	}
}