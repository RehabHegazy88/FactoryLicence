using System.ComponentModel.DataAnnotations;
 

namespace PdfExtractorRazor.Models
{
     
        public class CalibrationCertificate
        {
            [Display(Name = "Certificate No")]
            public string CertificateNo { get; set; } = string.Empty;

            [Display(Name = "Equipment Type")]
            public string EquipmentType { get; set; } = string.Empty;

            [Display(Name = "Serial No")]
            public string SerialNo { get; set; } = string.Empty;

            public string Manufacturer { get; set; } = string.Empty;

            [Display(Name = "Model No")]
            public string ModelNo { get; set; } = string.Empty;

            public string Range { get; set; } = string.Empty;

            public string Units { get; set; } = string.Empty;

            [Display(Name = "Accuracy Grade")]
            public string AccuracyGrade { get; set; } = string.Empty;

            [Display(Name = "Calibration Date")]
            [DataType(DataType.Date)]
            public string CalibrationDate { get; set; } = string.Empty;

            [Display(Name = "Next Cal Date")]
            [DataType(DataType.Date)]
            public string NextCalDate { get; set; } = string.Empty;

            public string Location { get; set; } = string.Empty;

            public string Status { get; set; } = string.Empty;

            [Display(Name = "Max Deviation")]
            public string MaxDeviation { get; set; } = string.Empty;

            [Display(Name = "Acceptance Criteria")]
            public string AcceptanceCriteria { get; set; } = string.Empty;
        }

    public class ExtractionResult
    {
        public List<CalibrationCertificate> Certificates { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public int ProcessedFiles { get; set; }
        public bool HasResults => Certificates.Any();
        public string JsonOutput { get; set; } = string.Empty;
        public string? SavedFilePath { get; set; }
        public string? SavedFileName => !string.IsNullOrEmpty(SavedFilePath) ? Path.GetFileName(SavedFilePath) : null;
        public string? DownloadUrl => !string.IsNullOrEmpty(SavedFilePath) ? $"/exports/{Path.GetFileName(SavedFilePath)}" : null;
    }

}
