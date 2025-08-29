using CsvExporterLibrary.Exceptions;
using CsvExporterLibrary.Interfaces;
using ExportExcel.Interfaces;
using ExportExcel.Models;
using ExportToExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using PresentationExporter.Interfaces;
using PresentationExporter.Models;
using System.Diagnostics;
using System.Text;
using TimeZoneConvertorLibrary.Interfaces;

namespace ExportToExcel.Controllers
{
    public class HomeController : Controller
    {

        private readonly IExcelExporter _excelExporter;
        private readonly IPresentationExportService _presentationExportService;
        private readonly string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data");
        private readonly ILogger<HomeController> _logger;
        private readonly ITimeZoneConversionService _timeZoneConversionService;
        private readonly IAsyncExcelImporter _excelToJsonConverter;
        private readonly ICsvExportService _csvExportService;


        public HomeController(ICsvExportService csvExportService, ITimeZoneConversionService timeZoneConversionService,
            ILogger<HomeController> logger, IExcelExporter excelExporter, IAsyncExcelImporter excelImporter , IPresentationExportService presentationExporter)
        {

            _excelToJsonConverter = excelImporter ?? throw new ArgumentNullException(nameof(excelImporter));
            _excelExporter = excelExporter ?? throw new ArgumentNullException(nameof(excelExporter));

            _presentationExportService = presentationExporter ?? throw new ArgumentNullException(nameof(presentationExporter));
            _timeZoneConversionService = timeZoneConversionService ?? throw new ArgumentNullException(nameof(timeZoneConversionService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _csvExportService = csvExportService ?? throw new ArgumentNullException(nameof(csvExportService));
        }



        public IActionResult Index() => View();

        public IActionResult Privacy() => View();

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet("ExportToExcel")]
        public IActionResult ExportToExcel(string? filename = null)
        {
            string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "jsondata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            try
            {
                string jsonString = System.IO.File.ReadAllText(filePath);

                var memoryStream = _excelExporter.ExportJsonToExcel(jsonString);

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"Export_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx"
                    : $"{filename}.xlsx";

                return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputFilename);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("ExportToExcel2")]
        public IActionResult ExportToExcel2(string? filename = null)
        {
            string filePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "cosmosdata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");

            try
            {
                string jsonString = System.IO.File.ReadAllText(filePath);

                var content = _excelExporter.ExportFlattenedJsonToExcel(jsonString);

                if (content.Length == 0)
                    return BadRequest("Invalid JSON format or no data found");

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"CosmosExport_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx"
                    : $"{filename}.xlsx";

                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputFilename);
            }
            catch (Exception ex)
            {
                return BadRequest($"An error occurred: {ex.Message}");
            }
        }

        [HttpGet("ExportToPptx")]
        public async Task<IActionResult> ExportToPptx(string? filename = null)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "jsondata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data file for presentation not found.");

            try
            {
                string jsonString = await System.IO.File.ReadAllTextAsync(filePath);

                var request = new ExportRequest
                {
                    Content = jsonString,
                    Format = InputFormat.Json
                };

                byte[] pptxBytes = await _presentationExportService.ExportPresentationAsync(request);

                if (pptxBytes?.Length > 0)
                {
                    string outputFilename = string.IsNullOrWhiteSpace(filename)
                        ? $"PresentationExport_{DateTime.Now:yyyy-MM-dd_HHmm}.pptx"
                        : $"{filename}.pptx";

                    return File(
                        pptxBytes,
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        outputFilename
                    );
                }

                return BadRequest("Failed to generate presentation from the provided JSON data.");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error creating the presentation: {ex.Message}");
            }
        }

        [HttpGet("ExportToPptx2")]
        public async Task<IActionResult> ExportToPptx2(string? filename = null)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "cosmosdata.json");

            if (!System.IO.File.Exists(filePath))
                return NotFound("Cosmos JSON data file for presentation not found.");

            try
            {
                string jsonString = await System.IO.File.ReadAllTextAsync(filePath);

                var request = new ExportRequest
                {
                    Content = jsonString,
                    Format = InputFormat.CosmosJson
                };

                byte[] pptxBytes = await _presentationExportService.ExportPresentationAsync(request);

                if (pptxBytes?.Length > 0)
                {
                    string outputFilename = string.IsNullOrWhiteSpace(filename)
                        ? $"CosmosPresentationExport_{DateTime.Now:yyyy-MM-dd_HHmm}.pptx"
                        : $"{filename}.pptx";

                    return File(
                        pptxBytes,
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        outputFilename
                    );
                }

                return BadRequest("Failed to generate presentation from the provided Cosmos DB JSON data.");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error creating the Cosmos presentation: {ex.Message}");
            }
        }


        [HttpGet("ExportImagesToPptx")]
        public async Task<IActionResult> ExportImagesToPptx(string? filename = null)
        {
            if (!Directory.Exists(dataFolder))
                return NotFound("Data folder not found.");

            try
            {
                var imageFilePaths = Directory.GetFiles(dataFolder, "*.*")
                    .Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                                f.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (!imageFilePaths.Any())
                    return NotFound("No images found in data folder.");

                var imageBytesList = imageFilePaths
                    .Select(System.IO.File.ReadAllBytes)
                    .ToList();

                var request = new ExportRequest
                {
                    Images = imageBytesList,
                    Format = InputFormat.Images
                };

                byte[] pptxBytes = await _presentationExportService.ExportPresentationAsync(request);

                if (pptxBytes == null || pptxBytes.Length == 0)
                {
                    return BadRequest("Failed to generate a presentation from the provided images.");
                }

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"AllImagesExport_{DateTime.Now:yyyy-MM-dd_HHmm}.pptx"
                    : $"{filename}.pptx";

                return File(pptxBytes,
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    outputFilename);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error creating the presentation: {ex.Message}");
            }
        }


        //[HttpPost("ConvertExcelTimeZone")]
        //public async Task<IActionResult> ConvertExcelTimeZone(IFormFile excelFile, string customFileName , string columnName, string sourceTimeZone, string targetTimeZone, CancellationToken cancellationToken = default)
        //{
        //    if (excelFile == null || excelFile.Length == 0)
        //    {
        //        return BadRequest("No file uploaded or file is empty.");
        //    }

        //    if (string.IsNullOrWhiteSpace(columnName))
        //    {
        //        return BadRequest("Column name is required.");
        //    }

        //    if (string.IsNullOrWhiteSpace(sourceTimeZone))
        //    {
        //        return BadRequest("Source time zone is required.");
        //    }

        //    if (string.IsNullOrWhiteSpace(targetTimeZone))
        //    {
        //        return BadRequest("Target time zone is required.");
        //    }

        //    // Validating file extension
        //    var allowedExtensions = new[] { ".xlsx", ".xlsm" };
        //    var fileExtension = Path.GetExtension(excelFile.FileName).ToLowerInvariant();

        //    if (!allowedExtensions.Contains(fileExtension))
        //    {
        //        return BadRequest("Invalid file type. Only .xlsx and .xlsm files are allowed.");
        //    }

        //    // Check the file size (50MB limit by default)
        //    const long maxFileSize = 50 * 1024 * 1024; // 50 MB
        //    if (excelFile.Length > maxFileSize)
        //    {
        //        return BadRequest("File size exceeds the maximum limit of 50 MB.");
        //    }

        //    try
        //    {
        //        byte[] excelBytes;
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            await excelFile.CopyToAsync(memoryStream, cancellationToken);
        //            excelBytes = memoryStream.ToArray();
        //        }

        //        // Create the conversion request using the new model
        //        var request = new TimeZoneConversionRequest
        //        {
        //            ExcelData = excelBytes,
        //            ColumnName = columnName,
        //            SourceTimeZone = sourceTimeZone,
        //            TargetTimeZone = targetTimeZone,
        //            MaxFileSizeBytes = maxFileSize
        //        };

        //        // Create progress reporter for real-time updates
        //        var progress = new Progress<ConversionProgress>(p =>
        //        {
        //            // You can log progress or send to SignalR hub for real-time updates
        //            // For now, just log it
        //            Console.WriteLine($"Progress: {p.PercentageComplete:F1}% - {p.CurrentOperation}");
        //        });

        //        // Use the orchestrator to perform the conversion
        //        var result = await _orchestator.ConvertExcelTimeStampsAsync(request, progress);

        //        if (!result.Success)
        //        {
        //            return BadRequest($"Conversion failed: {result.Message}");
        //        }

        //        if (result.ConvertedExcelData == null)
        //        {
        //            return BadRequest("Conversion succeeded, but the output data is empty.");
        //        }

        //        // Generate output filename
        //        var originalFileName = Path.GetFileNameWithoutExtension(excelFile.FileName);
        //        var outputFileName = string.IsNullOrWhiteSpace(customFileName) ? $"{originalFileName}_converted_{sourceTimeZone.Replace("/", "_")}_to_{targetTimeZone.Replace("/", "_")}_{DateTime.Now:yyyy-MM-dd_HHmm}.xlsx" : $"{customFileName}.xlsx";

        //        return File(result.ConvertedExcelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputFileName);
        //    }
        //    catch (ArgumentException ex)
        //    {
        //        return BadRequest($"Invalid Input: {ex.Message}");
        //    }
        //    catch (InvalidOperationException ex)
        //    {
        //        return BadRequest($"Processing error: {ex.Message}");
        //    }
        //    catch (OperationCanceledException)
        //    {
        //        return StatusCode(408, "Request was cancelled or timed out");
        //    }
        //    catch (Exception ex)
        //    {
        //        // Log the actual error for debugging
        //        Console.WriteLine($"Unexpected error: {ex}");
        //        return StatusCode(500, "An unexpected error occurred while processing the file");
        //    }
        //}

        //// Helper actions to get available time zones
        //[HttpGet("GetAvailableTimeZones")]
        //public JsonResult GetAvailableTimeZones()
        //{
        //    try
        //    {
        //        var timeZones = _orchestator.GetAvailableTimeZones()
        //            .Select(tz => new { value = tz, text = tz })
        //            .OrderBy(tz => tz.text)
        //            .ToList();

        //        return Json(timeZones);
        //    }
        //    catch
        //    {
        //        // Return some common timezones as fallback
        //        var fallbackTimeZones = new[]
        //        {
        //            new { value = "UTC", text = "UTC" },
        //            new { value = "America/New_York", text = "America/New_York" },
        //            new { value = "America/Los_Angeles", text = "America/Los_Angeles" },
        //            new { value = "Europe/London", text = "Europe/London" },
        //            new { value = "Europe/Paris", text = "Europe/Paris" },
        //            new { value = "Asia/Kolkata", text = "Asia/Kolkata" },
        //            new { value = "Asia/Tokyo", text = "Asia/Tokyo" },
        //            new { value = "Australia/Sydney", text = "Australia/Sydney" }
        //        };

        //        return Json(fallbackTimeZones);
        //    }
        //}

        //// Helper action to validate timezone
        //[HttpPost("ValidateTimeZone")]
        //public JsonResult ValidateTimeZone(string timeZoneId)
        //{
        //    try
        //    {
        //        var isValid = _orchestator.IsValidTimeZone(timeZoneId);
        //        return Json(new { isValid });
        //    }
        //    catch
        //    {
        //        return Json(new { isValid = false });
        //    }
        //}


        [HttpGet("ConvertExcelToJson")]
        public async Task<IActionResult> ConvertExcelToJson(string? filename = null)
        {
            ConversionMode mode = ConversionMode.Auto;
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "Export_2025-08-19_1139.xlsx");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("The specified Excel file was not found.");
            }

            try
            {
                var excelBytes = await System.IO.File.ReadAllBytesAsync(filePath);

                var data = await _excelToJsonConverter.ConvertToDataAsync(excelBytes, mode);

                string jsonString = JsonConvert.SerializeObject(data, Formatting.Indented);

                var jsonBytes = Encoding.UTF8.GetBytes(jsonString);

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"ConvertedFromExcel_{DateTime.Now:yyyy-MM-dd_HHmm}.json"
                    : $"{filename}.json";

                return File(jsonBytes, "application/json", outputFilename);
            }
            catch (Exception ex)
            {
                return BadRequest($"An error occurred during conversion: {ex.Message}");
            }
        }


        [HttpGet("JsonToCsvExporter")]
        public async Task<IActionResult> JsonToCsvExporter(string? filename = null)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "jsondata.json");
            if (!System.IO.File.Exists(filePath))
                return NotFound("JSON data not found");
            try
            {
                string jsonString = await System.IO.File.ReadAllTextAsync(filePath);

                byte[] csvBytes = await _csvExportService.ConvertJsonToCsvAsync(jsonString);

                if (csvBytes.Length == 0)
                    return BadRequest("Conversion resulted in empty CSV data.");

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"Export_{DateTime.Now:yyyy-MM-dd_HHmm}.csv"
                    : $"{filename}.csv";

                return File(csvBytes, "text/csv", outputFilename);
            }
            catch (CsvExportException ex)
            {
                return BadRequest($"Conversion Failed : {ex.UserFriendlyMessage}");
            }
            catch (Exception ex)
            {
                return BadRequest($"An error occurred: {ex.Message}");
            }
        }


        [HttpGet("CsvToJsonExporter")]
        public async Task<IActionResult> CsvToJsonExporter(string? filename = null)
        {
            string filepath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "Export_2025-08-20_1624.csv");

            try
            {
                var csvBytes = await System.IO.File.ReadAllBytesAsync(filepath);

                string jsonString = await _csvExportService.ConvertToJsonAsync(csvBytes);

                if (string.IsNullOrWhiteSpace(jsonString) || jsonString == "[]")
                    return BadRequest("Invalid CSV format or no data found to convert.");

                string outputFilename = string.IsNullOrWhiteSpace(filename)
                    ? $"Export_{DateTime.Now:yyyy-MM-dd_HHmm}.json"
                    : $"{filename}.json";

                return File(Encoding.UTF8.GetBytes(jsonString), "application/json", outputFilename);
            }
            catch (CsvExportException ex)
            {
                return BadRequest($"Conversion Failed : {ex.UserFriendlyMessage}");
            }
            catch (Exception ex)
            {
                return BadRequest($"An error occured : {ex.Message}");
            }
        }


        



    }
}