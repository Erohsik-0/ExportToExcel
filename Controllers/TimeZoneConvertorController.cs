using Microsoft.AspNetCore.Mvc;
using TimeZoneConvertorLibrary.Interfaces;
using TimeZoneConvertorLibrary.Models;

namespace ExportToExcel.Controllers
{

    [ApiController]
    [Route("api/[controller]")]
    public class TimeZoneConvertorController : Controller
    {

        private readonly ITimeZoneConversionService _service;
        private readonly ILogger<TimeZoneConvertorController> _logger;

        public TimeZoneConvertorController(ITimeZoneConversionService service, ILogger<TimeZoneConvertorController> logger)
        {
            _service = service;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }


        [HttpPost("convert-async")]
        public async Task<IActionResult> ConvertFromAsyncBody([FromBody] TimeZoneConversionRequest request, CancellationToken cancellationToken)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            try
            {
                // Call your library's async method
                var result = await _service.ConvertDateTimeAsync(request, cancellationToken);

                if (result.Success)
                {
                    // If successful, return 200 OK with the result
                    return Ok(result);
                }

                // If the library handled a known error (e.g., invalid timezone), return 400 Bad Request
                return BadRequest(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An unexpected error occurred in ConvertFromAsyncBody.");
                // For any other unexpected errors, return a 500 Internal Server Error
                return StatusCode(500, "An internal server error occurred.");
            }
        }

    }
}
