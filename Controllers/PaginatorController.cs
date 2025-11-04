using Microsoft.AspNetCore.Mvc;

namespace ExportToExcel.Controllers
{
    public class PaginatorController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
