using Data.Repository;
using Microsoft.AspNetCore.Mvc;

namespace HoiThao_Core.Controllers
{
    public class PrintController : Controller
    {
        private readonly IAseanRepository _aseanRepository;

        public PrintController(IAseanRepository aseanRepository)
        {
            _aseanRepository = aseanRepository;
        }

        public IActionResult Index()
        {
            return View();
        }
    }
}