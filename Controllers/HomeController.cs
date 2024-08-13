using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WeeklyStatusReport.Models;

namespace WeeklyStatusReport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Submit(TeamSelection model)
        {
           /* if(model.SelectedTeam=="Cloud Sailors")
                return RedirectToAction("WSRForm","Testing", new { selectedTeam = model.SelectedTeam });*/
            return RedirectToAction("WSRForm", "Testing", new { selectedTeam = model.SelectedTeam });
        }

    }
}
