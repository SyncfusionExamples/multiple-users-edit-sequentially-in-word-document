using DocumentEditorApp.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;

namespace DocumentEditorApp.Controllers
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
        public IActionResult Documenteditor(string serviceName, string userName)
        {
            ViewBag.serviceName = serviceName;
            ViewBag.userName = userName;
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}