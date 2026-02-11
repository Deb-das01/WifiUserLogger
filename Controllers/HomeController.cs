using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
using WifiUserLogger.Models;
using WifiUserLogger.Utils;

namespace WifiUserLogger.Controllers
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
        public IActionResult UpdateAttandance(string UserName)
        {
            string ip = HttpContext.Connection.RemoteIpAddress?.ToString();

            // Wi-Fi validation
            if (ip == null || !(ip.StartsWith(ConfigHelper.GetConfigValue("GatewayIP")) || ip == "::1" || ip == "127.0.0.1"))
            {
                TempData["ErrorMessage"] = "Not on office Wi-Fi";
                return RedirectToAction("Index");
            }


            string dateToday = DateTime.Now.ToString("yyyy-MM-dd");
            string sheetName = $"{ConfigHelper.GetConfigValue("SheetNamePrefix")}{dateToday}";
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), ConfigHelper.GetConfigValue("ExcelFileName"));

            // Create file if not exists
            if (!System.IO.File.Exists(filePath))
            {
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(sheetName);
                    workbook.SaveAs(filePath);
                }
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                // Create today's sheet if missing
                var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == sheetName)
                                ?? workbook.Worksheets.Add(sheetName);

                // Add header if empty
                if (worksheet.LastRowUsed() == null)
                {
                    worksheet.Cell(1, 1).Value = "UserName";
                    worksheet.Cell(1, 2).Value = "IP Address";
                    worksheet.Cell(1, 3).Value = "Login Time";
                }

                // Prevent duplicate attendance
                bool alreadyMarked = worksheet.RowsUsed()
                    .Skip(1)
                    .Any(r => r.Cell(1).GetString() == UserName);

                if (alreadyMarked)
                {
                    TempData["ErrorMessage"] = "Attendance already marked for today.";
                    return RedirectToAction("Index");

                }

                int lastRow = worksheet.LastRowUsed().RowNumber() + 1;

                worksheet.Cell(lastRow, 1).Value = UserName;
                worksheet.Cell(lastRow, 2).Value = ip;
                worksheet.Cell(lastRow, 3).Value = DateTime.Now;

                workbook.Save();
            }

            return View("Success");
        }

    }
}
