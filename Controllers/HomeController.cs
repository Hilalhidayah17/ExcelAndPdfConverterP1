using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Rotativa.AspNetCore;
using ClosedXML.Excel;
using PDFGenerator.Models;

namespace PDFGenerator.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        private ListSample GetMockData()
        {
            return new ListSample
            {
                SampleList = new List<Sample>
                {
                    new Sample { Name = "Hilal", Email = "hilal@gmail.com" },
                    new Sample { Name = "Budi", Email = "budi@gmail.com" },
                    new Sample { Name = "Siti", Email = "siti@gmail.com" }
                }
            };
        }

        public IActionResult Index()
        {
            var model = GetMockData();
            return View(model);
        }

        public IActionResult PDFGenerator()
        {
            var data = GetMockData();
            return new ViewAsPdf("PdfResult", data.SampleList)
            {
                FileName = "SamplePdf.pdf"
            };
        }

        public IActionResult ExcelGenerator()
        {
            var data = GetMockData();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Data");

            // Header
            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Email";
            worksheet.Row(1).Style.Font.Bold = true;

            // Data
            for (int i = 0; i < data.SampleList.Count; i++)
            {
                var row = i + 2;
                worksheet.Cell(row, 1).Value = data.SampleList[i].Name;
                worksheet.Cell(row, 2).Value = data.SampleList[i].Email;
            }

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            // Save to byte array to avoid closing stream prematurely
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            var content = ms.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "SampleData.xlsx"
            );
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel
            {
                RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
            });
        }
    }
}
