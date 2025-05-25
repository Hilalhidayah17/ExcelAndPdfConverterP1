using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Rotativa.AspNetCore;

namespace PDFGenerator.Controllers
{
    public class ExportController : Controller
    {
        public IActionResult Index() => View();

        public IActionResult DownloadExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Data");

            worksheet.Cells[1, 1].Value = "Nama";
            worksheet.Cells[1, 2].Value = "Umur";

            worksheet.Cells[2, 1].Value = "Andi";
            worksheet.Cells[2, 2].Value = 25;

            worksheet.Cells[3, 1].Value = "Budi";
            worksheet.Cells[3, 2].Value = 30;

            var stream = new MemoryStream(package.GetAsByteArray());
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Data.xlsx");
        }

        public IActionResult DownloadPdf()
        {
            var data = new List<(string Nama, int Umur)>
            {
                ("Andi", 25),
                ("Budi", 30)
            };
            return new ViewAsPdf("PdfView", data)
            {
                FileName = "Data.pdf"
            };
        }
    }
}
