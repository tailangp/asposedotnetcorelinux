using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using MyCart.Services;
using Microsoft.AspNetCore.Mvc;

namespace MyCart.Controllers
{
    [Route("api/[controller]")]
    public class ReportsController : Controller
    {
        [HttpGet("[action]")]
        public async Task<IActionResult> DownloadExcel()
        {
            ReportsService rs = new ReportsService();
            var pdf = rs.GetExcel();
            string fileName = Guid.NewGuid() + ".xlsx";
            pdf.Save(fileName, Aspose.Cells.SaveFormat.Xlsx);
            return GetFile(fileName);
        }

        [HttpGet("[action]")]
        public async Task<IActionResult> DownloadPdf()
        {
            ReportsService rs = new ReportsService();
            var pdf = rs.GetPDF();
            string fileName = Guid.NewGuid() + ".pdf";
            pdf.Save(fileName, Aspose.Pdf.SaveFormat.Pdf);
            return GetFile(fileName);
        }

        private IActionResult GetFile(string fileName)
        {
            var dataBytes = System.IO.File.ReadAllBytes(fileName);
            var dataStream = new MemoryStream(dataBytes);

            if (dataStream == null)
                return NotFound();

            return File(dataStream, "application/octet-stream", fileName); // returns a FileStreamResult
        }
    }
}
