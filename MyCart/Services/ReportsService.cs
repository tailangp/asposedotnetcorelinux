using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Pdf;

namespace MyCart.Services
{
    public class ReportsService
    {
        public Aspose.Pdf.Document GetPDF()
        {
            // Instantiate Document object
            Document doc = new Document();
            // Add a page to pages collection of PDF file
            Page page = doc.Pages.Add();
            // Instantiate HtmlFragment with HTML contnets
            HtmlFragment titel = new HtmlFragment("<fontsize=10><b><i>Table</i></b></fontsize>");
            // Set bottom margin information
            titel.Margin.Bottom = 10;
            // Set top margin information
            titel.Margin.Top = 200;
            // Add HTML Fragment to paragraphs collection of page
            page.Paragraphs.Add(titel);
            return doc;
        }

        public Aspose.Cells.Workbook GetExcel()
        {
            // Instantiate a Workbook object that represents Excel file.
            Workbook wb = new Workbook();

            // When you create a new workbook, a default "Sheet1" is added to the workbook.
            Worksheet sheet = wb.Worksheets[0];

            // Access the "A1" cell in the sheet.
            Aspose.Cells.Cell cell = sheet.Cells["A1"];

            // Input the "Hello World!" text into the "A1" cell
            cell.PutValue("Hello World!");

            // Save the Excel file.
            return wb;
        }
    }
}
