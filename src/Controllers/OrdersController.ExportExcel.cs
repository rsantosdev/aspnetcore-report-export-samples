using System.IO;
using System.Threading.Tasks;
using AspnetCoreReportDataExportSamples.Data;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace AspnetCoreReportDataExportSamples.Controllers
{
    public partial class OrdersController
    {
        /// <summary>
        /// In case we need to add more complexity, styles and/or information for our users
        /// we will need to export data as real excel files.
        /// Since xlsx files are xml based you can generate yourself but that's not recommended,
        /// there are some really good packages out there and we will use ClosedXml here
        /// For more info visit: https://github.com/ClosedXML/ClosedXML
        /// </summary>
        /// <returns></returns>
        [HttpGet("/orders/export/excel")]
        public async Task<IActionResult> ExportExcel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Orders");
            var currentRow = 1;

            worksheet.Cell(currentRow, 1).Value = "Id";
            worksheet.Cell(currentRow, 2).Value = "Customer";
            worksheet.Cell(currentRow, 3).Value = "Date";
            worksheet.Cell(currentRow, 4).Value = "Total";

            foreach (var order in DataLoader.Orders.Value)
            {
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = order.Id;
                worksheet.Cell(currentRow, 2).Value = order.CustomerName;
                worksheet.Cell(currentRow, 3).Value = order.Date;
                worksheet.Cell(currentRow, 4).Value = order.Total;
            }

            currentRow++;
            worksheet.Cell(currentRow, 3).Value = "Total";
            worksheet.Cell(currentRow, 3).Style.Font.Bold = true;

            worksheet.Cell(currentRow, 4).FormulaA1 = $"=SUM(D2:D{currentRow-1})";
            worksheet.Cell(currentRow, 4).Style.Font.Bold = true;

            await using var memory = new MemoryStream();
            workbook.SaveAs(memory);

            return File(memory.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "orders.xlsx");
        }
    }
}
