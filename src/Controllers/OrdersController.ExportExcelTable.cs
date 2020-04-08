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
        [HttpGet("/orders/export/excel-table")]
        public async Task<IActionResult> ExportExcelTable()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Orders");
            
            // Title
            worksheet.Cell(1, 1).Value = "Orders";
            
            var currentRow = 2;
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

            // Creates Table
            var rngTable = worksheet.Range($"A1:D{currentRow}");
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // Title
            rngTable.Cell(1, 1).Style.Font.Bold = true;
            rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngTable.Range("A1:D1").Merge();

            // Format Headers
            var rngHeaders = rngTable.Range("A2:D2");
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSkyBlue;

            // Creates Excel Table
            var rngData = worksheet.Range($"A2:D{currentRow}");
            var excelTable = rngData.CreateTable();
            excelTable.ShowTotalsRow = true;
            excelTable.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
            excelTable.Field("Date").TotalsRowLabel = "TOTAL:";

            await using var memory = new MemoryStream();
            workbook.SaveAs(memory);

            return File(memory.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "orders_table.xlsx");
        }
    }
}
