using System.Text;
using AspnetCoreReportDataExportSamples.Data;
using Microsoft.AspNetCore.Mvc;

namespace AspnetCoreReportDataExportSamples.Controllers
{
    public partial class OrdersController
    {
        /// <summary>
        /// This is the simplest way of exporting data.
        /// This method creates a simple comma-separated values file (CSV)
        /// There are NuGet packages for this, but for here I will keep as simple as possible
        /// </summary>
        /// <returns></returns>
        [HttpGet("/orders/export/csv")]
        public IActionResult ExportCsv()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id,Customer,Date,Total");

            foreach (var order in DataLoader.Orders.Value)
            {
                builder.AppendLine($"{order.Id},{order.CustomerName},{order.Date:yyyy-MM-dd},{order.Total}");
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "orders.csv");
        }
    }
}
