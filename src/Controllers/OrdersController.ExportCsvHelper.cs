using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using AspnetCoreReportDataExportSamples.Data;
using CsvHelper;
using Microsoft.AspNetCore.Mvc;

namespace AspnetCoreReportDataExportSamples.Controllers
{
    public partial class OrdersController
    {
        /// <summary>
        /// This method uses the CsvHelper package an easy and simple package for write/read
        /// csv files.
        /// For more info visit: https://joshclose.github.io/CsvHelper/getting-started#writing-a-csv-file
        /// </summary>
        /// <returns></returns>
        [HttpGet("/orders/export/csv-helper")]
        public async Task<IActionResult> ExportCsvHelper()
        {
            await using var memory = new MemoryStream();
            await using var stw = new StreamWriter(memory);
            await using var csv = new CsvWriter(stw, CultureInfo.InvariantCulture);

            await csv.WriteRecordsAsync(DataLoader.Orders.Value);

            await csv.FlushAsync();
            await stw.FlushAsync();
            return File(memory.ToArray(), "text/csv", "orders_csv_helper.csv");
        }
    }
}
