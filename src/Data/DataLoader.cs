using System;
using System.Collections.Generic;
using AspnetCoreReportDataExportSamples.Model;

namespace AspnetCoreReportDataExportSamples.Data
{
    public static class DataLoader
    {
        public static Lazy<IEnumerable<Order>> Orders = new Lazy<IEnumerable<Order>>(GetOrders);

        private static IEnumerable<Order> GetOrders()
        {
            return new[]
            {
                new Order{Id = 1, CustomerName = "Bob", Date = new DateTime(2020, 01, 01), Total = 100}, 
                new Order{Id = 2, CustomerName = "Sally", Date = new DateTime(2020, 01, 15), Total = 100}, 
                new Order{Id = 3, CustomerName = "Bob", Date = new DateTime(2020, 02, 01), Total = 100}, 
                new Order{Id = 4, CustomerName = "Bob", Date = new DateTime(2020, 02, 15), Total = 100}, 
                new Order{Id = 5, CustomerName = "Dimitri", Date = new DateTime(2020, 03, 01), Total = 100}, 
                new Order{Id = 6, CustomerName = "Xpto", Date = new DateTime(2020, 03, 15), Total = 100}
            };
        }
    }
}
