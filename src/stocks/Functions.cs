using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YahooFinanceApi;

namespace stocks
{
    public static class Functions
    {

        [ExcelFunction(Description = "My first .NET function")]
        public static string GetStocksData(string ticker, string field)
        {
            return $"{GetStocksDataAsync(ticker, field).Result}";
        }

        public static async Task<dynamic> GetStocksDataAsync(string ticker, string field)
        {
            var securities = await Yahoo.Symbols(ticker).Fields((Field)Enum.Parse(typeof(Field), field)).QueryAsync();
            var aapl = securities[ticker];
            var price = aapl[(Field)Enum.Parse(typeof(Field), field)]; // or, you could use aapl.RegularMarketPrice directly for typed-value
            
            return price;
        }

    }
}
