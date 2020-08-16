using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using YahooFinanceApi;

namespace stocks
{
    public static class Functions
    {
        private const string FIELDS = "Ask,AskSize,AverageDailyVolume10Day,AverageDailyVolume3Month,Bid,BidSize,BookValue,Currency,DividendDate,EarningsTimestamp,EarningsTimestampEnd,EarningsTimestampStart,EpsForward,EpsTrailingTwelveMonths,Exchange,ExchangeDataDelayedBy,ExchangeTimezoneName,ExchangeTimezoneShortName,FiftyDayAverage,FiftyDayAverageChange,FiftyDayAverageChangePercent,FiftyTwoWeekHigh,FiftyTwoWeekHighChange,FiftyTwoWeekHighChangePercent,FiftyTwoWeekLow,FiftyTwoWeekLowChange,FiftyTwoWeekLowChangePercent,FinancialCurrency,ForwardPE,FullExchangeName,GmtOffSetMilliseconds,Language,LongName,Market,MarketCap,MarketState,MessageBoardId,PriceHint,PriceToBook,QuoteSourceName,QuoteType,RegularMarketChange,RegularMarketChangePercent,RegularMarketDayHigh,RegularMarketDayLow,RegularMarketOpen,RegularMarketPreviousClose,RegularMarketPrice,RegularMarketTime,RegularMarketVolume,PostMarketChange,PostMarketChangePercent,PostMarketPrice,PostMarketTime,SharesOutstanding,ShortName,SourceInterval,Symbol,Tradeable,TrailingAnnualDividendRate,TrailingAnnualDividendYield,TrailingPE,TwoHundredDayAverage,TwoHundredDayAverageChange,TwoHundredDayAverageChangePercent";

        [ExcelFunction(Description = "My first .NET function")]
        public static string GetStocksData(string ticker, string field)
        {
            return $"{GetStocksDataAsync(ticker, field).Result}";
        }

        public static async Task<dynamic> GetStocksDataAsync(string ticker, string field)
        {
            var securities = await Yahoo.Symbols(ticker).Fields((Field)Enum.Parse(typeof(Field), field)).QueryAsync();
            var aapl = securities[ticker];
            var price = aapl[(Field)Enum.Parse(typeof(Field), field)];
            
            return price;
        }

        public static Microsoft.Office.Interop.Excel.Application GetApp()
        {
            return ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application;
        }

        [ExcelCommand(MenuName = "Stocks", MenuText = "Create Table")]
        public static void CreateStocksTable()
        {
            Workbook wb = GetApp().ActiveWorkbook;
            Worksheet ws = GetApp().ActiveSheet;
            Range cell = GetApp().ActiveCell;

            if (wb == null)
            {
                MessageBox.Show("엑셀 파일을 먼저 실행해주세요.", "error", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ListObject table = ws.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, XlListObjectHasHeaders: XlYesNoGuess.xlNo);
            table.Name = table.Name + "_stock";
            table.Range[1, 1] = "ticker";
            table.Range[1, 2] = "RegularMarketPrice";
            table.Range[1, 2].Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, FIELDS);

            table.ShowAutoFilter = false;

            table.Range[2, 1] = "AAPL";
            table.Range[3, 1] = "GOOG";

        }

        [ExcelCommand(MenuName = "Stocks", MenuText = "Refresh Tables", ShortCut = "^Q")]
        public static void RefreshTables()
        {
            Workbook wb = GetApp().ActiveWorkbook;
            
            if (wb == null)
            {
                MessageBox.Show("엑셀 파일을 먼저 실행해주세요.", "error", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ExcelAsyncTask.Run(RefreashTableTaskAsync);
        }


        public async static Task RefreashTableTaskAsync()
        {
            Workbook wb = GetApp().ActiveWorkbook;

            // 모든 업데이트 테이블 가져오기
            List<ListObject> tables = new List<ListObject>();

            foreach (Worksheet ws in wb.Worksheets)
            {
                foreach (ListObject table in ws.ListObjects)
                {
                    if (table.Name.EndsWith("_stock"))
                    {
                        tables.Add(table);
                    }
                }
            }


            foreach (ListObject table in tables)
            {
                // table의 param 가져오기
                List<string> tickers = new List<string>();
                List<Field> fields = new List<Field>();

                for (int r = 2; r <= 1 + table.ListRows.Count; r++)
                {
                    tickers.Add(table.Range[r, 1].Text);
                }

                for (int c = 2; c <= table.ListColumns.Count; c++)
                {
                    if (Enum.TryParse<Field>(table.Range[1, c].Text, out Field result))
                    {
                        fields.Add(result);
                    }
                    else
                    {
                        MessageBox.Show($"알 수 없는 Field가 있습니다! '{table.Range[1, c].Text}'", "error", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // yahoo finance 정보 가져오기 (비동기)
                var securities = await Yahoo.Symbols(tickers.ToArray()).Fields(fields.ToArray()).QueryAsync();

                int row = 2;
                // 시트에 정보 입력
                foreach (string ticker in tickers)
                {
                    int col = 2;

                    if (securities.ContainsKey(ticker) == false)
                    {
                        continue;
                    }

                    var data = securities[ticker];

                    foreach (Field field in fields)
                    {
                        table.Range[row, col].Value = data[field];

                        ++col;
                    }

                    ++row;
                }

            }
        }

    }
}
