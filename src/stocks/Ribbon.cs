using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using YahooFinanceApi;

namespace stocks
{
    public partial class Ribbon
    {
        private const string FIELDS = "Ask,AskSize,AverageDailyVolume10Day,AverageDailyVolume3Month,Bid,BidSize,BookValue,Currency,DividendDate,EarningsTimestamp,EarningsTimestampEnd,EarningsTimestampStart,EpsForward,EpsTrailingTwelveMonths,Exchange,ExchangeDataDelayedBy,ExchangeTimezoneName,ExchangeTimezoneShortName,FiftyDayAverage,FiftyDayAverageChange,FiftyDayAverageChangePercent,FiftyTwoWeekHigh,FiftyTwoWeekHighChange,FiftyTwoWeekHighChangePercent,FiftyTwoWeekLow,FiftyTwoWeekLowChange,FiftyTwoWeekLowChangePercent,FinancialCurrency,ForwardPE,FullExchangeName,GmtOffSetMilliseconds,Language,LongName,Market,MarketCap,MarketState,MessageBoardId,PriceHint,PriceToBook,QuoteSourceName,QuoteType,RegularMarketChange,RegularMarketChangePercent,RegularMarketDayHigh,RegularMarketDayLow,RegularMarketOpen,RegularMarketPreviousClose,RegularMarketPrice,RegularMarketTime,RegularMarketVolume,PostMarketChange,PostMarketChangePercent,PostMarketPrice,PostMarketTime,SharesOutstanding,ShortName,SourceInterval,Symbol,Tradeable,TrailingAnnualDividendRate,TrailingAnnualDividendYield,TrailingPE,TwoHundredDayAverage,TwoHundredDayAverageChange,TwoHundredDayAverageChangePercent";

        private Timer timer = new Timer();

        public Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            timer.Tick += Timer_Tick;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            RefreashTableTaskAsync();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = App.ActiveWorkbook;
            Worksheet ws = App.ActiveSheet;
            Range cell = App.ActiveCell;

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

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = App.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("엑셀 파일을 먼저 실행해주세요.", "error", buttons: MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            RefreashTableTaskAsync();
        }

        public async void RefreashTableTaskAsync()
        {
            Workbook wb = App.ActiveWorkbook;

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

            // table의 param 가져오기
            List<string> tickers = new List<string>();
            List<Field> fields = new List<Field>();

            foreach (ListObject table in tables)
            {

                for (int r = 2; r <= 1 + table.ListRows.Count; r++)
                {
                    string text = table.Range[r, 1].Text;

                    if (tickers.Contains(text) == false)
                    {
                        tickers.Add(text);
                    }
                }

                for (int c = 2; c <= table.ListColumns.Count; c++)
                {
                    if (Enum.TryParse(table.Range[1, c].Text, out Field field))
                    {
                        if (fields.Contains(field) == false)
                        {
                            fields.Add(field);
                        }
                    }
                }
            }

            // yahoo finance 정보 가져오기 (비동기)
            var securities = await Yahoo.Symbols(tickers.ToArray()).Fields(fields.ToArray()).QueryAsync();
            
            try
            {
                foreach (ListObject table in tables)
                {
                    // 시트에 정보 입력
                    for (int r = 2; r <= 1 + table.ListRows.Count; r++)
                    {
                        string ticker = table.Range[r, 1].Text;

                        var data = securities[ticker];

                        for (int c = 2; c <= table.ListColumns.Count; c++)
                        {
                            if (Enum.TryParse(table.Range[1, c].Text, out Field field))
                            {
                                table.Range[r, c].Value = data[field];
                            }
                        }
                    }
                }
            }
            catch 
            {
                // Do nothing
                // 셀 값을 입력할 수 없는 타이밍이 때 걸린 경우.
            }
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.toggleButton1.Checked)
            {
                // 최소 0.5초
                if (int.TryParse(editBox1.Text, out int result))
                {
                    timer.Interval = Math.Max(result, 500);
                }
                else
                {
                    timer.Interval = 1000;
                }

                editBox1.Text = timer.Interval.ToString();
                timer.Start();
            }
            else
            {
                timer.Stop();
            }
        }
    }
}
