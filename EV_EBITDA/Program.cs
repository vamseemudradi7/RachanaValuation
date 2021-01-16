using Newtonsoft.Json;
using OfficeOpenXml;
using EV_EBITDA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;

namespace BhavCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            WebClient myWebClient = new WebClient();
            var i = 0;

            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/Companies_proposed_to_be_delisted.xlsx", @"C:\Trading\BhavCopy\Rachana\EVEBITDA\ToBeDelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/delisted.xlsx", @"C:\Trading\BhavCopy\Rachana\EVEBITDA\DelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://www1.nseindia.com/content/equities/suspension.xls", @"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls");
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.DelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols1 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "Suspension prior SOP").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols2 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "SOP Suspended").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols3 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "Liquidation").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols4 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "Surveillence measures").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols5 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "ALF suspended").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbols6 = EpPlusHelper.ReadExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\SuspendedStockSymbols.xls", "OTHERS").Select(x => x.Symbol).ToList();
            List<string> suspendedStockSymbolsAll = suspendedStockSymbols1.Union(suspendedStockSymbols2).Union(suspendedStockSymbols3).Union(suspendedStockSymbols4).Union(suspendedStockSymbols5).Union(suspendedStockSymbols6).ToList();
            
            Dictionary<string, List<NSE.StockData>> last100DaysStockData = new Dictionary<string, List<NSE.StockData>>();
            Dictionary<string, List<McapData>> last100DaysMcapData = new Dictionary<string, List<McapData>>();
            List<McapData> MCapData = new List<McapData>();

            foreach (int item in Enumerable.Range(1, 100).ToList())
            {
                var dateConsidered = DateTime.Now.Date.AddDays(-item);
                if (dateConsidered.DayOfWeek != DayOfWeek.Saturday && dateConsidered.DayOfWeek != DayOfWeek.Sunday)
                {
                    var dateString = dateConsidered.ToString("dd-MM-yy").Split("-");
                    var stringKey = dateString[0] + dateString[1] + dateString[2]; // Date agianst which we are interetsed to get the records.
                    var url = "https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip"; // NSE Data for a given day.

                    try
                    {
                        var downloadFilePath = @"C:\Trading\BhavCopy\Last50DaysNSE\" + stringKey + ".zip";
                        var extractPath = @"C:\Trading\BhavCopy\NSEResponse";
                        myWebClient.DownloadFile(url, downloadFilePath);
                        try { Directory.Delete(extractPath, true); } catch { } // Clearing out NSEResponse folder after every day's stockData is added up
                        System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                        CreateXlsxFile(extractPath, stringKey); // Moving CSV File contents to XLSX format as EPPlus can only read xlsx formatted data.
                        List<NSE.StockData> stockData = new EpPlusHelper().ReadFromExcel<List<NSE.StockData>>(extractPath + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                        var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && x.OPEN_PRICE != null && x.OPEN_PRICE != "" && Decimal.TryParse(x.OPEN_PRICE, out decimal openPriceTemp) && openPriceTemp > 50m && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL) && !suspendedStockSymbolsAll.Contains(x.SYMBOL)).ToList(); // Here, we get stocks which arent delisted or on the delistable notice list.
                        last100DaysStockData.Add(stringKey, delistFilteredEquityStockData);
                    }
                    catch (Exception ex) { continue; }
                }
            }

            Dictionary<string, int> exponentialFactorPerStock = new Dictionary<string, int>();
            Dictionary<string, int> denominatorCountAggregateForVwapCalculationPerStock = new Dictionary<string, int>();
            Dictionary<string, decimal> exponentialVwapAggregateForDayPerStock = new Dictionary<string, decimal>();

            foreach (var item in last100DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null && x.NET_TRDQTY != null && x.NET_TRDQTY != "0" && x.NET_TRDVAL != null && x.NET_TRDVAL != "0"))
                {
                    var isValidTradedQuantity = decimal.TryParse(stock.NET_TRDQTY, out decimal tradedQuantityForDay);
                    var isValidTradedValue = decimal.TryParse(stock.NET_TRDVAL, out decimal tradedValueForDay);
                    if (isValidTradedQuantity && isValidTradedValue) // If Open & PreClose prices are not null in excel sheet
                    {
                        if (!exponentialFactorPerStock.ContainsKey(stock.SYMBOL))
                            exponentialFactorPerStock.Add(stock.SYMBOL, 100);
                        else
                            exponentialFactorPerStock[stock.SYMBOL] -= 1;
                        decimal vwapForDay = tradedValueForDay / tradedQuantityForDay;

                        if (!exponentialVwapAggregateForDayPerStock.ContainsKey(stock.SYMBOL))
                            exponentialVwapAggregateForDayPerStock.Add(stock.SYMBOL, vwapForDay * exponentialFactorPerStock[stock.SYMBOL]);
                        else
                            exponentialVwapAggregateForDayPerStock[stock.SYMBOL] += vwapForDay * exponentialFactorPerStock[stock.SYMBOL];

                        if (!denominatorCountAggregateForVwapCalculationPerStock.ContainsKey(stock.SYMBOL))
                            denominatorCountAggregateForVwapCalculationPerStock.Add(stock.SYMBOL, exponentialFactorPerStock[stock.SYMBOL]);
                        else
                            denominatorCountAggregateForVwapCalculationPerStock[stock.SYMBOL] += exponentialFactorPerStock[stock.SYMBOL];
                    }
                }
                i++;
            }

            Dictionary<string, decimal> exponentialVwapPerStock = new Dictionary<string, decimal>();
            foreach (var symbol in exponentialFactorPerStock.Keys)
                exponentialVwapPerStock.Add(symbol, exponentialVwapAggregateForDayPerStock[symbol] / denominatorCountAggregateForVwapCalculationPerStock[symbol]);

            var consideredDate = DateTime.Now.Date.AddDays(-1);
            if (consideredDate.DayOfWeek != DayOfWeek.Saturday && consideredDate.DayOfWeek != DayOfWeek.Sunday)
            {
                var dateString = consideredDate.ToString("dd-MM-yy").Split("-");
                var stringKey = dateString[0] + dateString[1] + dateString[2]; // Date against which we are interested to get records.
                var url = "https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip"; // NSE Data for a given day.
                int counter = 0;
                try
                {
                    var downloadFilePath = @"C:\Trading\BhavCopy\Last50DaysNSE\" + stringKey + ".zip";
                    var extractPath = @"C:\Trading\BhavCopy\NSEResponse";
                    myWebClient.DownloadFile(url, downloadFilePath);
                    try { Directory.Delete(extractPath, true); } catch { } // Clearing out NSEResponse folder after every day's stockData is added up
                    System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                    CreateXlsxFile(extractPath, stringKey); // Moving CSV File contents to XLSX format as EPPlus can only read xlsx formatted data.
                    List<NSE.StockData> stockData = new EpPlusHelper().ReadFromExcel<List<NSE.StockData>>(extractPath + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                    var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && x.OPEN_PRICE != null && x.OPEN_PRICE != "" && Decimal.TryParse(x.OPEN_PRICE, out decimal openPriceTemp) && openPriceTemp > 50m && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL) && !suspendedStockSymbolsAll.Contains(x.SYMBOL)).ToList(); 
                    
                    foreach (var delistFilteredStock in delistFilteredEquityStockData.OrderBy(x=>x.CLOSE_PRICE))
                    { 
                        string urlAddress = "https://www.google.com/search?q=https://www.moneycontrol.com/stocks/company_info/stock_news.php+" + delistFilteredStock.SYMBOL + "&oq=https://www.moneycontrol.com/stocks/company_info/stock_news.php+" + delistFilteredStock.SYMBOL;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
                        request.Headers.Add("User-Agent", @"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\AppleWebKit / 537.36(KHTML, like Gecko) Safari / 537.36");
                        if (counter % 2 == 0)
                            Thread.Sleep(120000);
                        counter++;
                        try
                        {
                            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                            if (response.StatusCode == HttpStatusCode.OK)
                            {
                                Stream receiveStream = response.GetResponseStream();
                                StreamReader readStream = string.IsNullOrWhiteSpace(response.CharacterSet) ? new StreamReader(receiveStream) : new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                                string data = readStream.ReadToEnd();
                                var moneyControlSymbolCode = data.Split("https://www.moneycontrol.com/stocks/company_info/stock_news.php%3Fsc_id%3D")[1].Split("%")[0];
                                response.Close();
                                readStream.Close();

                                string urlAddressForBalanceSheet = "https://www.moneycontrol.com/financials/britanniaindustries/consolidated-balance-sheetVI/" + moneyControlSymbolCode + "#" + moneyControlSymbolCode;
                                HttpWebRequest requestForBalanceSheet = (HttpWebRequest)WebRequest.Create(urlAddressForBalanceSheet);
                                requestForBalanceSheet.Headers.Add("User-Agent", @"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\AppleWebKit / 537.36(KHTML, like Gecko) Safari / 537.36");
                                HttpWebResponse responseForBalanceSheet = (HttpWebResponse)requestForBalanceSheet.GetResponse();
                                Stream receiveStreamForBalanceSheet = responseForBalanceSheet.GetResponseStream();
                                StreamReader readStreamForBalanceSheet = string.IsNullOrWhiteSpace(responseForBalanceSheet.CharacterSet) ? new StreamReader(receiveStreamForBalanceSheet) : new StreamReader(receiveStreamForBalanceSheet, Encoding.GetEncoding(responseForBalanceSheet.CharacterSet));
                                string balanceSheetData = readStreamForBalanceSheet.ReadToEnd();

                                if (balanceSheetData.Split("Long Term Borrowings").Length <= 1)
                                    continue;

                                decimal longTermBorrowings = decimal.Parse(balanceSheetData.Split("Long Term Borrowings")[1].Split("<td>")[1].Split("</td>")[0]);
                                decimal shortTermBorrowings = decimal.Parse(balanceSheetData.Split("Short Term Borrowings")[1].Split("<td>")[1].Split("</td>")[0]);
                                decimal debt = longTermBorrowings + shortTermBorrowings;
                                decimal cashAndCashEquivalents = decimal.Parse(balanceSheetData.Split("Cash And Cash Equivalents")[1].Split("<td>")[1].Split("</td>")[0]);

                                string urlAddressForConsolidatedQuarterlyResults = "https://www.moneycontrol.com/financials/britanniaindustries/results/consolidated-quarterly-results/" + moneyControlSymbolCode + "#" + moneyControlSymbolCode;
                                HttpWebRequest requestForConsolidatedQuarterlyResults = (HttpWebRequest)WebRequest.Create(urlAddressForConsolidatedQuarterlyResults);
                                requestForConsolidatedQuarterlyResults.Headers.Add("User-Agent", @"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\AppleWebKit / 537.36(KHTML, like Gecko) Safari / 537.36");
                                HttpWebResponse responseForConsolidatedQuarterlyResults = (HttpWebResponse)requestForConsolidatedQuarterlyResults.GetResponse();
                                Stream receiveStreamForConsolidatedQuarterlyResults = responseForConsolidatedQuarterlyResults.GetResponseStream();
                                StreamReader readStreamForConsolidatedQuarterlyResults = string.IsNullOrWhiteSpace(responseForConsolidatedQuarterlyResults.CharacterSet) ? new StreamReader(receiveStreamForConsolidatedQuarterlyResults) : new StreamReader(receiveStreamForConsolidatedQuarterlyResults, Encoding.GetEncoding(responseForConsolidatedQuarterlyResults.CharacterSet));
                                string consolidatedQuarterlyResultsData = readStreamForConsolidatedQuarterlyResults.ReadToEnd();

                                if (consolidatedQuarterlyResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax").Length <= 1)
                                    continue;

                                var quarter1PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[1], "[^0-9.]", "");
                                var quarter2PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[2], "[^0-9.]", "");
                                var quarter3PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[3], "[^0-9.]", "");
                                var quarter4PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[4], "[^0-9.]", "");

                                Decimal.TryParse(quarter1PLBeforeTaxDepreciationString, out decimal quarter1PLBeforeTaxDepreciation);
                                Decimal.TryParse(quarter2PLBeforeTaxDepreciationString, out decimal quarter2PLBeforeTaxDepreciation);
                                Decimal.TryParse(quarter3PLBeforeTaxDepreciationString, out decimal quarter3PLBeforeTaxDepreciation);
                                Decimal.TryParse(quarter4PLBeforeTaxDepreciationString, out decimal quarter4PLBeforeTaxDepreciation);

                                decimal PLBeforeTaxDepreciationLast4QuartersAggregate = quarter1PLBeforeTaxDepreciation + quarter2PLBeforeTaxDepreciation + quarter3PLBeforeTaxDepreciation + quarter4PLBeforeTaxDepreciation;

                                var quarter1DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("Depreciation")[1].Split("<td>")[1], "[^0-9.]", "");
                                var quarter2DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("Depreciation")[1].Split("<td>")[2], "[^0-9.]", "");
                                var quarter3DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("Depreciation")[1].Split("<td>")[3], "[^0-9.]", "");
                                var quarter4DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedQuarterlyResultsData.Split("Depreciation")[1].Split("<td>")[4], "[^0-9.]", "");

                                Decimal.TryParse(quarter1DepreciationString, out decimal quarter1Depreciation);
                                Decimal.TryParse(quarter2DepreciationString, out decimal quarter2Depreciation);
                                Decimal.TryParse(quarter3DepreciationString, out decimal quarter3Depreciation);
                                Decimal.TryParse(quarter4DepreciationString, out decimal quarter4Depreciation);

                                decimal DepreciationLast4QuartersAggregate = quarter1Depreciation + quarter2Depreciation + quarter3Depreciation + quarter4Depreciation;

                                //In Rachana Ranade's document, she compared 2017 to 2020's EBITDA from Annual Rpeorts. 
                                //In my case, I am comparing only last quarter to this quarter(Q1 and Q2 of above variables) EBITDA to speculate the next quarters EBIDTA and compare based on higher growth spurts.

                                var EBITDAQuarter1 = (quarter1PLBeforeTaxDepreciation + quarter1Depreciation);
                                var EBITDAQuarter3 = (quarter3PLBeforeTaxDepreciation + quarter3Depreciation);
                                var latestBiQuarterEBITDACAGR = Math.Pow((double)EBITDAQuarter1 / (double)EBITDAQuarter3, 2) - 1;

                                decimal consolidatedEBITDA = PLBeforeTaxDepreciationLast4QuartersAggregate + DepreciationLast4QuartersAggregate;

                                var moneyControlApiUrlForMCap = "https://api.moneycontrol.com/mcapi/v1/stock/get-stock-price?scIdList=" + moneyControlSymbolCode + "&scId=" + moneyControlSymbolCode;
                                using var client = new HttpClient();
                                string content = "";
                                client.BaseAddress = new Uri("https://api.moneycontrol.com");
                                moneyControlApiUrlForMCap = moneyControlApiUrlForMCap.Split("https://api.moneycontrol.com")[1];
                                client.DefaultRequestHeaders.Accept.Clear();
                                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                                HttpResponseMessage moneyControlApiResponse = client.GetAsync(moneyControlApiUrlForMCap).GetAwaiter().GetResult();
                                if (moneyControlApiResponse.IsSuccessStatusCode)
                                {
                                    content = moneyControlApiResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                                    var deserialisedContent = JsonConvert.DeserializeObject<WrappingMcapData>(content);
                                    var singleMCapData = deserialisedContent.Data.First();
                                    singleMCapData.Symbol = delistFilteredStock.SYMBOL;
                                    singleMCapData.NoOfShares = Convert.ToInt64((singleMCapData.MarketCap * Convert.ToDecimal(Math.Pow(10,7))) / (decimal)singleMCapData.LastPrice);
                                    singleMCapData.ExponentialVwapValue = exponentialVwapPerStock[delistFilteredStock.SYMBOL];
                                    singleMCapData.ComputedMarketCapInCrores = singleMCapData.ExponentialVwapValue * singleMCapData.NoOfShares / (decimal)Math.Pow(10,7);
                                    singleMCapData.ComputedDebt = debt;
                                    singleMCapData.ComputedCCE = cashAndCashEquivalents;
                                    singleMCapData.ComputedEnterpriseValue = singleMCapData.ComputedMarketCapInCrores + debt - cashAndCashEquivalents;
                                    singleMCapData.ComputedEBITDA = consolidatedEBITDA;
                                    singleMCapData.ComputedNetDebt = debt - cashAndCashEquivalents;
                                    singleMCapData.EVEBITDARatio = singleMCapData.ComputedEnterpriseValue / singleMCapData.ComputedEBITDA;
                                    singleMCapData.LatestQuarterEBITDACAGR = Convert.ToDecimal(latestBiQuarterEBITDACAGR);
                                    var estimatedEBITDA = singleMCapData.ComputedEBITDA * (1 + (singleMCapData.LatestQuarterEBITDACAGR / 100m));
                                    var expectedEV = singleMCapData.EVEBITDARatio * estimatedEBITDA;
                                    var expectedEquityValueInRupees = (expectedEV - singleMCapData.ComputedNetDebt) * Convert.ToDecimal(Math.Pow(10,7));
                                    singleMCapData.ExpectedPriceValueNextQuarter = expectedEquityValueInRupees / (decimal)singleMCapData.NoOfShares;
                                    if (singleMCapData.ExpectedPriceValueNextQuarter >= 1.4m * singleMCapData.LastPrice)
                                    {
                                        singleMCapData.ExpectedPriceLastPriceRatio = singleMCapData.ExpectedPriceValueNextQuarter / singleMCapData.LastPrice;
                                        MCapData.Add(singleMCapData);
                                    }
                                }
                            }
                        }
                        catch (Exception ex2)
                        {

                        }
                    }
                }
                catch (Exception ex3) { }
            }
           Console.WriteLine(JsonConvert.SerializeObject(MCapData.OrderByDescending(x=>x.ExpectedPriceLastPriceRatio)));
           //var screenedStocks = MCapData.Where(X => X.ExpectedPriceValueNextQuarter > 1.22m * X.LastPrice);
           //Console.WriteLine(JsonConvert.SerializeObject(screenedStocks));
           Console.ReadLine();
        }

        private static void CreateXlsxFile(string extractPath, string stringKey)
        {
            string csvFileName = extractPath + "\\" + "Pd" + stringKey + ".csv";
            string excelFileName = extractPath + "\\" + "Pd" + stringKey + ".xlsx";
            string worksheetsName = "Pd" + stringKey;
            bool firstRowIsHeader = true;
            var format = new ExcelTextFormat
            {
                Delimiter = ',',
                EOL = "\n"
            };
            using ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName));
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
            worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
            package.Save();
        }
    }
}
