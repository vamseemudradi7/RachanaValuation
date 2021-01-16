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
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.DelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Rachana\EVEBITDA\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();
            Dictionary<string, List<NSE.StockData>> last100DaysStockData = new Dictionary<string, List<NSE.StockData>>();
            Dictionary<string, List<McapData>> last100DaysMcapData = new Dictionary<string, List<McapData>>();
            List<McapData> MCapData = new List<McapData>();

            foreach (int item in Enumerable.Range(0, 100).ToList())
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
                        var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList(); // Here, we get stocks which arent delisted or on the delistable notice list.
                        last100DaysStockData.Add(stringKey, delistFilteredEquityStockData);
                    }
                    catch (Exception ex) { continue; }
                }
            }

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
                    var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList(); // Here, we get stocks which arent delisted or on the delistable notice list.

                    foreach (var delistFilteredStock in delistFilteredEquityStockData)
                    {
                        string urlAddress = "https://www.google.com/search?q=https://www.moneycontrol.com/stocks/company_info/stock_news.php+" + delistFilteredStock.SYMBOL + "&oq=https://www.moneycontrol.com/stocks/company_info/stock_news.php+" + delistFilteredStock.SYMBOL;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
                        request.Headers.Add("User-Agent", @"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\AppleWebKit / 537.36(KHTML, like Gecko) Safari / 537.36");
                        if (counter % 2 == 0)
                            Thread.Sleep(15000);
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
                                //moneyControlApiUrlForMCap += moneyControlSymbolCode + "%2C";
                                response.Close();
                                readStream.Close();
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
                                    singleMCapData.NoOfShares = (long)((decimal)singleMCapData.MarketCap / (decimal)singleMCapData.LastPrice);
                                    
                                    MCapData.Add(singleMCapData); // Data for Xth day completely with all symbols.
                                }
                            }
                        }
                        catch (Exception e2)
                        {

                        }
                    }
                    last100DaysMcapData.Add(stringKey, MCapData);
                    //last100DaysStockData.Add(stringKey, delistFilteredEquityStockData);
                }
                catch (Exception ex) { }
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
            Dictionary<string, decimal> lastClosePriceOfStocks = new Dictionary<string, decimal>();
            Dictionary<string, decimal> vwapGreaterThan20PercentClosePriceStocks = new Dictionary<string, decimal>();
            foreach (var symbol in exponentialFactorPerStock.Keys)
            {
                decimal exponentialVwapForSymbol = exponentialVwapAggregateForDayPerStock[symbol] / denominatorCountAggregateForVwapCalculationPerStock[symbol];
                exponentialVwapPerStock.Add(symbol, exponentialVwapForSymbol);
                if (!last100DaysStockData.First().Value.Exists(x => x.SYMBOL == symbol))
                    continue;
                var IslastClosePriceValid = decimal.TryParse(last100DaysStockData.First().Value.Find(x => x.SYMBOL == symbol).CLOSE_PRICE, out decimal lastClosePrice);
                if (IslastClosePriceValid)
                {
                    lastClosePriceOfStocks.Add(symbol, lastClosePrice);
                    if (exponentialVwapPerStock[symbol] >= 1.2m * lastClosePriceOfStocks[symbol])
                        vwapGreaterThan20PercentClosePriceStocks.Add(symbol, exponentialVwapPerStock[symbol]);
                }
            }
            string json = JsonConvert.SerializeObject(vwapGreaterThan20PercentClosePriceStocks, Formatting.Indented);
            Console.WriteLine(json);
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
