using HtmlAgilityPack;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace BhavCopy
{
    class Program // Will revisit ater the bonus one is completed.
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            WebClient myWebClient = new WebClient();
            var i = 0;
            Dictionary<string, decimal?> currentStockValue = new Dictionary<string, decimal?>();
            Dictionary<string, decimal?> stocksAndTheirHighAvgd = new Dictionary<string, decimal?>();
            Dictionary<string, decimal?> stocksAndTheirLowAvgd = new Dictionary<string, decimal?>();
            Dictionary<string, List<NSE.StockData>> last12DaysStockData = new Dictionary<string, List<NSE.StockData>>();
            Dictionary<string, int> hasDailyStdDevBeenHigherThan30percentAvgStdDevDays = new Dictionary<string, int>();

            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/Companies_proposed_to_be_delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx");
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.DelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<NSE.ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();
            foreach (int item in Enumerable.Range(0, 12).ToList())
            {
                var consideredDate = DateTime.Now.Date.AddDays(-item);
                if (consideredDate.DayOfWeek != DayOfWeek.Saturday && consideredDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    var dateString = consideredDate.ToString("dd-MM-yy").Split("-");
                    var stringKey = dateString[0] + dateString[1] + dateString[2]; // Date against which we are interetsed to get the records.
                    var url = "https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip"; // NSE Data for a given day.

                    try
                    {
                        var downloadFilePath = @"C:\Trading\BhavCopy\Last50DaysNSE\" + stringKey + ".zip";
                        var extractPath = @"C:\Trading\BhavCopy\NSEResponse";
                        myWebClient.DownloadFile(url, downloadFilePath);
                        try { Directory.Delete(extractPath, true); } catch { } // Clearing out NSEResponse folder after every day's stockData is added up
                        System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                        CreatexlsxFile(extractPath, stringKey); // Moving CSV File contents to XLSX format as EPPlus can only read xlsx formatted data.
                        List<NSE.StockData> stockData = new EpPlusHelper().ReadFromExcel<List<NSE.StockData>>(extractPath + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                        var delistFilteredEquityStockData = stockData.Where(x => x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList(); // Here, we get stocks which arent delisted or on the delistable notice list.
                        last12DaysStockData.Add(stringKey, delistFilteredEquityStockData);
                        var keyValues = delistFilteredEquityStockData.Select(x => new { Symbol = x.SYMBOL, CurrentPrice = x.CLOSE_PRICE });
                        foreach (var keyValue in keyValues.Where(x => !currentStockValue.ContainsKey(x.Symbol)))
                        {
                            decimal.TryParse(keyValue.CurrentPrice, out decimal currentPrice);
                            currentStockValue.Add(keyValue.Symbol, currentPrice);
                        }
                    }
                    catch (Exception ex) { continue; }
                }
            }
            Dictionary<string, decimal?>[] dailyStdDeviationAvgd = new Dictionary<string, decimal?>[last12DaysStockData.Values.Count];

            foreach (var item in last12DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null && x.OPEN_PRICE != null && x.PREV_CL_PR != null))
                {
                    var validHighPrice = decimal.TryParse(stock.HIGH_PRICE, out decimal highPrice);
                    var validLowPrice = decimal.TryParse(stock.LOW_PRICE, out decimal lowPrice);
                    if (validHighPrice && validLowPrice) // If Open & PreClose prices are not null in excel sheet
                    {
                        var priceDifference = highPrice - lowPrice;
                        if (!stocksAndTheirHighAvgd.ContainsKey(stock.SYMBOL))
                            stocksAndTheirHighAvgd.Add(stock.SYMBOL, highPrice);
                        else
                            stocksAndTheirHighAvgd[stock.SYMBOL] += highPrice;

                        if (!stocksAndTheirLowAvgd.ContainsKey(stock.SYMBOL))
                            stocksAndTheirLowAvgd.Add(stock.SYMBOL, lowPrice);
                        else
                            stocksAndTheirLowAvgd[stock.SYMBOL] += lowPrice;

                        // Storing the Difference of High - Low. Close for each day (denoted by i) and for each stock symbol into : diffOfOpenPrevCloseForEachStock
                        if (dailyStdDeviationAvgd[i] == null)
                            dailyStdDeviationAvgd[i] = new Dictionary<string, decimal?> { { stock.SYMBOL, priceDifference } };
                        else
                            dailyStdDeviationAvgd[i].Add(stock.SYMBOL, priceDifference);
                    }
                }
                i++;
            }

            i = 0;
            Dictionary<string, decimal?> eachStockStdDevAverageOver12Days = new Dictionary<string, decimal?>();
            List<string> stockNames = new List<string>();
            Dictionary<string, int> counterOfIthStock = new Dictionary<string, int>();

            foreach (var item in last12DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null))
                {
                    dailyStdDeviationAvgd[i].TryGetValue(stock.SYMBOL, out decimal? stdDeviation);
                    if (stdDeviation != null)
                    {
                        if (!stockNames.Contains(stock.SYMBOL))
                            stockNames.Add(stock.SYMBOL);

                        if (!counterOfIthStock.ContainsKey(stock.SYMBOL))
                            counterOfIthStock.Add(stock.SYMBOL, 1); // get total count to use later for diving total calculated in line 113
                        else
                            counterOfIthStock[stock.SYMBOL] += 1;

                        if (!eachStockStdDevAverageOver12Days.ContainsKey(stock.SYMBOL))
                            eachStockStdDevAverageOver12Days.Add(stock.SYMBOL, stdDeviation);
                        else
                            eachStockStdDevAverageOver12Days[stock.SYMBOL] += stdDeviation; // eachStockAverageOver50Days , currently add all values and save total
                    }
                }
                i++;
            }

            foreach (var name in stockNames.Distinct().Where(x => eachStockStdDevAverageOver12Days.ContainsKey(x) && counterOfIthStock.ContainsKey(x)))
            {
                try { eachStockStdDevAverageOver12Days[name] = ((decimal)eachStockStdDevAverageOver12Days[name] / (decimal)counterOfIthStock[name]); } catch { } // find average used in line 113 and line 106
                try { stocksAndTheirLowAvgd[name] = ((decimal)stocksAndTheirLowAvgd[name] / (decimal)counterOfIthStock[name]); } catch { }
                try { stocksAndTheirHighAvgd[name] = ((decimal)stocksAndTheirHighAvgd[name] / (decimal)counterOfIthStock[name]); } catch { }
            }

            i = 0;
            foreach (var item in last12DaysStockData)
            {
                foreach (var stock in item.Value.Where(x => x.SYMBOL != null && x.OPEN_PRICE != null && x.PREV_CL_PR != null))
                {
                    var validHighPrice = decimal.TryParse(stock.HIGH_PRICE, out decimal highPrice);
                    var validLowPrice = decimal.TryParse(stock.LOW_PRICE, out decimal lowPrice);
                    if (validHighPrice && validLowPrice) // If Open & PreClose prices are not null in excel sheet
                    {
                        var priceDifference = highPrice - lowPrice;
                        if (priceDifference > 0.98m * dailyStdDeviationAvgd[i][stock.SYMBOL])
                        {
                            if (!hasDailyStdDevBeenHigherThan30percentAvgStdDevDays.ContainsKey(stock.SYMBOL))
                                hasDailyStdDevBeenHigherThan30percentAvgStdDevDays.Add(stock.SYMBOL, 1);
                            else
                                hasDailyStdDevBeenHigherThan30percentAvgStdDevDays[stock.SYMBOL] += 1;
                        }
                        else if (!hasDailyStdDevBeenHigherThan30percentAvgStdDevDays.ContainsKey(stock.SYMBOL))
                                hasDailyStdDevBeenHigherThan30percentAvgStdDevDays.Add(stock.SYMBOL, 0);
                    }
                }
                i++;
            }
            var numbers = new List<string>();
            foreach (var num in Enumerable.Range(0, 9).ToArray())
                numbers.Add("'" + num + "'");
            var screenedAvgStdDevStocks = eachStockStdDevAverageOver12Days.Where(x => currentStockValue.ContainsKey(x.Key) && x.Value >= (currentStockValue[x.Key] * 0.07m));
            var screenedStocks = from avgStdDev in screenedAvgStdDevStocks
                                 join lowPrice in stocksAndTheirLowAvgd on avgStdDev.Key equals lowPrice.Key
                                 join highPrice in stocksAndTheirHighAvgd on avgStdDev.Key equals highPrice.Key
                                 join positive40PercentileStdDevDay in hasDailyStdDevBeenHigherThan30percentAvgStdDevDays on avgStdDev.Key equals positive40PercentileStdDevDay.Key
                                 where currentStockValue[avgStdDev.Key] < (1.05m * lowPrice.Value) && (avgStdDev.Value >= 5) && (((decimal?)positive40PercentileStdDevDay.Value) / ((decimal)last12DaysStockData.Count)) > 0.93m // && currentStockValue[avgStdDev.Key] < ((lowPrice.Value + highPrice.Value) /2)
                                 select new { Stock = highPrice.Key, AverageStdDev = avgStdDev.Value, CurrentStockPrice = currentStockValue[avgStdDev.Key], StdDevToCurrentValueRatio = (avgStdDev.Value / currentStockValue[avgStdDev.Key]) * 100, StdDeviationMorethan90thPercentileOfAvgStdDevDays = hasDailyStdDevBeenHigherThan30percentAvgStdDevDays[avgStdDev.Key], PricePositiveByTotalDays = (((decimal?)hasDailyStdDevBeenHigherThan30percentAvgStdDevDays[avgStdDev.Key]) / ((decimal)last12DaysStockData.Count)) };
            screenedStocks = screenedStocks.OrderByDescending(x => x.StdDevToCurrentValueRatio);                
            string json = JsonConvert.SerializeObject(screenedStocks, Formatting.Indented);
            Console.WriteLine(json);
            Console.ReadLine();
        }

        private static void CreatexlsxFile(string extractPath, string stringKey)
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
