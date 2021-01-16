using BhavCopy.NSE;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace BhavCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            WebClient myWebClient = new WebClient();
            Dictionary<string, decimal?> currentStockValue = new Dictionary<string, decimal?>();
            Dictionary<string, decimal?> stocksAndTheirHighAvgd = new Dictionary<string, decimal?>();
            Dictionary<string, decimal?> stocksAndTheirLowAvgd = new Dictionary<string, decimal?>();
            Dictionary<string, List<DividendData>> last20YearsStockData = new Dictionary<string, List<DividendData>>();
            Dictionary<string, List<DividendDataOutput>> symbolWiseBonuses = new Dictionary<string, List<DividendDataOutput>>();
            Dictionary<string, int> hasDailyStdDevBeenHigherThan30percentAvgStdDevDays = new Dictionary<string, int>();
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/Companies_proposed_to_be_delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx");
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<DelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();

            var stocksWhichHaveBeenAnnounced = new List<DividendData>();
            var consideredDate = DateTime.Now.Date.AddDays(-1);
            while (consideredDate.DayOfWeek == DayOfWeek.Saturday || consideredDate.DayOfWeek == DayOfWeek.Sunday)
                consideredDate = consideredDate.AddDays(-1);
            var tempDt = consideredDate.ToString("dd-MM-yy").Split("-");
            var stringKey = tempDt[0] + tempDt[1] + tempDt[2];
            try
            {
                var downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                var extractPath = @"C:\Trading\BhavCopy\NSEBonus";
                myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                try { Directory.Delete(extractPath, true); } catch { }
                System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                CreateBcXlsxFile(extractPath, stringKey);
                List<DividendData> bonusAndDividendData = new EpPlusHelper().ReadFromExcel<List<DividendData>>(extractPath + @"\Bc" + stringKey + ".xlsx", "Bc" + stringKey);
                var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("DIV") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                DateTime result;
                foreach (var localBonusData in delistFilteredEquityBonusData)
                {
                    if ((Regex.Match(localBonusData.PURPOSE, @"[\d][.][\d]+").Value) != "")
                        localBonusData.PURPOSE = (decimal.Parse(Regex.Match(localBonusData.PURPOSE, @"[\d][.][\d]+").Value)).ToString(); // Defines the percentage of shares on existing number of shares.
                    else if (Regex.Match(localBonusData.PURPOSE, @"\d+").Value != "")
                        localBonusData.PURPOSE = (int.Parse(Regex.Match(localBonusData.PURPOSE, @"\d+").Value)).ToString();
                    if (!DateTime.TryParseExact(localBonusData.EX_DT, "yyyy-dd-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                        DateTime.TryParseExact(localBonusData.EX_DT, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    localBonusData.EX_DT = result.ToString("dd-MM-yyyy");
                }
                
                var bonusUpcomingStockData = delistFilteredEquityBonusData.Where(x=> Convert.ToDateTime(x.EX_DT) > DateTime.Today.AddDays(2)).Select(x => new DividendData { SYMBOL = x.SYMBOL, EX_DT = x.EX_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
                stocksWhichHaveBeenAnnounced.AddRange(bonusUpcomingStockData);
            }
            catch(Exception ex) { }

            foreach (int month in Enumerable.Range(1, 300).ToList())
            {
                consideredDate = DateTime.Now.Date.AddDays(-1).AddMonths(-month);
                while (consideredDate.DayOfWeek == DayOfWeek.Saturday || consideredDate.DayOfWeek == DayOfWeek.Sunday)
                    consideredDate = consideredDate.AddDays(-1);
                tempDt = consideredDate.ToString("MM-dd-yy").Split("-");
                stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                try
                {
                    var downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                    var extractPath = @"C:\Trading\BhavCopy\NSEBonus";
                    myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                    try { Directory.Delete(extractPath, true); } catch { }
                    System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPath, true);
                    CreateBcXlsxFile(extractPath, stringKey);
                    List<DividendData> bonusAndDividendData = new EpPlusHelper().ReadFromExcel<List<DividendData>>(extractPath + @"\Bc" + stringKey + ".xlsx", "Bc" + stringKey);
                    var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("BONUS") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                    DateTime result;
                    foreach (var localBonusData in delistFilteredEquityBonusData)
                    {
                        if ((Regex.Match(localBonusData.PURPOSE, @"[\d][.][\d]+").Value) != "")
                            localBonusData.PURPOSE = (Decimal.Parse(Regex.Match(localBonusData.PURPOSE, @"[\d][.][\d]+").Value)).ToString(); // Defines the percentage of shares on existing number of shares.
                        else if (Regex.Match(localBonusData.PURPOSE, @"\d+").Value != "")
                            localBonusData.PURPOSE = (Int32.Parse(Regex.Match(localBonusData.PURPOSE, @"\d+").Value)).ToString();
                        if (!DateTime.TryParseExact(localBonusData.EX_DT, "yyyy-dd-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                            DateTime.TryParseExact(localBonusData.EX_DT, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                        localBonusData.EX_DT = result.ToString("dd-MM-yyyy");
                    }
                    var bonusStockData = delistFilteredEquityBonusData.Select(x => new DividendData { SYMBOL = x.SYMBOL, EX_DT = x.EX_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
                    last20YearsStockData.Add(stringKey, bonusStockData.ToList());
                }
                catch (Exception ex) { continue; }
            }

            Dictionary<string, int> daysForWhichBonusIsStable = new Dictionary<string, int>();
            var extractPathForStableDate = @"C:\Trading\BhavCopy\NSEBonusStableDate";
            foreach (var currentYearBonusData in last20YearsStockData)
            {
                foreach (var bonusData in currentYearBonusData.Value.Where(x => x.EX_DT.Trim() != ""))
                {
                    bool IsYearDateFormat = DateTime.TryParseExact(bonusData.EX_DT, "yyyy-MM-dd", CultureInfo.InvariantCulture,DateTimeStyles.None, out DateTime result);
                    if(!IsYearDateFormat)
                        DateTime.TryParseExact(bonusData.EX_DT, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    var recordDateMinus1Day = result.AddDays(-3);
                    while (recordDateMinus1Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus1Day.DayOfWeek == DayOfWeek.Sunday)
                        recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                    tempDt = recordDateMinus1Day.ToString("MM-dd-yy").Split("-");
                    stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                    var downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                    try
                    {
                        myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                    }
                    catch (Exception ex) {
                        recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                        while (recordDateMinus1Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus1Day.DayOfWeek == DayOfWeek.Sunday)
                            recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                        tempDt = recordDateMinus1Day.ToString("MM-dd-yy").Split("-");
                        stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                        downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                        try
                        {
                            myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                        }
                        catch
                        {
                            recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                            while (recordDateMinus1Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus1Day.DayOfWeek == DayOfWeek.Sunday)
                                recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                            tempDt = recordDateMinus1Day.ToString("MM-dd-yy").Split("-");
                            stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                            downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                            try { myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath); }
                            catch
                            {
                                recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                                while (recordDateMinus1Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus1Day.DayOfWeek == DayOfWeek.Sunday)
                                    recordDateMinus1Day = recordDateMinus1Day.AddDays(-1);
                                tempDt = recordDateMinus1Day.ToString("MM-dd-yy").Split("-");
                                stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                                downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                                myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                            }
                        }
                    }

                    try { Directory.Delete(extractPathForStableDate, true); } catch { }
                    System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPathForStableDate, true);
                    CreateXlsxFile(extractPathForStableDate, stringKey);
                    List<StockData> recordDateMinus6DaysStockData = new EpPlusHelper().ReadFromExcel<List<StockData>>(extractPathForStableDate + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                    var delistFilteredEquityStockDataRMinus6 = recordDateMinus6DaysStockData.FirstOrDefault(x => x.SYMBOL == bonusData.SYMBOL && x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)); // Here, we get stocks which arent delisted or on the delistable notice list.
                    if (delistFilteredEquityStockDataRMinus6 == null)
                        continue;

                    bool IsYearDateFormatForRecordDate = DateTime.TryParseExact(bonusData.EX_DT, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    if (!IsYearDateFormatForRecordDate)
                        DateTime.TryParseExact(bonusData.EX_DT, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    var recordDateAfter = result.AddDays(1);
                    while (recordDateAfter.DayOfWeek == DayOfWeek.Saturday || recordDateAfter.DayOfWeek == DayOfWeek.Sunday)
                        recordDateAfter = recordDateAfter.AddDays(1);
                    if (recordDateAfter.Date > DateTime.Today.Date)
                        continue;
                    tempDt = recordDateAfter.ToString("MM-dd-yy").Split("-");
                    stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                    downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                    try
                    {
                        myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                    }
                    catch (Exception ex)
                    {
                        recordDateAfter = recordDateAfter.AddDays(1);
                        while (recordDateAfter.DayOfWeek == DayOfWeek.Saturday || recordDateAfter.DayOfWeek == DayOfWeek.Sunday)
                            recordDateAfter = recordDateAfter.AddDays(1);
                        if (recordDateAfter.Date > DateTime.Today.Date)
                            continue;
                        tempDt = recordDateAfter.ToString("MM-dd-yy").Split("-");
                        stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                        downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                        try
                        {
                            myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                        }
                        catch
                        {
                            recordDateAfter = recordDateAfter.AddDays(1);
                            while (recordDateAfter.DayOfWeek == DayOfWeek.Saturday || recordDateAfter.DayOfWeek == DayOfWeek.Sunday)
                                recordDateAfter = recordDateAfter.AddDays(1);
                            tempDt = recordDateAfter.ToString("MM-dd-yy").Split("-");
                            stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                            downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                            try { myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath); }
                            catch
                            {
                                recordDateAfter = recordDateAfter.AddDays(1);
                                while (recordDateAfter.DayOfWeek == DayOfWeek.Saturday || recordDateAfter.DayOfWeek == DayOfWeek.Sunday)
                                    recordDateAfter = recordDateAfter.AddDays(1);
                                
                                tempDt = recordDateAfter.ToString("MM-dd-yy").Split("-");
                                stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                                downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                                myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                            }
                        }
                    }

                    try { Directory.Delete(extractPathForStableDate, true); } catch { }
                    System.IO.Compression.ZipFile.ExtractToDirectory(downloadFilePath, extractPathForStableDate, true);
                    CreateXlsxFile(extractPathForStableDate, stringKey);
                    List<StockData> recordDateStockData = new EpPlusHelper().ReadFromExcel<List<StockData>>(extractPathForStableDate + @"\Pd" + stringKey + ".xlsx", "Pd" + stringKey);
                    var delistFilteredEquityStockDataRDate = recordDateStockData.FirstOrDefault(x => x.SYMBOL == bonusData.SYMBOL && x.SERIES == "EQ" && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)); // Here, we get stocks which arent delisted or on the delistable notice list.
                    if (delistFilteredEquityStockDataRDate == null)
                        continue;

                    // If Dividend provided is such that it has resulted in 8% profit overall compared to price 6 days ago.
                    if (decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE) + (Decimal.Parse(bonusData.PURPOSE)) >= 1.08m * (decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE)))
                    {
                        if (!symbolWiseBonuses.ContainsKey(bonusData.SYMBOL))
                            symbolWiseBonuses.Add(bonusData.SYMBOL, new List<DividendDataOutput> { new DividendDataOutput { ExDividendDate = bonusData.EX_DT, DividendValue = bonusData.PURPOSE.Trim(), ClosePriceOnExDPlus2Date = decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE), ClosePriceOnExDMinus1Date = decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE) } });
                       else if (!symbolWiseBonuses[bonusData.SYMBOL].Exists(x => x.ExDividendDate.Trim() == bonusData.EX_DT.Trim()))
                            symbolWiseBonuses[bonusData.SYMBOL].Add(new DividendDataOutput { ExDividendDate = bonusData.EX_DT, DividendValue = bonusData.PURPOSE.Trim(), ClosePriceOnExDPlus2Date = decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE), ClosePriceOnExDMinus1Date = decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE) });
                    }
                }
            }
            var historicalBonusShareData = symbolWiseBonuses.OrderByDescending(x => x.Value.Count);
            string json = JsonConvert.SerializeObject(historicalBonusShareData, Formatting.Indented);
            Console.WriteLine("Historical Share Bonus Data: -----\n" + json + "\n-------------\n");
            var upcomingStocksWithGenuineDividend = stocksWhichHaveBeenAnnounced.Where(x => historicalBonusShareData.Select(y => y.Key).Contains(x.SYMBOL));
            string upcomingGenuineStockDividends = JsonConvert.SerializeObject(upcomingStocksWithGenuineDividend, Formatting.Indented);
            Console.WriteLine("Upcoming Genuine Stock Bonus Data: -----\n" + upcomingGenuineStockDividends + "\n-------------\n");
            Console.ReadLine();
        }

        private static void CreateBcXlsxFile(string extractPath, string stringKey)
        {
            string csvFileName = extractPath + "\\" + "Bc" + stringKey + ".csv";
            string excelFileName = extractPath + "\\" + "Bc" + stringKey + ".xlsx";
            string worksheetsName = "Bc" + stringKey;
            bool firstRowIsHeader = true;
            var format = new ExcelTextFormat
            {
                Delimiter = ',',
                EOL = "\r\n"
            };
            format.TextQualifier = '"';
            using ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName));
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
            worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium23, firstRowIsHeader);
            worksheet.Column(7).Style.Numberformat.Format = "yyyy-MM-dd";
            package.Save();
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
