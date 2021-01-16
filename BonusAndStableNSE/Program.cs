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
            Dictionary<string, List<BonusData>> last20YearsStockData = new Dictionary<string, List<BonusData>>();
            Dictionary<string, List<BonusDataOutput>> symbolWiseBonuses = new Dictionary<string, List<BonusDataOutput>>();
            Dictionary<string, int> hasDailyStdDevBeenHigherThan30percentAvgStdDevDays = new Dictionary<string, int>();
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/Companies_proposed_to_be_delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx");
            myWebClient.DownloadFile("https://archives.nseindia.com/content/equities/delisted.xlsx", @"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx");
            List<string> delistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<DelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\DelistedStockSymbols.xlsx", "delisted").Select(x => x.Symbol).ToList();
            List<string> toBeDelistedStockSymbols = new EpPlusHelper().ReadFromExcel<List<ToBeDelistedStockData>>(@"C:\Trading\BhavCopy\Last50DaysNSE\ToBeDelistedStockSymbols.xlsx", "Sheet1").Select(x => x.Symbol).ToList();

            var stocksWhichHaveBeenAnnounced = new List<BonusData>();
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
                List<BonusData> bonusAndDividendData = new EpPlusHelper().ReadFromExcel<List<BonusData>>(extractPath + @"\Bc" + stringKey + ".xlsx", "Bc" + stringKey);
                var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("BONUS") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                foreach (var localBonusData in delistFilteredEquityBonusData)
                {
                    int[] tempResultArray = new int[2];
                    int i = 0;
                    foreach (var split in localBonusData.PURPOSE.Split(":"))
                    {
                        tempResultArray[i] = Int32.Parse(Regex.Match(split, @"\d+").Value);
                        i++;
                    }
                    localBonusData.PURPOSE = ((decimal)tempResultArray[0] / (decimal)tempResultArray[1]).ToString(); // Defines the percentage of shares on existing number of shares.
                }
                var bonusUpcomingStockData = delistFilteredEquityBonusData.Where(x=>DateTime.ParseExact(x.RECORD_DT, "yyyy-dd-MM", CultureInfo.InvariantCulture) > DateTime.Today.AddDays(2)).Select(x => new BonusData { SYMBOL = x.SYMBOL, RECORD_DT = x.RECORD_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
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
                    List<BonusData> bonusAndDividendData = new EpPlusHelper().ReadFromExcel<List<BonusData>>(extractPath + @"\Bc" + stringKey + ".xlsx", "Bc" + stringKey);
                    var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("BONUS") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                    foreach (var localBonusData in delistFilteredEquityBonusData)
                    {
                        int[] tempResultArray = new int[2];
                        int i = 0;
                        foreach (var split in localBonusData.PURPOSE.Split(":"))
                        {
                            tempResultArray[i] = Int32.Parse(Regex.Match(split, @"\d+").Value);
                            i++;
                        }
                        localBonusData.PURPOSE = ((decimal)tempResultArray[0] / (decimal)tempResultArray[1]).ToString(); // Defines the percentage of shares on existing number of shares.
                    }
                    var bonusStockData = delistFilteredEquityBonusData.Select(x => new BonusData { SYMBOL = x.SYMBOL, RECORD_DT = x.RECORD_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
                    last20YearsStockData.Add(stringKey, bonusStockData.ToList());
                }
                catch (Exception ex) { continue; }
            }

            Dictionary<string, int> daysForWhichBonusIsStable = new Dictionary<string, int>();
            var extractPathForStableDate = @"C:\Trading\BhavCopy\NSEBonusStableDate";
            foreach (var currentYearBonusData in last20YearsStockData)
            {
                foreach (var bonusData in currentYearBonusData.Value.Where(x => x.RECORD_DT.Trim() != ""))
                {
                    bool IsYearDateFormat = DateTime.TryParseExact(bonusData.RECORD_DT, "yyyy-MM-dd", CultureInfo.InvariantCulture,DateTimeStyles.None, out DateTime result);
                    if(!IsYearDateFormat)
                        DateTime.TryParseExact(bonusData.RECORD_DT, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    var recordDateMinus6Day = result.AddDays(-6);
                    while (recordDateMinus6Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus6Day.DayOfWeek == DayOfWeek.Sunday)
                        recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                    tempDt = recordDateMinus6Day.ToString("MM-dd-yy").Split("-");
                    stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                    var downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                    try
                    {
                        myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                    }
                    catch (Exception ex) {
                        recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                        while (recordDateMinus6Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus6Day.DayOfWeek == DayOfWeek.Sunday)
                            recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                        tempDt = recordDateMinus6Day.ToString("MM-dd-yy").Split("-");
                        stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                        downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                        try
                        {
                            myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath);
                        }
                        catch
                        {
                            recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                            while (recordDateMinus6Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus6Day.DayOfWeek == DayOfWeek.Sunday)
                                recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                            tempDt = recordDateMinus6Day.ToString("MM-dd-yy").Split("-");
                            stringKey = tempDt[1] + tempDt[0] + tempDt[2];
                            downloadFilePath = @"C:\Trading\BhavCopy\Last10Year\" + stringKey + ".zip";
                            try { myWebClient.DownloadFile("https://www1.nseindia.com/archives/equities/bhavcopy/pr/PR" + stringKey + ".zip", downloadFilePath); }
                            catch
                            {
                                recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                                while (recordDateMinus6Day.DayOfWeek == DayOfWeek.Saturday || recordDateMinus6Day.DayOfWeek == DayOfWeek.Sunday)
                                    recordDateMinus6Day = recordDateMinus6Day.AddDays(-1);
                                tempDt = recordDateMinus6Day.ToString("MM-dd-yy").Split("-");
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

                    bool IsYearDateFormatForRecordDate = DateTime.TryParseExact(bonusData.RECORD_DT, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    if (!IsYearDateFormatForRecordDate)
                        DateTime.TryParseExact(bonusData.RECORD_DT, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
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

                    // Only if after giving Bonus has led to atleast 40 % profit overall:
                    if (decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE) * (1m + Decimal.Parse(bonusData.PURPOSE)) >= 1.40m * (decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE)))
                    {
                        if (!symbolWiseBonuses.ContainsKey(bonusData.SYMBOL))
                            symbolWiseBonuses.Add(bonusData.SYMBOL, new List<BonusDataOutput> { new BonusDataOutput { RecordDate = bonusData.RECORD_DT, BonusRatio = bonusData.PURPOSE.Trim(), ClosePriceOnRPlus1Date = decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE), ClosePriceOnRMinus6Date = decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE) } });
                       else if (!symbolWiseBonuses[bonusData.SYMBOL].Exists(x => x.RecordDate.Trim() == bonusData.RECORD_DT.Trim()))
                            symbolWiseBonuses[bonusData.SYMBOL].Add(new BonusDataOutput { RecordDate = bonusData.RECORD_DT, BonusRatio = bonusData.PURPOSE.Trim(), ClosePriceOnRPlus1Date = decimal.Parse(delistFilteredEquityStockDataRDate.CLOSE_PRICE), ClosePriceOnRMinus6Date = decimal.Parse(delistFilteredEquityStockDataRMinus6.CLOSE_PRICE) });
                    }
                }
            }
            var historicalBonusShareData = symbolWiseBonuses.OrderByDescending(x => x.Value.Count);
            string json = JsonConvert.SerializeObject(historicalBonusShareData, Formatting.Indented);
            Console.WriteLine("Historical Share Bonus Data: -----\n" + json + "\n-------------\n");
            var upcomingStocksWithGenuineBonus = stocksWhichHaveBeenAnnounced.Where(x => historicalBonusShareData.Select(y => y.Key).Contains(x.SYMBOL));
            string upcomingGenuineStockBonuses = JsonConvert.SerializeObject(upcomingStocksWithGenuineBonus, Formatting.Indented);
            Console.WriteLine("Upcoming Genuine Stock Bonus Data: -----\n" + upcomingGenuineStockBonuses + "\n-------------\n");
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
            worksheet.Column(4).Style.Numberformat.Format = "yyyy-MM-dd";
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
