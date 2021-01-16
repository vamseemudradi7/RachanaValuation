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
                var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("RIGHTS") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                var bonusUpcomingStockData = delistFilteredEquityBonusData.Where(x=> Convert.ToDateTime(x.EX_DT) > DateTime.Today.AddDays(2)).Select(x => new DividendData { SYMBOL = x.SYMBOL, EX_DT = x.EX_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
                stocksWhichHaveBeenAnnounced.AddRange(bonusUpcomingStockData);
            }
            catch(Exception ex) { }

            foreach (int month in Enumerable.Range(0, 300).ToList())
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
                    var delistFilteredEquityBonusData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("RIGHTS") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                    if (delistFilteredEquityBonusData.Count() > 0)
                    {
                        var bonusStockData = delistFilteredEquityBonusData.Select(x => new DividendData { SYMBOL = x.SYMBOL, EX_DT = x.EX_DT.Replace("/", "-").Trim(), PURPOSE = x.PURPOSE.Trim() });
                        last20YearsStockData.Add(stringKey, bonusStockData.ToList());
                    }                        
                }
                catch (Exception ex) { continue; }
            }

            Dictionary<string, int> daysForWhichBonusIsStable = new Dictionary<string, int>();
            foreach (var currentYearBonusData in last20YearsStockData)
            {
                foreach (var bonusData in currentYearBonusData.Value.Where(x => x.EX_DT.Trim() != ""))
                {
                    bool IsYearDateFormat = DateTime.TryParseExact(bonusData.EX_DT, "yyyy-MM-dd", CultureInfo.InvariantCulture,DateTimeStyles.None, out DateTime result);
                    if(!IsYearDateFormat)
                        DateTime.TryParseExact(bonusData.EX_DT, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    bonusData.EX_DT = result.ToString("dd-MM-yyyy");
                }
            }
            var historicalRightsIssueData = last20YearsStockData.OrderByDescending(x => x.Value.Count);
            string json = JsonConvert.SerializeObject(historicalRightsIssueData, Formatting.Indented);
            Console.WriteLine("Historical Share Rights Data: -----\n" + json + "\n-------------\n");
            var upcomingStocksWithGenuineRights = stocksWhichHaveBeenAnnounced.Where(x => historicalRightsIssueData.Select(y => y.Key).Contains(x.SYMBOL));
            string upcomingGenuineStockRights = JsonConvert.SerializeObject(upcomingStocksWithGenuineRights, Formatting.Indented);
            Console.WriteLine("Upcoming Genuine Stock Bonus Data: -----\n" + upcomingGenuineStockRights + "\n-------------\n");
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
