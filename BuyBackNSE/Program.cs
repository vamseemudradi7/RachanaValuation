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
            List<DividendData> delistFilteredEquityBuybackData = new List<DividendData>();
            List<DividendData> bonusUpcomingStockData = new List<DividendData>();
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
                delistFilteredEquityBuybackData = bonusAndDividendData.Where(x => x.PURPOSE != null && x.PURPOSE.Contains("BUYBACK") && x.SERIES != null && x.SERIES == "EQ" && x.SYMBOL != null && !delistedStockSymbols.Contains(x.SYMBOL) && !toBeDelistedStockSymbols.Contains(x.SYMBOL)).ToList();
                DateTime result;
                foreach (var buyBackData in delistFilteredEquityBuybackData)
                {
                    bool IsYearDateFormatForRecordDate = DateTime.TryParseExact(buyBackData.EX_DT, "yyyy-dd-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    if (!IsYearDateFormatForRecordDate)
                        DateTime.TryParseExact(buyBackData.EX_DT, "dd-MM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
                    if(result > DateTime.Today.AddDays(2))
                        bonusUpcomingStockData.Add(new DividendData { SYMBOL = buyBackData.SYMBOL , EX_DT = buyBackData.EX_DT, PURPOSE = buyBackData.PURPOSE });
                }
                stocksWhichHaveBeenAnnounced.AddRange(bonusUpcomingStockData);
            }
            catch(Exception ex) { }
            string json = JsonConvert.SerializeObject(delistFilteredEquityBuybackData, Formatting.Indented);
            Console.WriteLine("Historical Share Bonus Data: -----\n" + json + "\n-------------\n");
            string upcomingGenuineStockBonuses = JsonConvert.SerializeObject(bonusUpcomingStockData, Formatting.Indented);
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
