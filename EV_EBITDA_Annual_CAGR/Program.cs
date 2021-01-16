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

            // In the below computation from line # 41 till # 65, we get list of all stocks with their (High-Low-Open-Close-Traded Qty-Net Value against each stock and filter out delisted and suspended stocks. These stocks however have upper circuit and lower circuit stocks)
            // The variable : last100DaysStockData contains this data. 
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
            // Calculate the Exponential VWap per stocks taken exponentially for the last 100 days.
            // Example :  If Yesterday's VWap = 100 and Day Before Yesterday's Vwap = 110, then exponential VWap for 2 days,
            // then taking a sample weight of 100, we have ((100 * 100) + (99 * 110))/(100 +99)) = 104.97

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
                        // Right now, its a headache to run this progrm as it takes very very long amount of time to execute. (Maybe like 1 .5 days continuiusly due to waiting for 2 mins every other request.
                        // Issue is with google returning 429 (Too many requests from a single IP error whenever I change this value below 2 mins / 2 requests)
                        try
                        {
                            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                            if (response.StatusCode == HttpStatusCode.OK)
                            {
                                Stream receiveStream = response.GetResponseStream();
                                StreamReader readStream = string.IsNullOrWhiteSpace(response.CharacterSet) ? new StreamReader(receiveStream) : new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                                string data = readStream.ReadToEnd();
                                // MoneyControl Symbol Code is different than the NSE Symbol.
                                var moneyControlSymbolCode = data.Split("https://www.moneycontrol.com/stocks/company_info/stock_news.php%3Fsc_id%3D")[1].Split("%")[0];
                                response.Close();
                                readStream.Close();

                                // Get the consolidated balance sheet from money control.
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
                                // We need the above values like Debt and Cash&Cash Equivalents(CCE). We need the above data to capture the Enterprise Value.
                                // Net Debt = Debt - CCE [ Debt - CCE ]
                                // Enterprise Value (EV) = [NetDebt + MarketCap] in Crores.

                                string urlAddressForConsolidatedAnnualResults = "https://www.moneycontrol.com/financials/britanniaindustries/results/consolidated-yearly/" + moneyControlSymbolCode + "#" + moneyControlSymbolCode;
                                HttpWebRequest requestForConsolidatedAnnualResults = (HttpWebRequest)WebRequest.Create(urlAddressForConsolidatedAnnualResults);
                                requestForConsolidatedAnnualResults.Headers.Add("User-Agent", @"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5)\AppleWebKit / 537.36(KHTML, like Gecko) Safari / 537.36");
                                HttpWebResponse responseForConsolidatedAnnualResults = (HttpWebResponse)requestForConsolidatedAnnualResults.GetResponse();
                                Stream receiveStreamForConsolidatedAnnualResults = responseForConsolidatedAnnualResults.GetResponseStream();
                                StreamReader readStreamForConsolidatedAnnualResults = string.IsNullOrWhiteSpace(responseForConsolidatedAnnualResults.CharacterSet) ? new StreamReader(receiveStreamForConsolidatedAnnualResults) : new StreamReader(receiveStreamForConsolidatedAnnualResults, Encoding.GetEncoding(responseForConsolidatedAnnualResults.CharacterSet));
                                string consolidatedAnnualResultsData = readStreamForConsolidatedAnnualResults.ReadToEnd();

                                if (consolidatedAnnualResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax").Length <= 1)
                                    continue;

                                // Year1 'Profit Before Tax, Depreciation & Amortization' = Last March 2020 PBT (After April 2021, we would be getting March 2021 PBT automatically in this code.)
                                var year1PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[1], "[^0-9.]", "");

                                // Year2 'Profit Before Tax, Depreciation & Amortization' = March 2019 PBT (After April 2021, we would be getting March 2020 PBT automatically in this code.)
                                var year2PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[2], "[^0-9.]", "");

                                // Year3 'Profit Before Tax, Depreciation & Amortization' = March 2018 PBT (After April 2021, we would be getting March 2019 PBT automatically in this code.)
                                var year3PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[3], "[^0-9.]", "");

                                // Year4 'Profit Before Tax, Depreciation & Amortization' = March 2017 PBT (After April 2021, we would be getting March 2018 PBT automatically in this code.)
                                var year4PLBeforeTaxDepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("P/L Before Int., Excpt. Items &amp; Tax")[1].Split("<td>")[4], "[^0-9.]", "");

                                Decimal.TryParse(year1PLBeforeTaxDepreciationString, out decimal year1PLBeforeTaxDepreciation);
                                Decimal.TryParse(year2PLBeforeTaxDepreciationString, out decimal year2PLBeforeTaxDepreciation);
                                Decimal.TryParse(year3PLBeforeTaxDepreciationString, out decimal year3PLBeforeTaxDepreciation);
                                Decimal.TryParse(year4PLBeforeTaxDepreciationString, out decimal year4PLBeforeTaxDepreciation);

                                decimal PLBeforeTaxDepreciationLast4YearsAggregate = year1PLBeforeTaxDepreciation + year2PLBeforeTaxDepreciation + year3PLBeforeTaxDepreciation + year4PLBeforeTaxDepreciation;

                                var year1DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("Depreciation")[1].Split("<td>")[1], "[^0-9.]", "");
                                var year2DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("Depreciation")[1].Split("<td>")[2], "[^0-9.]", "");
                                var year3DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("Depreciation")[1].Split("<td>")[3], "[^0-9.]", "");
                                var year4DepreciationString = System.Text.RegularExpressions.Regex.Replace(consolidatedAnnualResultsData.Split("Depreciation")[1].Split("<td>")[4], "[^0-9.]", "");

                                Decimal.TryParse(year1DepreciationString, out decimal year1Depreciation);
                                Decimal.TryParse(year2DepreciationString, out decimal year2Depreciation);
                                Decimal.TryParse(year3DepreciationString, out decimal year3Depreciation);
                                Decimal.TryParse(year4DepreciationString, out decimal year4Depreciation);

                                decimal DepreciationLast4YearsAggregate = year1Depreciation + year2Depreciation + year3Depreciation + year4Depreciation;

                                var EBITDAYear1 = (year1PLBeforeTaxDepreciation + year1Depreciation); // EBITDA for 2020
                                var EBITDAYear4 = (year4PLBeforeTaxDepreciation + year4Depreciation); // EBITDA for 2017

                                // CAGR of the company when comparing the results of 2020 and comparing the results of 2017.
                                // Formula for CAGR is ("2020 EBITDA" / "2017 EBITDA")^(1/n) - 1; // So, here, n = number of years difference = 3 (2020 - 2017).
                                var latest3yearEBITDACAGR = Math.Pow((double)EBITDAYear1 / (double)EBITDAYear4, 0.334) - 1;

                                // Consolidated EBITDA for the last 4 years.
                                decimal consolidatedEBITDA = PLBeforeTaxDepreciationLast4YearsAggregate + DepreciationLast4YearsAggregate;

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
                                    singleMCapData.NoOfShares = Convert.ToInt64((singleMCapData.MarketCap * Convert.ToDecimal(Math.Pow(10,7))) / (decimal)singleMCapData.LastPrice); // * 10^7  = 1 Crore. So, since the money control api & site has the Market Cap in crores, we multiply by 10^7. Ultimately we need to get the # of shares of the company. It can also be gotten from Shareholder pattern in moneycontrol site but could not get a way to do it programatically.
                                    singleMCapData.ExponentialVwapValue = exponentialVwapPerStock[delistFilteredStock.SYMBOL];
                                    singleMCapData.ComputedMarketCapInCrores = singleMCapData.ExponentialVwapValue * singleMCapData.NoOfShares / (decimal)Math.Pow(10,7);
                                    singleMCapData.ComputedDebt = debt;
                                    singleMCapData.ComputedCCE = cashAndCashEquivalents;
                                    singleMCapData.ComputedEnterpriseValue = singleMCapData.ComputedMarketCapInCrores + debt - cashAndCashEquivalents; // Get the current EV as per exponential VWAP.
                                    singleMCapData.ComputedEBITDA = consolidatedEBITDA;
                                    singleMCapData.ComputedNetDebt = debt - cashAndCashEquivalents;
                                    singleMCapData.EVEBITDARatio = singleMCapData.ComputedEnterpriseValue / singleMCapData.ComputedEBITDA; // (EV/EBITDA) ratio is important to calucate the estimated share price next quarter.
                                    singleMCapData.LatestYearEBITDACAGR = Convert.ToDecimal(latest3yearEBITDACAGR);
                                    var estimatedEBITDA = singleMCapData.ComputedEBITDA * (1 + (singleMCapData.LatestYearEBITDACAGR / 100m)); // Estimated EBITDA next year would be = [(Calculated EBITDA this year) * ( 1 + CAGR)]
                                    var expectedEV = singleMCapData.EVEBITDARatio * estimatedEBITDA; // (Static EV/EBITDA calculated in line # 240 * Estimated EBITDA for next year calcuated in above line)
                                    var expectedEquityValueInRupees = (expectedEV - singleMCapData.ComputedNetDebt) * Convert.ToDecimal(Math.Pow(10,7));
                                    singleMCapData.ExpectedPriceValueNextYear = expectedEquityValueInRupees / (decimal)singleMCapData.NoOfShares; // If # of shares is staying constant next year as well, then expected share value next year would be = [Expected Equity Value / # Of Shares]
                                    
                                    //Rachana Ranade actually mentioned till above point to evaluate the expected price value for March 2021 based on the EBITDA CAGR of 2017 (or of any previous year for that matter) AND (current share value)
                                    if (singleMCapData.ExpectedPriceValueNextYear >= 1.15m * singleMCapData.LastPrice)
                                    {
                                        // Expected Price / Todays Price suggests the % increase when compared to today's price. If expected price is greater than 15% of today's price, then it would be a good valuation to buy according to me . Rachana didnt suggest anything about this though. 
                                        singleMCapData.ExpectedPriceLastPriceRatio = singleMCapData.ExpectedPriceValueNextYear / singleMCapData.LastPrice;
                                        MCapData.Add(singleMCapData);
                                    }
                                }
                            }
                        }
                        catch (Exception ex2)
                        {
                            // Exception occurs if the Profit Loss statements arent available in Money Control for the given company.
                        }
                    }
                }
                catch (Exception ex3) { }
            }
            // PRogram is a little error prone due to Google not returning the right MoneyControl Code. However, this error is easily caught 
            // by checking the CompanyName & Company Symbol in the below response. Only if its matching, then its an accurate result as per this program.
           Console.WriteLine(JsonConvert.SerializeObject(MCapData.OrderByDescending(x=>x.ExpectedPriceLastPriceRatio)));
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
