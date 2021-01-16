namespace BhavCopy.NSE
{
    public class DividendData
    {
        public string SERIES { get; set; }
        public string SYMBOL { get; set; }
        public string EX_DT { get; set; }
        public string PURPOSE { get; set; }
    }

    public class DividendDataOutput
    {    
        public string ExDividendDate { get; set; }    
        public string DividendValue { get; set; }
        public decimal? ClosePriceOnExDMinus1Date { get; set; }
        public decimal? ClosePriceOnExDPlus2Date { get; set; }
    }
}
