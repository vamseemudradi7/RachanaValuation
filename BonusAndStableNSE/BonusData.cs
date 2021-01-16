namespace BhavCopy.NSE
{
    public class BonusData
    {
        public string SERIES { get; set; }
        public string SYMBOL { get; set; }
        public string RECORD_DT { get; set; }
        public string PURPOSE { get; set; }
    }

    public class BonusDataOutput
    {    
        public string RecordDate { get; set; }    
        public string BonusRatio { get; set; }
        public decimal? ClosePriceOnRMinus6Date { get; set; }
        public decimal? ClosePriceOnRPlus1Date { get; set; }
    }
}
