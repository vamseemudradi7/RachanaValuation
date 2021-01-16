using System.Collections.Generic;

namespace EV_EBITDA
{
    public class McapData
    {
        public string Symbol { get; set; }
        public decimal ExponentialVwapValue { get; set; }
        public decimal ComputedMarketCapInCrores { get; set; }
        public decimal ComputedEBITDA { get; set; }
        public decimal ExpectedPriceLastPriceRatio { get; set; }
        public decimal ComputedNetDebt { get; set; }
        public decimal LatestYearEBITDACAGR { get; set; }
        public decimal ExpectedPriceValueNextYear { get; set; }
        public decimal LatestCAGR { get; set; }
        public decimal EVEBITDARatio { get; set; }
        public decimal ComputedDebt { get; set; }
        public decimal ComputedCCE { get; set; }
        public decimal ComputedEnterpriseValue { get; set; }
        public long NoOfShares { get; set; }
        public string CompanyName { get; set; }
        public decimal LastPrice { get; set; }
        public string PerChange { get; set; }
        public decimal MarketCap { get; set; }
        public string ScTtm { get; set; }
        public string Perform1yr { get; set; }
        public string PriceBook { get; set; }
    }

    public class WrappingMcapData
    {
        public int Success { get; set; }
        public List<McapData> Data { get; set; }
    }    
}
