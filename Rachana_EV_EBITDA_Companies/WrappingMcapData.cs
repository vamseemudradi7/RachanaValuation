namespace EV_EBITDA
{
    public class McapData
    {
       public string Symbol   {get;set;}
       public long NoOfShares   {get;set;}
       public string CompanyName   {get;set;}
       public decimal LastPrice     {get;set;}
       public decimal PerChange     {get;set;}
       public decimal MarketCap     {get;set;}
       public decimal ScTtm         {get;set;}
       public decimal Perform1yr    {get;set;}
       public decimal PriceBook { get; set; }
    }
}
