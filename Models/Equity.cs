using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SecurityMaster_ST.Models
{
    public class Equity
    {
        public int ID { get; set; }
        public string SecurityName { get; set; }
        public string SecurityDescription { get; set; }
        public string HasPosition { get; set; }
        public string ActiveSecurity { get; set; }
        public int LotSize { get; set; }
        public string BBGUniqueName { get; set; }
        public string CUSIP { get; set; }
        public string ISIN { get; set; }
        public string SEDOL { get; set; }
        public string BBGTicker { get; set; }
        public string BBGUID { get; set; }
        public string BBGGlobalID { get; set; }
        public string TickerAndExchange { get; set; }
        public string ADRFlag { get; set; }
        public string ADRUnderlyingTicker { get; set; }
        public string ADRUnderlyingCurrency { get; set; }
        public Nullable<float> SharesPerADR { get; set; }
        public Nullable<System.DateTime> IPODate { get; set; }
        public string PricingCurrency { get; set; }
        public Nullable<int> SettleDays { get; set; }
        public Nullable<float> OutstandingShares { get; set; }
        public Nullable<float> VotingRightsPerShare { get; set; }
        public Nullable<float> AverageVolume { get; set; }
        public Nullable<float> Beta { get; set; }
        public Nullable<float> ShortInterest { get; set; }
        public Nullable<float> ReturnYTD { get; set; }
        public Nullable<float> Volatility { get; set; }
        public string AssetClass { get; set; }
        public string PFCountry { get; set; }
        public string PFCreditRating { get; set; }
        public string PFCurrency { get; set; }
        public string PFInstrument { get; set; }
        public string PFLiquidityProfile { get; set; }
        public Nullable<System.DateTime> PFMaturity { get; set; }
        public string PFNAICSCode { get; set; }
        public string PFRegion { get; set; }
        public string PFSector { get; set; }
        public string PFSubAssetClass { get; set; }
        public string CountryOfIssuance { get; set; }
        public string Exchange { get; set; }
        public string IssuerName { get; set; }
        public string IssueCurrency { get; set; }
        public string TradingCurrency { get; set; }
        public string BBGIndustrySubGroup { get; set; }
        public string BBGIndustryGroup { get; set; }
        public string BBGSector { get; set; }
        public string CountryOfIncorporation { get; set; }
        public string RiskCurrency { get; set; }
        public Nullable<float> OpenPrice { get; set; }
        public Nullable<float> ClosePrice { get; set; }
        public Nullable<float> Volume { get; set; }
        public Nullable<float> LastPrice { get; set; }
        public Nullable<float> AskPrice { get; set; }
        public Nullable<float> BidPrice { get; set; }
        public Nullable<float> PERatio { get; set; }
        public Nullable<System.DateTime> DividendDeclareDate { get; set; }
        public Nullable<System.DateTime> DividendExDate { get; set; }
        public Nullable<System.DateTime> DividendRecordDate { get; set; }
        public Nullable<System.DateTime> DividendPayDate { get; set; }
        public Nullable<float> DividendAmount { get; set; }
        public string Frequency { get; set; }
        public string DividendType { get; set; }
    }
}