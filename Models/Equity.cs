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
        public float SharesPerADR { get; set; }
        public string IPODate { get; set; }
        public string PricingCurrency { get; set; }
        public int SettleDays { get; set; }
        public float OutstandingShares { get; set; }
        public float VotingRightsPerShare { get; set; }
        public float AverageVolume { get; set; }
        public float Beta { get; set; }
        public float ShortInterest { get; set; }
        public float ReturnYTD { get; set; }
        public float Volatility { get; set; }
        public string AssetClass { get; set; }
        public string PFCountry { get; set; }
        public string PFCreditRating { get; set; }
        public string PFCurrency { get; set; }
        public string PFInstrument { get; set; }
        public string PFLiquidityProfile { get; set; }
        public string PFMaturity { get; set; }
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
        public float OpenPrice { get; set; }
        public float ClosePrice { get; set; }
        public float Volume { get; set; }
        public float LastPrice { get; set; }
        public float AskPrice { get; set; }
        public float BidPrice { get; set; }
        public float PERatio { get; set; }
        public string DividendDeclareDate { get; set; }
        public string DividendExDate { get; set; }
        public string DividendRecordDate { get; set; }
        public string DividendPayDate { get; set; }
        public float DividendAmount { get; set; }
        public string Frequency { get; set; }
        public string DividendType { get; set; }
    }
}