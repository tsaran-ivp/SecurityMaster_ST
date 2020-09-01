using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SecurityMaster_ST.Models
{
    public class Bond
    {   
        public int ID { get; set; }
        public string SecurityDescription { get; set; }
        public string SecurityName { get; set; }
        public string AssetType { get; set; }
        public string InvestmentType { get; set; }
        public string TradingFactor { get; set; }
        public string PricingFactor { get; set; }
        public string ISIN { get; set; }
        public string BBGTicker { get; set; }
        public string BBGUniqueID { get; set; }
        public string CUSIP { get; set; }
        public string SEDOL { get; set; }
        public string FirstCouponDate { get; set; }
        public string Cap { get; set; }
        public string FloorValue { get; set; }
        public string CouponFrequency { get; set; }
        public string Coupon { get; set; }
        public string CouponType { get; set; }
        public string Spread { get; set; }
        public string CallableFlag { get; set; }
        public string FixtoFloatFlag { get; set; }
        public string PutableFlag { get; set; }
        public string IssueDate { get; set; }
        public string LastResetDate { get; set; }
        public string Maturity { get; set; }
        public string CallNotificationMaxDays { get; set; }
        public string PutNotificationMaxDays { get; set; }
        public string PenultimateCouponDate { get; set; }
        public string ResetFrequency { get; set; }
        public string HasPosition { get; set; }
        public string MacaulayDuration { get; set; }
        public string Volatility_30D { get; set; }
        public string Volatility_90D { get; set; }
        public string Convexity { get; set; }
        public string AverageVolume_30Day { get; set; }
        public string PFAssetClass { get; set; }
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
        public string BloombergIndustryGroup { get; set; }
        public string BloombergIndustrySubGroup { get; set; }
        public string BloombergIndustrySector { get; set; }
        public string CountryofIssuance { get; set; }
        public string IssueCurrency { get; set; }
        public string IssuerName { get; set; }
        public string RiskCurrency { get; set; }
        public string PutDate { get; set; }
        public string PutPrice { get; set; }
        public string AskPrice { get; set; }
        public string HighPrice { get; set; }
        public string LowPrice { get; set; }
        public string OpenPrice { get; set; }
        public string Volume { get; set; }
        public string BidPrice { get; set; }
        public string LastPrice { get; set; }
        public string CallDate { get; set; }
        public string CallPrice { get; set; }
    }
    
}