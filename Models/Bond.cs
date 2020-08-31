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
        public Nullable<double> TradingFactor { get; set; }
        public Nullable<double> PricingFactor { get; set; }
        public string ISIN { get; set; }
        public string BBGTicker { get; set; }
        public string BBGUniqueID { get; set; }
        public string CUSIP { get; set; }
        public string SEDOL { get; set; }
        public string FirstCouponDate { get; set; }
        public string Cap { get; set; }
        public string FloorValue { get; set; }
        public Nullable<int> CouponFrequency { get; set; }
        public Nullable<double> Coupon { get; set; }
        public string CouponType { get; set; }
        public string Spread { get; set; }
        public string CallableFlag { get; set; }
        public string FixtoFloatFlag { get; set; }
        public string PutableFlag { get; set; }
        public string IssueDate { get; set; }
        public string LastResetDate { get; set; }
        public string Maturity { get; set; }
        public Nullable<double> CallNotificationMaxDays { get; set; }
        public Nullable<double> PutNotificationMaxDays { get; set; }
        public string PenultimateCouponDate { get; set; }
        public string ResetFrequency { get; set; }
        public string HasPosition { get; set; }
        public Nullable<double> MacaulayDuration { get; set; }
        public Nullable<double> Volatility_30D { get; set; }
        public Nullable<double> Volatility_90D { get; set; }
        public Nullable<double> Convexity { get; set; }
        public Nullable<double> AverageVolume_30Day { get; set; }
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
        public Nullable<double> PutPrice { get; set; }
        public Nullable<double> AskPrice { get; set; }
        public Nullable<double> HighPrice { get; set; }
        public Nullable<double> LowPrice { get; set; }
        public Nullable<double> OpenPrice { get; set; }
        public Nullable<double> Volume { get; set; }
        public Nullable<double> BidPrice { get; set; }
        public Nullable<double> LastPrice { get; set; }
        public string CallDate { get; set; }
        public Nullable<double> CallPrice { get; set; }
    }
    
}