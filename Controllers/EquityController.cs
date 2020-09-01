using SecurityMaster_ST.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.UI.WebControls;
using ExcelDataReader;

namespace SecurityMaster_ST.Controllers
{
    public class EquityController : ApiController
    {
        string connectionString = "data source=192.168.0.104\\sql_express1,63862; database=SecurityMaster_ST; user=Sa; password=valley@1234"; //Hardcoded Connection String

        public HttpResponseMessage Get()
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                DataTable table = new DataTable();
                string spName = "Equity.usp_ViewRecords";
                SqlCommand cmd = new SqlCommand(spName,con);
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(table);
                return Request.CreateResponse(HttpStatusCode.OK, table);

            }
        }

        public string Put(Equity equity)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    DataTable table = new DataTable();
                    string spName = "Equity.UpdateRecordsInDB";
                    SqlCommand cmd = new SqlCommand(spName, con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ID", equity.ID);
                    cmd.Parameters.AddWithValue("@SecurityNameVar", equity.SecurityName);
                    cmd.Parameters.AddWithValue("@SecurityDescriptionVar", equity.SecurityDescription);
                    cmd.Parameters.AddWithValue("@HasPositionVar", equity.HasPosition);
                    cmd.Parameters.AddWithValue("@ActiveSecurityVar", equity.ActiveSecurity);
                    cmd.Parameters.AddWithValue("@LotSizeVar", equity.LotSize);
                    cmd.Parameters.AddWithValue("@BBGUniqueNameVar", equity.BBGUniqueName);

                    cmd.Parameters.AddWithValue("@CUSIPVar", equity.CUSIP);
                    cmd.Parameters.AddWithValue("@ISINVar", equity.ISIN);
                    cmd.Parameters.AddWithValue("@SEDOLVar", equity.SEDOL);
                    cmd.Parameters.AddWithValue("@BBGtickerVar", equity.BBGTicker);
                    cmd.Parameters.AddWithValue("@BBGUIDVar", equity.BBGUID);
                    cmd.Parameters.AddWithValue("@BBGGlobalIDVar", equity.BBGGlobalID);
                    cmd.Parameters.AddWithValue("@BBGTickerandExchangeVar", equity.TickerAndExchange);

                    cmd.Parameters.AddWithValue("@ADRFlagVar", equity.ADRFlag);
                    cmd.Parameters.AddWithValue("@ADRUnderlyingTickerVar", equity.ADRUnderlyingTicker);
                    cmd.Parameters.AddWithValue("@ADRUnderlyingCurrencyVar", equity.ADRUnderlyingCurrency);
                    cmd.Parameters.AddWithValue("@SharesPerADRVar", equity.SharesPerADR);
                    cmd.Parameters.AddWithValue("@IPODateVar", equity.IPODate);

                    cmd.Parameters.AddWithValue("@PricingCurrencyVar", equity.PricingCurrency);
                    cmd.Parameters.AddWithValue("@SettleDaysVar", equity.SettleDays);
                    cmd.Parameters.AddWithValue("@OutstandingSharesVar", equity.OutstandingShares);
                    cmd.Parameters.AddWithValue("@VotingRightsPerShareVar", equity.VotingRightsPerShare);

                    cmd.Parameters.AddWithValue("@AverageVolume20DVar", equity.AverageVolume);
                    cmd.Parameters.AddWithValue("@BetaVar", equity.Beta);
                    cmd.Parameters.AddWithValue("@ShortInterestVar", equity.ShortInterest);
                    cmd.Parameters.AddWithValue("@ReturnYTDVar", equity.ReturnYTD);
                    cmd.Parameters.AddWithValue("@Volatility90DayVar", equity.Volatility);


                    cmd.Parameters.AddWithValue("@AssetClassVar", equity.AssetClass);
                    cmd.Parameters.AddWithValue("@PFCountryVar", equity.PFCountry);
                    cmd.Parameters.AddWithValue("@PFCreditRatingVar", equity.PFCreditRating);
                    cmd.Parameters.AddWithValue("@PFCurrencyVar", equity.PFCurrency);
                    cmd.Parameters.AddWithValue("@PFInstrumentVar", equity.PFInstrument);
                    cmd.Parameters.AddWithValue("@PFLiquidityProfileVar", equity.PFLiquidityProfile);
                    cmd.Parameters.AddWithValue("@PFMaturityVar", equity.PFMaturity);
                    cmd.Parameters.AddWithValue("@PFNAISCCodeVar", equity.PFNAICSCode);

                    cmd.Parameters.AddWithValue("@PFRegionVar", equity.PFRegion);
                    cmd.Parameters.AddWithValue("@PFSectorVar", equity.PFSector);
                    cmd.Parameters.AddWithValue("@PFSubAssetClassVar", equity.PFSubAssetClass);

                    cmd.Parameters.AddWithValue("@CountryOfIssuanceVar", equity.CountryOfIssuance);
                    cmd.Parameters.AddWithValue("@ExchangeVar", equity.Exchange);
                    cmd.Parameters.AddWithValue("@IssuerNameVar", equity.IssuerName);
                    cmd.Parameters.AddWithValue("@IssueCurrencyVar", equity.IssueCurrency);
                    cmd.Parameters.AddWithValue("@TradingCurrencyVar", equity.TradingCurrency);
                    cmd.Parameters.AddWithValue("@BBGIndustrySubGroupVar", equity.BBGIndustrySubGroup);
                    cmd.Parameters.AddWithValue("@BBGIndustryGroupVar", equity.BBGIndustryGroup);

                    cmd.Parameters.AddWithValue("@BBGSectorVar", equity.BBGSector);
                    cmd.Parameters.AddWithValue("@CountryOfIncorporationVar", equity.CountryOfIncorporation);
                    cmd.Parameters.AddWithValue("@RiskCurrencyVar", equity.RiskCurrency);
                    cmd.Parameters.AddWithValue("@OpenPriceVar", equity.OpenPrice);
                    cmd.Parameters.AddWithValue("@ClosePriceVar", equity.ClosePrice);
                    cmd.Parameters.AddWithValue("@VolumeVar", equity.Volume);
                    cmd.Parameters.AddWithValue("@LastPriceVar", equity.LastPrice);
                    cmd.Parameters.AddWithValue("@AskPriceVar", equity.AskPrice);
                    cmd.Parameters.AddWithValue("@BidPriceVar", equity.BidPrice);
                    cmd.Parameters.AddWithValue("@PERatioVar", equity.PERatio);


                    cmd.Parameters.AddWithValue("@DividendDeclareDateVar", equity.DividendDeclareDate);
                    cmd.Parameters.AddWithValue("@DividendExDateVar", equity.DividendExDate);
                    cmd.Parameters.AddWithValue("@DividendRecordDateVar", equity.DividendRecordDate);
                    cmd.Parameters.AddWithValue("@DividendPayDateVar", equity.DividendPayDate);
                    cmd.Parameters.AddWithValue("@DividendAmountVar", equity.DividendAmount);
                    cmd.Parameters.AddWithValue("@FrequencyVar", equity.Frequency);
                    cmd.Parameters.AddWithValue("@DividendTypeVar", equity.DividendType);
                    con.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);

                    return "Update Successful";
                }

            }
            catch (Exception e)
            {
                return $"Couldn't Update due to {e.Message}";
            }
        }

        public string Delete(int ID)
        {
            try
            {

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    DataTable table = new DataTable();
                    string spName = "Equity.usp_DeleteFromDB";
                    SqlCommand cmd = new SqlCommand(spName, con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@ID", ID);
                    con.Open();
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                    dataAdapter.Fill(table);


                    return "Security Deleted";
                }


            }
            catch (Exception e)
            {
                return $"Couldn't Delete due to {e.Message}";
            }
        }


        public string Post()
        {
            try
            {
                var httRequest = HttpContext.Current.Request;
                var postedFile = httRequest.Files[0];
                string filename = postedFile.FileName;
                var physicalpath = HttpContext.Current.Server.MapPath("~/EquityFiles/" + filename);
                postedFile.SaveAs(physicalpath);
                string ext = Path.GetExtension(postedFile.FileName);
                //check file extension if excel or csv
                if (ext.ToLower() == ".xls" || ext.ToLower() == ".xlsx")
                {

                    Stream stream = postedFile.InputStream;
                    IExcelDataReader reader = null;
                    //check which type of excel file
                    if (ext.ToLower() == ".xls")
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (ext.ToLower() == ".xlsx")
                    { reader = ExcelReaderFactory.CreateOpenXmlReader(stream); }
                  

                    DataSet dataSet = reader.AsDataSet();
                    reader.Close();
                    dataSet.Tables[0].Rows.RemoveAt(0);


                    //for each new record call the procedure and insert data

                    foreach (DataRow dr in dataSet.Tables[0].Rows)
                    {

                        using (SqlConnection con = new SqlConnection(connectionString))
                        {
                            //Execute Reader    
                            DataTable table = new DataTable();
                            string command = "Equity.usp_InsertIntoDB";
                            SqlCommand cmd = new SqlCommand(command, con);
                            cmd.CommandType = CommandType.StoredProcedure;


                            cmd.Parameters.AddWithValue("@SecurityNameVar", dr["Security Name"].ToString());
                            cmd.Parameters.AddWithValue("@SecurityDescriptionVar", dr["Security Description"].ToString());
                            cmd.Parameters.AddWithValue("@HasPositionVar", dr["Has  Position"].ToString());
                            cmd.Parameters.AddWithValue("@ActiveSecurityVar", dr["Active Security"].ToString());
                            cmd.Parameters.AddWithValue("@LotSizeVar", dr["Lot Size"]);
                            cmd.Parameters.AddWithValue("@BBGUniqueNameVar", dr["BBG Unique Name"].ToString());

                            cmd.Parameters.AddWithValue("@CUSIPVar", dr["CUSIP"].ToString());
                            cmd.Parameters.AddWithValue("@ISINVar", dr["ISIN"].ToString());
                            cmd.Parameters.AddWithValue("@SEDOLVar", dr["SEDOL"].ToString());
                            cmd.Parameters.AddWithValue("@BBGtickerVar", dr["Bloomberg Ticker"].ToString());
                            cmd.Parameters.AddWithValue("@BBGUIDVar", dr["Bloomberg Unique ID"].ToString());
                            cmd.Parameters.AddWithValue("@BBGGlobalIDVar", dr["BBG Global ID"].ToString());
                            cmd.Parameters.AddWithValue("@BBGTickerandExchangeVar", dr["Ticker and Exchange"].ToString());

                            cmd.Parameters.AddWithValue("@ADRFlagVar", dr["Is ADR Flag"].ToString());
                            cmd.Parameters.AddWithValue("@ADRUnderlyingTickerVar", dr["ADR Underlying Ticker"].ToString());
                            cmd.Parameters.AddWithValue("@ADRUnderlyingCurrencyVar", dr["ADR Underlying Currency"].ToString());
                            cmd.Parameters.AddWithValue("@SharesPerADRVar", dr["Shares Per ADR"]);
                            cmd.Parameters.AddWithValue("@IPODateVar", dr["IPO Date"]);

                            cmd.Parameters.AddWithValue("@PricingCurrencyVar", dr["Pricing Currency"].ToString());
                            cmd.Parameters.AddWithValue("@SettleDaysVar", dr["Settle Days"]);
                            cmd.Parameters.AddWithValue("@OutstandingSharesVar", dr["Total Shares Outstanding"]);
                            cmd.Parameters.AddWithValue("@VotingRightsPerShareVar", dr["Voting Rights Per Share"]);

                            cmd.Parameters.AddWithValue("@AverageVolume20DVar", dr["Average Volume - 20D"]);
                            cmd.Parameters.AddWithValue("@BetaVar", dr["Beta"]);
                            cmd.Parameters.AddWithValue("@ShortInterestVar", dr["Short Interest"]);
                            cmd.Parameters.AddWithValue("@ReturnYTDVar", dr["Return - YTD"]);
                            cmd.Parameters.AddWithValue("@Volatility90DayVar", dr["Volatility - 90D"]);


                            cmd.Parameters.AddWithValue("@AssetClassVar", dr["PF Asset Class"].ToString());
                            cmd.Parameters.AddWithValue("@PFCountryVar", dr["PF Country"].ToString());
                            cmd.Parameters.AddWithValue("@PFCreditRatingVar", dr["PF Credit Rating"].ToString());
                            cmd.Parameters.AddWithValue("@PFCurrencyVar", dr["PF Currency"].ToString());
                            cmd.Parameters.AddWithValue("@PFInstrumentVar", dr["PF Instrument"].ToString());
                            cmd.Parameters.AddWithValue("@PFLiquidityProfileVar", dr["PF Liquidity Profile"].ToString());
                            cmd.Parameters.AddWithValue("@PFMaturityVar", dr["PF Maturity"]);
                            cmd.Parameters.AddWithValue("@PFNAISCCodeVar", dr["PF NAICS Code"].ToString());

                            cmd.Parameters.AddWithValue("@PFRegionVar", dr["PF Region"].ToString());
                            cmd.Parameters.AddWithValue("@PFSectorVar", dr["PF Sector"].ToString());
                            cmd.Parameters.AddWithValue("@PFSubAssetClassVar", dr["PF Sub Asset Class"].ToString());

                            cmd.Parameters.AddWithValue("@CountryOfIssuanceVar", dr["Country of Issuance"].ToString());
                            cmd.Parameters.AddWithValue("@ExchangeVar", dr["Exchange"].ToString());
                            cmd.Parameters.AddWithValue("@IssuerNameVar", dr["Issuer"].ToString());
                            cmd.Parameters.AddWithValue("@IssueCurrencyVar", dr["Issue Currency"].ToString());
                            cmd.Parameters.AddWithValue("@TradingCurrencyVar", dr["Trading Currency"].ToString());
                            cmd.Parameters.AddWithValue("@BBGIndustrySubGroupVar", dr["BBG Industry Sub Group"].ToString());
                            cmd.Parameters.AddWithValue("@BBGIndustryGroupVar", dr["Bloomberg Industry Group"].ToString());

                            cmd.Parameters.AddWithValue("@BBGSectorVar", dr["Bloomberg Sector"].ToString());
                            cmd.Parameters.AddWithValue("@CountryOfIncorporationVar", dr["Country of Incorporation"].ToString());
                            cmd.Parameters.AddWithValue("@RiskCurrencyVar", dr["Risk Currency"].ToString());
                            cmd.Parameters.AddWithValue("@OpenPriceVar", dr["Open Price"]);
                            cmd.Parameters.AddWithValue("@ClosePriceVar", dr["Close Price"]);
                            cmd.Parameters.AddWithValue("@VolumeVar", dr["Volume"]);
                            cmd.Parameters.AddWithValue("@LastPriceVar", dr["Last Price"]);
                            cmd.Parameters.AddWithValue("@AskPriceVar", dr["Ask Price"]);
                            cmd.Parameters.AddWithValue("@BidPriceVar", dr["Bid Price"]);
                            cmd.Parameters.AddWithValue("@PERatioVar", dr["PE Ratio"]);


                            cmd.Parameters.AddWithValue("@DividendDeclareDateVar", dr["Dividend Declared Date"]);
                            cmd.Parameters.AddWithValue("@DividendExDateVar", dr["Dividend Ex Date"]);
                            cmd.Parameters.AddWithValue("@DividendRecordDateVar", dr["Dividend Record Date"]);
                            cmd.Parameters.AddWithValue("@DividendPayDateVar", dr["Dividend Pay Date"]);
                            cmd.Parameters.AddWithValue("@DividendAmountVar", dr["Dividend Amount"]);
                            cmd.Parameters.AddWithValue("@FrequencyVar", dr["Frequency"].ToString());
                            cmd.Parameters.AddWithValue("@DividendTypeVar", dr["Dividend Type"].ToString());

                            con.Open();
                            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
                            dataAdapter.Fill(table);
                            dataAdapter.Dispose();
                        }

                    }
                    return "Insertion Successful";
                }
                else if (ext.ToLower() == ".csv")   //if csv file
                {

                    string[] Lines = File.ReadAllLines(filename);
                    string[] Fields;
                    Lines = Lines.Skip(1).ToArray();
                    //read each line and exclude header
                    foreach (var line in Lines)
                    {
                        //read each column value for a record and store in fields
                        Fields = line.Split(new char[] { ',' });

                        using (SqlConnection con = new SqlConnection(connectionString))
                        {
                            //Execute Reader    
                            DataTable table = new DataTable();
                            string command = "Equity.usp_InsertIntoDB";
                            SqlCommand cmd = new SqlCommand(command, con);
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.AddWithValue("@SecurityNameVar", Fields[0].Replace("\"", ""));//to remove " " 
                            cmd.Parameters.AddWithValue("@SecurityDescriptionVar", Fields[1].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@HasPositionVar", Fields[2].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ActiveSecurityVar", Fields[3].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@LotSizeVar", Fields[4].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGUniqueNameVar", Fields[5].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@CUSIPVar", Fields[6].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ISINVar", Fields[7].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@SEDOLVar", Fields[8].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGtickerVar", Fields[9].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGUIDVar", Fields[10].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGGlobalIDVar", Fields[11].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGTickerandExchangeVar", Fields[12].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@ADRFlagVar", Fields[13].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ADRUnderlyingTickerVar", Fields[14].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ADRUnderlyingCurrencyVar", Fields[15].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@SharesPerADRVar", Fields[16].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IPODateVar", Fields[17].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@PricingCurrencyVar", Fields[18].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@SettleDaysVar", Fields[19].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@OutstandingSharesVar", Fields[20].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@VotingRightsPerShareVar", Fields[21].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@AverageVolume20DVar", Fields[22].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BetaVar", Fields[23].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ShortInterestVar", Fields[24].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ReturnYTDVar", Fields[25].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Volatility90DayVar", Fields[26].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@AssetClassVar", Fields[27].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCountryVar", Fields[28].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCreditRatingVar", Fields[29].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCurrencyVar", Fields[30].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFInstrumentVar", Fields[31].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFLiquidityProfileVar", Fields[32].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFMaturityVar", Fields[33].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFNAISCCodeVar", Fields[34].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@PFRegionVar", Fields[35].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFSectorVar", Fields[36].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFSubAssetClassVar", Fields[37].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CountryOfIssuanceVar", Fields[38].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ExchangeVar", Fields[39].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IssuerNameVar", Fields[40].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IssueCurrencyVar", Fields[41].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@TradingCurrencyVar", Fields[42].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGIndustrySubGroupVar", Fields[43].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGIndustryGroupVar", Fields[44].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@BBGSectorVar", Fields[45].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CountryOfIncorporationVar", Fields[46].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@RiskCurrencyVar", Fields[47].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@OpenPriceVar", Fields[48].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ClosePriceVar", Fields[49].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@VolumeVar", Fields[50].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@LastPriceVar", Fields[51].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@AskPriceVar", Fields[52].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BidPriceVar", Fields[53].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PERatioVar", Fields[54].Replace("\"", ""));

                            cmd.Parameters.AddWithValue("@DividendDeclareDateVar", Fields[55].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@DividendExDateVar", Fields[56].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@DividendRecordDateVar", Fields[57].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@DividendPayDateVar", Fields[58].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@DividendAmountVar", Fields[59].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@FrequencyVar", Fields[60].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@DividendTypeVar", Fields[61].Replace("\"", ""));

                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(table);
                            da.Dispose();
                        }

                    }
                    return "Insertion successful.";
                }
                else { return "Can upload only CSV/Excel file"; }
            }
            catch (Exception e)
            {
                return $"Insertion Failed due to {e.Message} ";
            }

        }
    }
}
