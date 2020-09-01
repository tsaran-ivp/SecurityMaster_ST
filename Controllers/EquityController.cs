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
                    cmd.Parameters.AddWithValue("@PFNAICSCodeVar", equity.PFNAICSCode);

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


                            cmd.Parameters.AddWithValue("@SecurityNameVar", dr[0]);
                            cmd.Parameters.AddWithValue("@SecurityDescriptionVar", dr[1]);
                            cmd.Parameters.AddWithValue("@HasPositionVar", dr[2].ToString());
                            cmd.Parameters.AddWithValue("@ActiveSecurityVar", dr[3].ToString());
                            cmd.Parameters.AddWithValue("@LotSizeVar",dr[4]);
                            cmd.Parameters.AddWithValue("@BBGUniqueNameVar", dr[5].ToString());

                            cmd.Parameters.AddWithValue("@CUSIPVar", dr[6].ToString());
                            cmd.Parameters.AddWithValue("@ISINVar", dr[7].ToString());
                            cmd.Parameters.AddWithValue("@SEDOLVar", dr[8].ToString());
                            cmd.Parameters.AddWithValue("@BBGtickerVar", dr[9].ToString());
                            cmd.Parameters.AddWithValue("@BBGUIDVar", dr[10].ToString());
                            cmd.Parameters.AddWithValue("@BBGGlobalIDVar", dr[11].ToString());
                            cmd.Parameters.AddWithValue("@BBGTickerandExchangeVar", dr[12].ToString());

                            cmd.Parameters.AddWithValue("@ADRFlagVar", dr[13].ToString());
                            cmd.Parameters.AddWithValue("@ADRUnderlyingTickerVar", dr[14].ToString());
                            cmd.Parameters.AddWithValue("@ADRUnderlyingCurrencyVar", dr[15].ToString());
                            cmd.Parameters.AddWithValue("@SharesPerADRVar", dr[16]);
                            cmd.Parameters.AddWithValue("@IPODateVar", dr[17]);

                            cmd.Parameters.AddWithValue("@PricingCurrencyVar", dr[18].ToString());
                            cmd.Parameters.AddWithValue("@SettleDaysVar", dr[19]);
                            cmd.Parameters.AddWithValue("@OutstandingSharesVar", dr[20]);
                            cmd.Parameters.AddWithValue("@VotingRightsPerShareVar",dr[21]);

                            cmd.Parameters.AddWithValue("@AverageVolume20DVar", dr[22]);
                            cmd.Parameters.AddWithValue("@BetaVar", dr[23]);
                            cmd.Parameters.AddWithValue("@ShortInterestVar", dr[24]);
                            cmd.Parameters.AddWithValue("@ReturnYTDVar", dr[25]);
                            cmd.Parameters.AddWithValue("@Volatility90DayVar", dr[26]);


                            cmd.Parameters.AddWithValue("@AssetClassVar", dr[27].ToString());
                            cmd.Parameters.AddWithValue("@PFCountryVar", dr[28].ToString());
                            cmd.Parameters.AddWithValue("@PFCreditRatingVar", dr[29].ToString());
                            cmd.Parameters.AddWithValue("@PFCurrencyVar", dr[30].ToString());
                            cmd.Parameters.AddWithValue("@PFInstrumentVar", dr[31].ToString());
                            cmd.Parameters.AddWithValue("@PFLiquidityProfileVar", dr[32].ToString());
                            cmd.Parameters.AddWithValue("@PFMaturityVar", dr[33]);
                            cmd.Parameters.AddWithValue("@PFNAICSCodeVar", dr[34].ToString());

                            cmd.Parameters.AddWithValue("@PFRegionVar", dr[35].ToString());
                            cmd.Parameters.AddWithValue("@PFSectorVar", dr[36].ToString());
                            cmd.Parameters.AddWithValue("@PFSubAssetClassVar", dr[37].ToString());

                            cmd.Parameters.AddWithValue("@CountryOfIssuanceVar", dr[28].ToString());
                            cmd.Parameters.AddWithValue("@ExchangeVar", dr[39].ToString());
                            cmd.Parameters.AddWithValue("@IssuerNameVar", dr[40].ToString());
                            cmd.Parameters.AddWithValue("@IssueCurrencyVar", dr[41].ToString());
                            cmd.Parameters.AddWithValue("@TradingCurrencyVar", dr[42].ToString());
                            cmd.Parameters.AddWithValue("@BBGIndustrySubGroupVar", dr[43].ToString());
                            cmd.Parameters.AddWithValue("@BBGIndustryGroupVar", dr[44].ToString());

                            cmd.Parameters.AddWithValue("@BBGSectorVar", dr[45].ToString());
                            cmd.Parameters.AddWithValue("@CountryOfIncorporationVar", dr[46].ToString());
                            cmd.Parameters.AddWithValue("@RiskCurrencyVar", dr[47].ToString());
                            cmd.Parameters.AddWithValue("@OpenPriceVar", dr[48]);
                            cmd.Parameters.AddWithValue("@ClosePriceVar", dr[49]);
                            cmd.Parameters.AddWithValue("@VolumeVar", dr[50]);
                            cmd.Parameters.AddWithValue("@LastPriceVar", dr[51]);
                            cmd.Parameters.AddWithValue("@AskPriceVar",dr[52]);
                            cmd.Parameters.AddWithValue("@BidPriceVar", dr[53]);
                            cmd.Parameters.AddWithValue("@PERatioVar", dr[54]);


                            cmd.Parameters.AddWithValue("@DividendDeclareDateVar",dr[55]);
                            cmd.Parameters.AddWithValue("@DividendExDateVar", dr[56]);
                            cmd.Parameters.AddWithValue("@DividendRecordDateVar", dr[57]);
                            cmd.Parameters.AddWithValue("@DividendPayDateVar", dr[58]);
                            cmd.Parameters.AddWithValue("@DividendAmountVar", dr[59]);
                            cmd.Parameters.AddWithValue("@FrequencyVar", dr[60]);
                            cmd.Parameters.AddWithValue("@DividendTypeVar", dr[61]);

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
