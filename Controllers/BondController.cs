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
    public class BondController : ApiController
    {
        string connectionString = "data source=192.168.0.104\\sql_express1,63862; database=SecurityMaster_ST; user=Sa; password=valley@1234"; //Hardcoded Connection String

        public HttpResponseMessage Get()
        {

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                //Execute Reader    
                DataTable table = new DataTable();
                string command = "SelectBonds";
                SqlCommand cmd = new SqlCommand(command, con);
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(table);

                return Request.CreateResponse(HttpStatusCode.OK, table);

            } 

        }

        public string Put(Bond bond)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    
                    //Execute Reader    
                    DataTable table = new DataTable();
                    string command = "UpdateBonds";
                    SqlCommand cmd = new SqlCommand(command, con);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@Id", bond.ID);
                    cmd.Parameters.AddWithValue("@SecurityDescription", bond.SecurityDescription);
                    cmd.Parameters.AddWithValue("@SecurityName", bond.SecurityName);
                    cmd.Parameters.AddWithValue("@AssetType", bond.AssetType);
                    cmd.Parameters.AddWithValue("@InvestmentType", bond.InvestmentType);
                    cmd.Parameters.AddWithValue("@TradingFactor", bond.TradingFactor);
                    cmd.Parameters.AddWithValue("@PricingFactor", bond.PricingFactor);
                    cmd.Parameters.AddWithValue("@ISIN", bond.ISIN);
                    cmd.Parameters.AddWithValue("@BBGTicker", bond.BBGTicker);
                    cmd.Parameters.AddWithValue("@BBGUniqueID", bond.BBGUniqueID);
                    cmd.Parameters.AddWithValue("@CUSIP", bond.CUSIP);
                    cmd.Parameters.AddWithValue("@SEDOL", bond.SEDOL);
                    cmd.Parameters.AddWithValue("@FirstCouponDate", bond.FirstCouponDate);
                    cmd.Parameters.AddWithValue("@Cap", bond.Cap);
                    cmd.Parameters.AddWithValue("@FloorValue", bond.FloorValue);
                    cmd.Parameters.AddWithValue("@CouponFrequency", bond.CouponFrequency);
                    cmd.Parameters.AddWithValue("@Coupon", bond.Coupon);
                    cmd.Parameters.AddWithValue("@CouponType", bond.CouponType);
                    cmd.Parameters.AddWithValue("@Spread", bond.Spread);
                    cmd.Parameters.AddWithValue("@CallableFlag", bond.CallableFlag);
                    cmd.Parameters.AddWithValue("@FixtoFloatFlag", bond.FixtoFloatFlag);
                    cmd.Parameters.AddWithValue("@PutableFlag", bond.PutableFlag);
                    cmd.Parameters.AddWithValue("@IssueDate", bond.IssueDate);
                    cmd.Parameters.AddWithValue("@LastResetDate", bond.LastResetDate);
                    cmd.Parameters.AddWithValue("@Maturity", bond.Maturity);
                    cmd.Parameters.AddWithValue("@CallNotificationMaxDays", bond.CallNotificationMaxDays);
                    cmd.Parameters.AddWithValue("@PutNotificationMaxDays", bond.PutNotificationMaxDays);
                    cmd.Parameters.AddWithValue("@PenultimateCouponDate", bond.PenultimateCouponDate);
                    cmd.Parameters.AddWithValue("@ResetFrequency", bond.ResetFrequency);
                    cmd.Parameters.AddWithValue("@HasPosition", bond.HasPosition);
                    cmd.Parameters.AddWithValue("@MacaulayDuration", bond.MacaulayDuration);
                    cmd.Parameters.AddWithValue("@Volatility_30D", bond.Volatility_30D);
                    cmd.Parameters.AddWithValue("@Volatility_90D", bond.Volatility_90D);
                    cmd.Parameters.AddWithValue("@Convexity", bond.Convexity);
                    cmd.Parameters.AddWithValue("@AverageVolume_30Day", bond.AverageVolume_30Day);
                    cmd.Parameters.AddWithValue("@PFAssetClass", bond.PFAssetClass);
                    cmd.Parameters.AddWithValue("@PFCountry", bond.PFCountry);
                    cmd.Parameters.AddWithValue("@PFCreditRating", bond.PFCreditRating);
                    cmd.Parameters.AddWithValue("@PFCurrency", bond.PFCurrency);
                    cmd.Parameters.AddWithValue("@PFInstrument", bond.PFInstrument);
                    cmd.Parameters.AddWithValue("@PFLiquidityProfile", bond.PFLiquidityProfile);
                    cmd.Parameters.AddWithValue("@PFMaturity", bond.PFMaturity);
                    cmd.Parameters.AddWithValue("@PFNAICSCode", bond.PFNAICSCode);
                    cmd.Parameters.AddWithValue("@PFRegion", bond.PFRegion);
                    cmd.Parameters.AddWithValue("@PFSector", bond.PFSector);
                    cmd.Parameters.AddWithValue("@PFSubAssetClass", bond.PFSubAssetClass);
                    cmd.Parameters.AddWithValue("@BloombergIndustryGroup", bond.BloombergIndustryGroup);
                    cmd.Parameters.AddWithValue("@BloombergIndustrySubGroup", bond.BloombergIndustrySubGroup);
                    cmd.Parameters.AddWithValue("@BloombergIndustrySector", bond.BloombergIndustrySector);
                    cmd.Parameters.AddWithValue("@CountryofIssuance", bond.CountryofIssuance);
                    cmd.Parameters.AddWithValue("@IssueCurrency", bond.IssueCurrency);
                    cmd.Parameters.AddWithValue("@IssuerName", bond.IssuerName);
                    cmd.Parameters.AddWithValue("@RiskCurrency", bond.RiskCurrency);
                    cmd.Parameters.AddWithValue("@PutDate", bond.PutDate);
                    cmd.Parameters.AddWithValue("@PutPrice", bond.PutPrice);
                    cmd.Parameters.AddWithValue("@AskPrice", bond.AskPrice);
                    cmd.Parameters.AddWithValue("@HighPrice", bond.HighPrice);
                    cmd.Parameters.AddWithValue("@LowPrice", bond.LowPrice);
                    cmd.Parameters.AddWithValue("@OpenPrice", bond.OpenPrice);
                    cmd.Parameters.AddWithValue("@Volume", bond.Volume);
                    cmd.Parameters.AddWithValue("@BidPrice", bond.BidPrice);
                    cmd.Parameters.AddWithValue("@LastPrice", bond.LastPrice);
                    cmd.Parameters.AddWithValue("@CallDate", bond.CallDate);
                    cmd.Parameters.AddWithValue("@CallPrice", bond.CallPrice);

                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(table);

                    return "put success";

                }

            }
            catch (Exception)
            {

                return "could not put";
            }

        }

        public string Delete(int id)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    //Execute Reader    
                    DataTable table = new DataTable();
                    string command = "DeleteBonds";
                    SqlCommand cmd = new SqlCommand(command, con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Id", id);
                    con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(table);
                    return "delete success";

                }

            }
            catch (Exception)
            {

                return "could not delete";
            }
        }
        public string Post() {

            try
            {
                //fetch and save file
                var httRequest = HttpContext.Current.Request;
                var postedFile = httRequest.Files[0];
                string filename = postedFile.FileName;
                var physicalpath = HttpContext.Current.Server.MapPath("~/BondFiles/" + filename);
                postedFile.SaveAs(physicalpath);
                 string ext = Path.GetExtension(postedFile.FileName);
                
                //check file extension if excel or csv
                if (ext.ToLower() == ".xls" || ext.ToLower() == ".xlsx") 
                {
                    Stream stream = postedFile.InputStream;
                    IExcelDataReader reader = null;
                
                //check which type of excel file
                    if (ext.ToLower() == ".xls") {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (ext.ToLower() == ".xlsx") {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                    DataSet ds = reader.AsDataSet();
                    reader.Close();
                    ds.Tables[0].Rows.RemoveAt(0);
                    //for each new record call the procedure and insert data

                foreach (DataRow dr in ds.Tables[0].Rows)
                {

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    //Execute Reader    
                    DataTable table = new DataTable();
                    string command = "InsertBonds";
                    SqlCommand cmd = new SqlCommand(command, con);
                    cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@SecurityDescription", dr[0].ToString());
                        cmd.Parameters.AddWithValue("@SecurityName", dr[1].ToString());
                        cmd.Parameters.AddWithValue("@AssetType", dr[2].ToString());
                        cmd.Parameters.AddWithValue("@InvestmentType", dr[3].ToString());
                        cmd.Parameters.AddWithValue("@TradingFactor", dr[4].ToString());
                        cmd.Parameters.AddWithValue("@PricingFactor", dr[5].ToString());
                        cmd.Parameters.AddWithValue("@ISIN", dr[6].ToString());
                        cmd.Parameters.AddWithValue("@BBGTicker", dr[7].ToString());
                        cmd.Parameters.AddWithValue("@BBGUniqueID", dr[8].ToString());
                        cmd.Parameters.AddWithValue("@CUSIP", dr[9].ToString());
                        cmd.Parameters.AddWithValue("@SEDOL", dr[10].ToString());
                        cmd.Parameters.AddWithValue("@FirstCouponDate", dr[11].ToString());
                        cmd.Parameters.AddWithValue("@Cap", dr[12].ToString());
                        cmd.Parameters.AddWithValue("@FloorValue", dr[13].ToString());
                        cmd.Parameters.AddWithValue("@CouponFrequency", dr[14].ToString());
                        cmd.Parameters.AddWithValue("@Coupon", dr[15].ToString());
                        cmd.Parameters.AddWithValue("@CouponType", dr[16].ToString());
                        cmd.Parameters.AddWithValue("@Spread", dr[17].ToString());
                        cmd.Parameters.AddWithValue("@CallableFlag", dr[18].ToString());
                        cmd.Parameters.AddWithValue("@FixtoFloatFlag", dr[19].ToString());
                        cmd.Parameters.AddWithValue("@PutableFlag", dr[20].ToString());
                        cmd.Parameters.AddWithValue("@IssueDate", dr[21].ToString());
                        cmd.Parameters.AddWithValue("@LastResetDate", dr[22].ToString());
                        cmd.Parameters.AddWithValue("@Maturity", dr[23].ToString());
                        cmd.Parameters.AddWithValue("@CallNotificationMaxDays", dr[24].ToString());
                        cmd.Parameters.AddWithValue("@PutNotificationMaxDays", dr[25].ToString());
                        cmd.Parameters.AddWithValue("@PenultimateCouponDate", dr[26].ToString());
                        cmd.Parameters.AddWithValue("@ResetFrequency", dr[27].ToString());
                        cmd.Parameters.AddWithValue("@HasPosition", dr[28].ToString());
                        cmd.Parameters.AddWithValue("@MacaulayDuration", dr[29].ToString());
                        cmd.Parameters.AddWithValue("@Volatility_30D", dr[30].ToString());
                        cmd.Parameters.AddWithValue("@Volatility_90D", dr[31].ToString());
                        cmd.Parameters.AddWithValue("@Convexity", dr[32].ToString());
                        cmd.Parameters.AddWithValue("@AverageVolume_30Day", dr[33].ToString());
                        cmd.Parameters.AddWithValue("@PFAssetClass", dr[34].ToString());
                        cmd.Parameters.AddWithValue("@PFCountry", dr[35].ToString());
                        cmd.Parameters.AddWithValue("@PFCreditRating", dr[36].ToString());
                        cmd.Parameters.AddWithValue("@PFCurrency", dr[37].ToString());
                        cmd.Parameters.AddWithValue("@PFInstrument", dr[38].ToString());
                        cmd.Parameters.AddWithValue("@PFLiquidityProfile", dr[39].ToString());
                        cmd.Parameters.AddWithValue("@PFMaturity", dr[40].ToString());
                        cmd.Parameters.AddWithValue("@PFNAICSCode", dr[41].ToString());
                        cmd.Parameters.AddWithValue("@PFRegion", dr[42].ToString());
                        cmd.Parameters.AddWithValue("@PFSector", dr[43].ToString());
                        cmd.Parameters.AddWithValue("@PFSubAssetClass", dr[44].ToString());
                        cmd.Parameters.AddWithValue("@BloombergIndustryGroup", dr[45].ToString());
                        cmd.Parameters.AddWithValue("@BloombergIndustrySubGroup", dr[46].ToString());
                        cmd.Parameters.AddWithValue("@BloombergIndustrySector", dr[47].ToString());
                        cmd.Parameters.AddWithValue("@CountryofIssuance", dr[48].ToString());
                        cmd.Parameters.AddWithValue("@IssueCurrency", dr[49].ToString());
                        cmd.Parameters.AddWithValue("@IssuerName", dr[50].ToString());
                        cmd.Parameters.AddWithValue("@RiskCurrency", dr[51].ToString());
                        cmd.Parameters.AddWithValue("@PutDate", dr[52].ToString());
                        cmd.Parameters.AddWithValue("@PutPrice", dr[53].ToString());
                        cmd.Parameters.AddWithValue("@AskPrice", dr[54].ToString());
                        cmd.Parameters.AddWithValue("@HighPrice", dr[55].ToString());
                        cmd.Parameters.AddWithValue("@LowPrice", dr[56].ToString());
                        cmd.Parameters.AddWithValue("@OpenPrice", dr[57].ToString());
                        cmd.Parameters.AddWithValue("@Volume", dr[58].ToString());
                        cmd.Parameters.AddWithValue("@BidPrice", dr[59].ToString());
                        cmd.Parameters.AddWithValue("@LastPrice", dr[60].ToString());
                        cmd.Parameters.AddWithValue("@CallDate", dr[61].ToString());
                        cmd.Parameters.AddWithValue("@CallPrice", dr[62].ToString());

                        con.Open();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(table);
                        da.Dispose();                        
                 }

                }
                    return "insert success";

                }
                //if csv file
                else if (ext.ToLower() == ".csv") {

                    string[] Lines = File.ReadAllLines(filename);
                    string[] Fields;
                    Lines = Lines.Skip(1).ToArray();
                    //read each line and exclude header
                    foreach (var line in Lines)
                    {
                        //read each column value for a record and store in fields
                        Fields = line.Split(new char[] {','});

                        using (SqlConnection con = new SqlConnection(connectionString))
                        {
                            //Execute Reader    
                            DataTable table = new DataTable();
                            string command = "InsertBonds";
                            SqlCommand cmd = new SqlCommand(command, con);
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.AddWithValue("@SecurityDescription", Fields[0].Replace("\"",""));//to remove " " 
                            cmd.Parameters.AddWithValue("@SecurityName", Fields[1].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@AssetType", Fields[2].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@InvestmentType", Fields[3].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@TradingFactor", Fields[4].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PricingFactor", Fields[5].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ISIN", Fields[6].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGTicker", Fields[7].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BBGUniqueID", Fields[8].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CUSIP", Fields[9].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@SEDOL", Fields[10].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@FirstCouponDate", Fields[11].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Cap", Fields[12].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@FloorValue", Fields[13].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CouponFrequency", Fields[14].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Coupon", Fields[15].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CouponType", Fields[16].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Spread", Fields[17].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CallableFlag", Fields[18].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@FixtoFloatFlag", Fields[19].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PutableFlag", Fields[20].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IssueDate", Fields[21].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@LastResetDate", Fields[22].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Maturity", Fields[23].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CallNotificationMaxDays", Fields[24].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PutNotificationMaxDays", Fields[25].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PenultimateCouponDate", Fields[26].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@ResetFrequency", Fields[27].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@HasPosition", Fields[28].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@MacaulayDuration", Fields[29].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Volatility_30D", Fields[30].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Volatility_90D", Fields[31].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Convexity", Fields[32].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@AverageVolume_30Day", Fields[33].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFAssetClass", Fields[34].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCountry", Fields[35].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCreditRating", Fields[36].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFCurrency", Fields[37].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFInstrument", Fields[38].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFLiquidityProfile", Fields[39].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFMaturity", Fields[40].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFNAICSCode", Fields[41].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFRegion", Fields[42].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFSector", Fields[43].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PFSubAssetClass", Fields[44].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BloombergIndustryGroup", Fields[45].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BloombergIndustrySubGroup", Fields[46].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BloombergIndustrySector", Fields[47].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CountryofIssuance", Fields[48].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IssueCurrency", Fields[49].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@IssuerName", Fields[50].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@RiskCurrency", Fields[51].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PutDate", Fields[52].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@PutPrice", Fields[53].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@AskPrice", Fields[54].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@HighPrice", Fields[55].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@LowPrice", Fields[56].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@OpenPrice", Fields[57].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@Volume", Fields[58].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@BidPrice", Fields[59].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@LastPrice", Fields[60].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CallDate", Fields[61].Replace("\"", ""));
                            cmd.Parameters.AddWithValue("@CallPrice", Fields[62].Replace("\"", ""));

                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(table);
                            da.Dispose();
                        }


                    }


                    return "insert success";

                } else { return "enter only csv/excel file"; }
             
            }
            catch (Exception e)
            {

                return e.Message;
            }

        }
    }
}
