using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace XSyncExcelUploader
{
    public partial class UploadExcel : System.Web.UI.Page
    {
        string excelPath = string.Empty;
        string conString = string.Empty;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // nothing
            }
        }        

        protected void btnUpload_Click(object sender, EventArgs e)
        {            
            if (FileUpload1.PostedFile != null)
            {
                try
                {                                                            
                    string extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    switch (extension)
                    {
                        case ".xls": //Excel 97-03
                            conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                            UploadExcelFile();
                            break;
                        case ".xlsx": //Excel 07 or higher
                            conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                            UploadExcelFile();
                            break;
                        case ".csv": //csv file
                            conString = ConfigurationManager.ConnectionStrings["CSVConString"].ConnectionString;
                            UploadCSVFile();
                            break;
                    }                   
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    lblMessage.Text = "File not uploaded.Please try again.";
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                }
            }            
            
        }


        protected void UploadCSVFile()
        {           
            excelPath = string.Concat(Server.MapPath("~/FilesUpload/"));
            FileUpload1.SaveAs(excelPath + Path.GetFileName(FileUpload1.PostedFile.FileName));
            conString = string.Format(conString, excelPath);
            using (OleDbConnection con = new OleDbConnection(conString))
            {
                con.Open();
                string sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();
                dtExcelData.Columns.AddRange(new DataColumn[31] {
                                                     new DataColumn("SysPK", typeof(Int64)),
                                                     new DataColumn("Res Name", typeof(string)),
                                                     new DataColumn("AreaZone", typeof(string)),
                                                     new DataColumn("City", typeof(string)),
                                                     new DataColumn("Date", typeof(DateTime)),
                                                     new DataColumn("Time", typeof(string)),
                                                     new DataColumn("Hour", typeof(int)),
                                                     new DataColumn("Invoice ID", typeof(string)),
                                                     new DataColumn("Online Order Number", typeof(Int64)),
                                                     new DataColumn("Payment Type", typeof(string)),
                                                     new DataColumn("Order Status", typeof(string)),
                                                     new DataColumn("Area", typeof(string)),
                                                     new DataColumn("Order Type", typeof(string)),
                                                     new DataColumn("Cancel Reason", typeof(string)),
                                                     new DataColumn("SapCode", typeof(int)),
                                                     new DataColumn("Category", typeof(String)),
                                                     new DataColumn("Item Name", typeof(String)),
                                                     new DataColumn("AddOn", typeof(String)),
                                                     new DataColumn("Variation", typeof(String)),
                                                     new DataColumn("Round Off", typeof(Decimal)),
                                                     new DataColumn("Item Qty", typeof(int)),
                                                     new DataColumn("Item Price", typeof(Decimal)),
                                                     new DataColumn("AddOn Qty.", typeof(int)),
                                                     new DataColumn("AddOn Price", typeof(int)),
                                                     new DataColumn("AddOn Sub Total", typeof(int)),
                                                     new DataColumn("Sub Total", typeof(int)),
                                                     new DataColumn("Discount", typeof(Decimal)),
                                                     new DataColumn("Delivery Charge", typeof(Decimal)),
                                                     new DataColumn("Packing Charge", typeof(Decimal)),
                                                     new DataColumn("Taxes", typeof(Decimal)),
                                                     new DataColumn("Final Total", typeof(Decimal))                                                    
                                             });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", con))
                {
                    oda.Fill(dtExcelData);
                    progresslbl.Text = "Data loaded";
                }
                con.Close();

                // break down restaurant_name column, merge date time into one in Excel file to different columns in sql
                foreach (DataRow row in dtExcelData.Rows)
                {
                    string fullresname = row["Res Name"].ToString();
                    System.Diagnostics.Debug.WriteLine(fullresname + "--" + row["SysPK"].ToString());
                    string[] divideresname = fullresname.Split('-');
                    row["Res Name"] = divideresname[0].Trim();
                    row["AreaZone"] = divideresname[1].Trim();
                    row["City"] = divideresname[2].Trim();

                    var timeval = row["Time"].ToString().Split(' ');
                    row["Time"] = timeval[1] +" " + timeval[2];

                }

                // begin upload to sql
                string consString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                using (SqlConnection sqlcon = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.OrderDetailsProd";

                        //Mapping the Excel columns with the sql table OrderDetails
                        //sqlBulkCopy.ColumnMappings.Add("SysPK", "SysPK");
                        sqlBulkCopy.ColumnMappings.Add("ï»¿SysPK", "SysPK");
                        sqlBulkCopy.ColumnMappings.Add("Res Name", "restaurant_name");
                        sqlBulkCopy.ColumnMappings.Add("AreaZone", "AreaZone");
                        sqlBulkCopy.ColumnMappings.Add("City", "City");
                        sqlBulkCopy.ColumnMappings.Add("Date", "date");
                        sqlBulkCopy.ColumnMappings.Add("Hour", "hourvalue");
                        sqlBulkCopy.ColumnMappings.Add("Time", "timevalue");
                        sqlBulkCopy.ColumnMappings.Add("Invoice ID", "invoice_no");
                        sqlBulkCopy.ColumnMappings.Add("Online Order Number", "onlineorder_no");
                        sqlBulkCopy.ColumnMappings.Add("Payment Type", "payment_type");
                        sqlBulkCopy.ColumnMappings.Add("Order Status", "orderstatus");
                        sqlBulkCopy.ColumnMappings.Add("Area", "channelname");
                        sqlBulkCopy.ColumnMappings.Add("Order Type", "order_type");
                        sqlBulkCopy.ColumnMappings.Add("Cancel Reason", "order_cancel_reason");
                        sqlBulkCopy.ColumnMappings.Add("SapCode", "sap_code");
                        sqlBulkCopy.ColumnMappings.Add("Category", "category_name");
                        sqlBulkCopy.ColumnMappings.Add("Item Name", "item_name");
                        sqlBulkCopy.ColumnMappings.Add("AddOn", "addon");
                        sqlBulkCopy.ColumnMappings.Add("Variation", "variation");
                        sqlBulkCopy.ColumnMappings.Add("Round Off", "round_off");
                        sqlBulkCopy.ColumnMappings.Add("Item Qty", "item_quantity");
                        sqlBulkCopy.ColumnMappings.Add("Item Price", "item_price");
                       // sqlBulkCopy.ColumnMappings.Add("AddOn Qty.", "addon_qty");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Qty#", "addon_qty");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Price", "addon_price");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Sub Total", "addon_subtotal");
                        sqlBulkCopy.ColumnMappings.Add("Sub Total", "subtotal");
                        sqlBulkCopy.ColumnMappings.Add("Discount", "discount");
                        sqlBulkCopy.ColumnMappings.Add("Delivery Charge", "delivery_charge");
                        sqlBulkCopy.ColumnMappings.Add("Packing Charge", "packing_charge");
                        sqlBulkCopy.ColumnMappings.Add("Taxes", "total_tax");
                        sqlBulkCopy.ColumnMappings.Add("Final Total", "final_total");

                        sqlcon.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        sqlcon.Close();
                    }
                }
                lblMessage.Text = "File uploaded successfully!";
                lblMessage.ForeColor = System.Drawing.Color.Green;                                
            }


        }

        protected void UploadExcelFile()
        {
            excelPath = string.Concat(Server.MapPath("~/FilesUpload/") + Path.GetFileName(FileUpload1.PostedFile.FileName));
            FileUpload1.SaveAs(excelPath);
            conString = string.Format(conString, excelPath);
            using (OleDbConnection con = new OleDbConnection(conString))
            {
                con.Open();
                string sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtExcelData = new DataTable();
                dtExcelData.Columns.AddRange(new DataColumn[31] {
                                                     new DataColumn("SysPK", typeof(Int64)),
                                                     new DataColumn("Res Name", typeof(string)),
                                                     new DataColumn("AreaZone", typeof(string)),
                                                     new DataColumn("City", typeof(string)),
                                                     new DataColumn("Date", typeof(DateTime)),
                                                     new DataColumn("Time", typeof(string)),
                                                     new DataColumn("Hour", typeof(int)),
                                                     new DataColumn("Invoice ID", typeof(string)),
                                                     new DataColumn("Online Order Number", typeof(Int64)),
                                                     new DataColumn("Payment Type", typeof(string)),
                                                     new DataColumn("Order Status", typeof(string)),
                                                     new DataColumn("Area", typeof(string)),
                                                     new DataColumn("Order Type", typeof(string)),
                                                     new DataColumn("Cancel Reason", typeof(string)),
                                                     new DataColumn("SapCode", typeof(int)),
                                                     new DataColumn("Category", typeof(String)),
                                                     new DataColumn("Item Name", typeof(String)),
                                                     new DataColumn("AddOn", typeof(String)),
                                                     new DataColumn("Variation", typeof(String)),
                                                     new DataColumn("Round Off", typeof(Decimal)),
                                                     new DataColumn("Item Qty", typeof(int)),
                                                     new DataColumn("Item Price", typeof(Decimal)),
                                                     new DataColumn("AddOn Qty.", typeof(int)),
                                                     new DataColumn("AddOn Price", typeof(int)),
                                                     new DataColumn("AddOn Sub Total", typeof(int)),
                                                     new DataColumn("Sub Total", typeof(int)),
                                                     new DataColumn("Discount", typeof(Decimal)),
                                                     new DataColumn("Delivery Charge", typeof(Decimal)),
                                                     new DataColumn("Packing Charge", typeof(Decimal)),
                                                     new DataColumn("Taxes", typeof(Decimal)),
                                                     new DataColumn("Final Total", typeof(Decimal))

                                                     //not required as per latest excel
                                                    /* new DataColumn("assign_to", typeof(string)),
                                                     new DataColumn("customer_phone", typeof(string)),
                                                     new DataColumn("customer_name", typeof(string)),
                                                     new DataColumn("customer_address", typeof(string)),
                                                     new DataColumn("persons", typeof(string)),                                                     
                                                     new DataColumn("my_amount", typeof(Decimal)),
                                                     new DataColumn("container_charge", typeof(Decimal)),
                                                     new DataColumn("service_charge", typeof(Decimal)),
                                                     new DataColumn("waived_off", typeof(Decimal)),                  
                                                     new DataColumn("item_total", typeof(decimal))*/

                                             });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", con))
                {
                    oda.Fill(dtExcelData);
                }
                con.Close();

                // break down restaurant_name column, merge date time into one in Excel file to different columns in sql
                foreach (DataRow row in dtExcelData.Rows)
                {
                    string fullresname = row["Res Name"].ToString();
                    System.Diagnostics.Debug.WriteLine(fullresname + "--" + row["SysPK"].ToString());
                    string[] divideresname = fullresname.Split('-');
                    row["Res Name"] = divideresname[0].Trim();
                    row["AreaZone"] = divideresname[1].Trim();
                    row["City"] = divideresname[2].Trim();

                    //merge date time

                    // var date = row["Date"];
                    // var time = row["Time"];
                    // DateTime timevalue = Convert.ToDateTime(row["Time"].ToString());
                    // var newdate = date + " " + time;
                    // row["Time"] = timevalue;
                }

                // begin upload to sql
                string consString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                using (SqlConnection sqlcon = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon))
                    {
                        //Set the database table name
                        sqlBulkCopy.DestinationTableName = "dbo.OrderDetailsProd";

                        //Mapping the Excel columns with the sql table OrderDetails
                        sqlBulkCopy.ColumnMappings.Add("SysPK", "SysPK");
                        sqlBulkCopy.ColumnMappings.Add("Res Name", "restaurant_name");
                        sqlBulkCopy.ColumnMappings.Add("AreaZone", "AreaZone");
                        sqlBulkCopy.ColumnMappings.Add("City", "City");
                        sqlBulkCopy.ColumnMappings.Add("Date", "date");
                        sqlBulkCopy.ColumnMappings.Add("Hour", "hourvalue");
                        sqlBulkCopy.ColumnMappings.Add("Time", "timevalue");
                        sqlBulkCopy.ColumnMappings.Add("Invoice ID", "invoice_no");
                        sqlBulkCopy.ColumnMappings.Add("Online Order Number", "onlineorder_no");
                        sqlBulkCopy.ColumnMappings.Add("Payment Type", "payment_type");
                        sqlBulkCopy.ColumnMappings.Add("Order Status", "orderstatus");
                        sqlBulkCopy.ColumnMappings.Add("Area", "channelname");
                        sqlBulkCopy.ColumnMappings.Add("Order Type", "order_type");
                        sqlBulkCopy.ColumnMappings.Add("Cancel Reason", "order_cancel_reason");
                        sqlBulkCopy.ColumnMappings.Add("SapCode", "sap_code");
                        sqlBulkCopy.ColumnMappings.Add("Category", "category_name");
                        sqlBulkCopy.ColumnMappings.Add("Item Name", "item_name");
                        sqlBulkCopy.ColumnMappings.Add("AddOn", "addon");
                        sqlBulkCopy.ColumnMappings.Add("Variation", "variation");
                        sqlBulkCopy.ColumnMappings.Add("Round Off", "round_off");
                        sqlBulkCopy.ColumnMappings.Add("Item Qty", "item_quantity");
                        sqlBulkCopy.ColumnMappings.Add("Item Price", "item_price");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Qty.", "addon_qty");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Price", "addon_price");
                        sqlBulkCopy.ColumnMappings.Add("AddOn Sub Total", "addon_subtotal");
                        sqlBulkCopy.ColumnMappings.Add("Sub Total", "subtotal");
                        sqlBulkCopy.ColumnMappings.Add("Discount", "discount");
                        sqlBulkCopy.ColumnMappings.Add("Delivery Charge", "delivery_charge");
                        sqlBulkCopy.ColumnMappings.Add("Packing Charge", "packing_charge");
                        sqlBulkCopy.ColumnMappings.Add("Taxes", "total_tax");
                        sqlBulkCopy.ColumnMappings.Add("Final Total", "final_total");

                        // not required as per latest excel
                        /*  sqlBulkCopy.ColumnMappings.Add("assign_to", "assign_to");
                          sqlBulkCopy.ColumnMappings.Add("customer_phone", "customer_phone");
                          sqlBulkCopy.ColumnMappings.Add("customer_name", "customer_name");
                          sqlBulkCopy.ColumnMappings.Add("customer_address", "customer_address");
                          sqlBulkCopy.ColumnMappings.Add("persons", "persons");                                
                          sqlBulkCopy.ColumnMappings.Add("my_amount", "my_amount");           
                          sqlBulkCopy.ColumnMappings.Add("container_charge", "container_charge");
                          sqlBulkCopy.ColumnMappings.Add("service_charge", "service_charge");
                          sqlBulkCopy.ColumnMappings.Add("waived_off", "waived_off");                                
                          sqlBulkCopy.ColumnMappings.Add("total", "total");
                          sqlBulkCopy.ColumnMappings.Add("item_total", "item_total");*/

                        sqlcon.Open();
                        sqlBulkCopy.WriteToServer(dtExcelData);
                        sqlcon.Close();
                    }
                }
                lblMessage.Text = "File uploaded successfully!";
                lblMessage.ForeColor = System.Drawing.Color.Green;
            }
        }      
    }
}