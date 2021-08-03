using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;

namespace XSyncExcelUploader
{
    public partial class UploadExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // BindGridview();
            }
        }

        //function to show all orders in Gridview
        private void BindGridview()
        {
            string CS = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
            using (SqlConnection con = new SqlConnection(CS))
            {
                SqlCommand cmd = new SqlCommand("spGetAllOrders", con);
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();
                GridView1.DataSource = cmd.ExecuteReader();
                GridView1.DataBind();
                con.Close();
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile != null)
            {
                try
                {                    
                    string excelPath = string.Concat(Server.MapPath("~/FilesUpload/") + Path.GetFileName(FileUpload1.PostedFile.FileName));
                    FileUpload1.SaveAs(excelPath);
                    string conString = string.Empty;
                    string extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    switch (extension)
                    {
                        case ".xls": //Excel 97-03
                            conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                            break;
                        case ".xlsx": //Excel 07 or higher
                            conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                            break;

                    }
                    conString = string.Format(conString, excelPath);                    
                    using (OleDbConnection con = new OleDbConnection(conString))
                    {
                        con.Open();
                        string sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                        DataTable dtExcelData = new DataTable();
                        dtExcelData.Columns.AddRange(new DataColumn[30] {                                                     
                                                     new DataColumn("restaurant_name", typeof(string)),
                                                     new DataColumn("AreaZone", typeof(string)),
                                                     new DataColumn("City", typeof(string)),
                                                     new DataColumn("invoice_no", typeof(string)),
                                                     new DataColumn("date", typeof(DateTime)),
                                                     new DataColumn("payment_type", typeof(string)),
                                                     new DataColumn("order_type", typeof(string)),
                                                     new DataColumn("orderstatus", typeof(string)),
                                                     new DataColumn("channelname", typeof(string)),
                                                     new DataColumn("assign_to", typeof(string)),
                                                     new DataColumn("customer_phone", typeof(string)),
                                                     new DataColumn("customer_name", typeof(string)),
                                                     new DataColumn("customer_address", typeof(string)),
                                                     new DataColumn("persons", typeof(string)),
                                                     new DataColumn("order_cancel_reason", typeof(string)),
                                                     new DataColumn("my_amount", typeof(Decimal)),
                                                     new DataColumn("total_tax", typeof(Decimal)),
                                                     new DataColumn("discount", typeof(Decimal)),
                                                     new DataColumn("delivery_charge", typeof(Decimal)),
                                                     new DataColumn("container_charge", typeof(Decimal)),
                                                     new DataColumn("service_charge", typeof(Decimal)),
                                                     new DataColumn("waived_off", typeof(Decimal)),
                                                     new DataColumn("round_off", typeof(Decimal)),
                                                     new DataColumn("total", typeof(Decimal)),
                                                     new DataColumn("item_name", typeof(String)),
                                                     new DataColumn("category_name", typeof(String)),
                                                     new DataColumn("sap_code", typeof(int)),
                                                     new DataColumn("item_price", typeof(Decimal)),
                                                     new DataColumn("item_quantity", typeof(int)),
                                                     new DataColumn("item_total", typeof(decimal))

                        });

                        using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [" + sheet1 + "]", con))
                        {
                            oda.Fill(dtExcelData);
                        }
                        con.Close();

                        // break down restaurant_name column in Excel file to different columns in sql
                        foreach (DataRow row in dtExcelData.Rows)
                        {
                            string fullresname = row["restaurant_name"].ToString();
                            string[] divideresname = fullresname.Split('-');
                            row["restaurant_name"] = divideresname[0].Trim();
                            row["AreaZone"] = divideresname[1].Trim();
                            row["City"] = divideresname[2].Trim();
                        }
                        
                        // begin upload to sql
                        string consString = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;
                        using (SqlConnection sqlcon = new SqlConnection(consString))
                        {
                            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlcon))
                            {
                                //Set the database table name
                                sqlBulkCopy.DestinationTableName = "dbo.OrderDetails";

                                //Mapping the Excel columns with the sql table OrderDetails
                                sqlBulkCopy.ColumnMappings.Add("restaurant_name", "restaurant_name");
                                sqlBulkCopy.ColumnMappings.Add("AreaZone", "AreaZone");
                                sqlBulkCopy.ColumnMappings.Add("City", "City");
                                sqlBulkCopy.ColumnMappings.Add("invoice_no", "invoice_no");
                                sqlBulkCopy.ColumnMappings.Add("date", "date");
                                sqlBulkCopy.ColumnMappings.Add("payment_type", "payment_type");
                                sqlBulkCopy.ColumnMappings.Add("order_type", "order_type");
                                sqlBulkCopy.ColumnMappings.Add("status", "orderstatus");
                                sqlBulkCopy.ColumnMappings.Add("area", "channelname");
                                sqlBulkCopy.ColumnMappings.Add("assign_to", "assign_to");
                                sqlBulkCopy.ColumnMappings.Add("customer_phone", "customer_phone");
                                sqlBulkCopy.ColumnMappings.Add("customer_name", "customer_name");
                                sqlBulkCopy.ColumnMappings.Add("customer_address", "customer_address");
                                sqlBulkCopy.ColumnMappings.Add("persons", "persons");
                                sqlBulkCopy.ColumnMappings.Add("order_cancel_reason", "order_cancel_reason");
                                sqlBulkCopy.ColumnMappings.Add("my_amount", "my_amount");
                                sqlBulkCopy.ColumnMappings.Add("total_tax", "total_tax");
                                sqlBulkCopy.ColumnMappings.Add("discount", "discount");
                                sqlBulkCopy.ColumnMappings.Add("delivery_charge", "delivery_charge");
                                sqlBulkCopy.ColumnMappings.Add("container_charge", "container_charge");
                                sqlBulkCopy.ColumnMappings.Add("service_charge", "service_charge");
                                sqlBulkCopy.ColumnMappings.Add("waived_off", "waived_off");
                                sqlBulkCopy.ColumnMappings.Add("round_off", "round_off");
                                sqlBulkCopy.ColumnMappings.Add("total", "total");
                                sqlBulkCopy.ColumnMappings.Add("item_name", "item_name");
                                sqlBulkCopy.ColumnMappings.Add("category_name", "category_name");
                                sqlBulkCopy.ColumnMappings.Add("sap_code", "sap_code");
                                sqlBulkCopy.ColumnMappings.Add("item_price", "item_price");
                                sqlBulkCopy.ColumnMappings.Add("item_quantity", "item_quantity");
                                sqlBulkCopy.ColumnMappings.Add("item_total", "item_total");
                                sqlcon.Open();
                                sqlBulkCopy.WriteToServer(dtExcelData);
                                sqlcon.Close();
                            }
                        }
                        lblMessage.Text = "File uploaded successfully!";
                        lblMessage.ForeColor = System.Drawing.Color.Green;
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
    }
}