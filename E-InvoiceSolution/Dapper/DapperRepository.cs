using Dapper;
using E_InvoiceSolution.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace E_InvoiceSolution.Dapper
{

    public static class DapperRepository
    {
        private static string ConnectionString = ConfigurationManager.ConnectionStrings["HP_Connection"].ConnectionString;

        /// <summary>
        /// Get List of Distributors to fill the dropdown
        /// </summary>
        /// <returns></returns>
        public static async Task<List<DistributorsDropDownListModel>> GetDistributorsList()
        {
            try
            {
                using (var connection = new SqlConnection(ConnectionString))
                {
                    var User = connection.Query<DistributorsDropDownListModel>(
                    "Select Distributorid,Distributorname from tblDistributor order by distributorname",
                    commandType: CommandType.Text);
                    return User.ToList();
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Get List of Invoices to update the PO
        /// </summary>
        /// <returns></returns>
        public static async Task<List<OutStandingPoModel>> GetOutstandingPOs(int DistributorID)
        {
            try
            {
                using (var connection = new SqlConnection(ConnectionString))
                {
                    var User = connection.Query<OutStandingPoModel>(
                    "select POID, POCODE, orderdate, po_comments from tblpurchaseorder " +
                    "where invoicenumber is null and distributorid = @DistributorId",
                    new
                    {
                        DistributorId = DistributorID
                    },
                    commandType: CommandType.Text);
                    return User.ToList();
                }
            }
            catch (Exception ex)

            {
                throw;
            }
        }

        public static async Task<KeheUploadResultModel> UploadExcelDataIntoNaturesBestPOTable(int DistributorID, int POID, string Path)
        {
            try
            {
                KeheUploadResultModel keheUploadResultModel = new KeheUploadResultModel();
                DataTable excelData = new DataTable();

                if (DistributorID == 39)
                {
                    //Read the excel file into datatable
                    excelData = ReadExcelFile(Path);
                    // Add SrNo into datatable 
                    excelData.Columns.Add("SNO", typeof(int));
                }
                if (DistributorID == 51)
                {
                    excelData = GenerateMcKeesonDataSetFromExcel(Path);
                }

                for (int count = 0; count < excelData.Rows.Count; count++)
                {
                    excelData.Rows[count]["SNO"] = count + 1;
                }

                using (var connection = new SqlConnection(ConnectionString))
                {
                    var data = connection.QueryMultiple
                        ("ImportNaturesBestPOFromExcel",
                        new
                        {
                            @POID = POID,
                            @ExcelData = excelData.AsTableValuedParameter()
                        },
                        commandType: CommandType.StoredProcedure
                    );

                    keheUploadResultModel.POID = POID;
                    //Get number of deleted records
                    keheUploadResultModel.RowsDeletedFromDailyPO = data.Read<int>().FirstOrDefault();
                    //Select Total Number of Rows From Excel
                    keheUploadResultModel.RowsInExcelSheet = data.Read<int>().FirstOrDefault();
                    //GET Duplicate SKUs as comma separated string
                    keheUploadResultModel.RepatedSKUsFoundInExcelSheet = data.Read<string>().FirstOrDefault();
                    //Delete the records from another supporting table with the give POID                    
                    keheUploadResultModel.RowsDeltedFromNaturesBestSupportingTable = data.Read<int>().FirstOrDefault(); ;
                    //retrieve the total products and quantity for the given POID from the purchase order details table and displays them
                    var PoDetails = data.Read<dynamic>().FirstOrDefault();
                    keheUploadResultModel.SkusFoundIntblPurchaseOrderDetailsForPOID = PoDetails.ProductsSum;
                    keheUploadResultModel.QuantityFoundIntblPurchaseOrderDetailsForPOID = PoDetails.QuantitySum;
                    // Retrieves the total products and qunatities in the given invoice file
                    var InvoiceDetails = data.Read<dynamic>().FirstOrDefault();
                    keheUploadResultModel.TotalSkusReceived = PoDetails.TotalProducts;
                    keheUploadResultModel.TotalQuantityReceived = PoDetails.TotalQty;
                    //Retrieves the purchase order level data for the purchase orders
                    keheUploadResultModel.PurchaseOrderDetails = data.Read<PurchaseOrderDetails>().ToList();
                    //Retrieves the items which are shipped but not ordered and displays them
                    keheUploadResultModel.ShippedButNotOrdered = data.Read<ItemDetails>().ToList();
                    //Retrieves the products which are extra shipped 
                    keheUploadResultModel.ExtraShipped = data.Read<ItemDetails>().ToList();
                    //retrieves the proudcts which are not shipped (ordered but not received)
                    keheUploadResultModel.OrderedButNotReceived = data.Read<ItemDetails>().ToList();

                }

                return keheUploadResultModel;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static DataTable ReadExcelFile(string Path)
        {
            try
            {

                string excelConnectString = $@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={Path};Extended Properties=""Excel 8.0;HDR=YES;""";

                if (Path.EndsWith(".xlsx"))
                {
                    excelConnectString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Path};Extended Properties=""Excel 12.0;HDR=YES;""";
                }

                OleDbConnection objConn = new OleDbConnection(excelConnectString);
                objConn.Open();

                List<string> sheets = new List<string>();
                // Get First Sheet Name
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dt.Rows)
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        string s = drSheet["TABLE_NAME"].ToString();
                        sheets.Add(s.StartsWith("'") ? s.Substring(1, s.Length - 3) : s.Substring(0, s.Length - 1));
                    }

                OleDbCommand objCmd = new OleDbCommand("Select * From [" + sheets.First() + "$]", objConn);

                OleDbDataAdapter objDatAdap = new OleDbDataAdapter();
                objDatAdap.SelectCommand = objCmd;
                DataSet ds = new DataSet();
                objDatAdap.Fill(ds);
                objConn.Close();
                return ds.Tables[0];
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static DataTable GenerateMcKeesonDataSetFromExcel(string path)
        {
            DataTable dt = CreateBlankDataSet();
            int i = 0;
            DataTable dtFromExcel = ReadExcelFile(path);
            foreach (DataRow dataRow in dtFromExcel.Rows)
            {
                i = i++;
                string InvoiceNumber = dataRow["Invoice/Transaction Number"]?.ToString();
                string PONumber = dataRow["PO#"]?.ToString();
                string SKU = dataRow["Item Number"]?.ToString();
                string ItemQuantity = dataRow["Item Quantity"]?.ToString();
                string UnitPrice = dataRow["Item Price Unit"]?.ToString();
                string ItemExtendedPrice = dataRow["Item Extended Price"]?.ToString();
                string InvoiceDate = dataRow["Invoice Date"]?.ToString();
                string Description = dataRow["Item Description"]?.ToString();
                string UPC = dataRow["NDC/UPC value"]?.ToString();


                dt.Rows.Add(i.ToString(),
                    0,
                    Convert.ToInt32(ItemQuantity),
                    SKU,
                    "",
                    "",
                    Description,
                    UPC,
                    Convert.ToDouble(UnitPrice),
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    InvoiceNumber,
                    Convert.ToDateTime(InvoiceDate),
                    0,
                    0);
            }

            return dt;
        }

        private static DataTable CreateBlankDataSet()
        {
            DataTable dt = new DataTable("Table");
            dt.Columns.Add(new DataColumn("Line", typeof(string)));
            dt.Columns.Add(new DataColumn("Order Qty", typeof(float)));

            dt.Columns.Add(new DataColumn("Ship Qty", typeof(float)));
            dt.Columns.Add(new DataColumn("Ship Item", typeof(string)));

            dt.Columns.Add(new DataColumn("Pack Size", typeof(string)));
            dt.Columns.Add(new DataColumn("BRAND", typeof(string)));

            dt.Columns.Add(new DataColumn("Description", typeof(string)));
            dt.Columns.Add(new DataColumn("UPC Code", typeof(string)));

            dt.Columns.Add(new DataColumn("Retail", typeof(float)));
            dt.Columns.Add(new DataColumn("Suggested Retail", typeof(float)));

            dt.Columns.Add(new DataColumn("Wholesale", typeof(float)));
            dt.Columns.Add(new DataColumn("Adj Wholesale", typeof(float)));

            dt.Columns.Add(new DataColumn("Discount %", typeof(float)));
            dt.Columns.Add(new DataColumn("Discount $", typeof(float)));


            dt.Columns.Add(new DataColumn("Net Each", typeof(float)));
            dt.Columns.Add(new DataColumn("Net Billable", typeof(float)));


            dt.Columns.Add(new DataColumn("UpCharges", typeof(float)));
            dt.Columns.Add(new DataColumn("BottleTax", typeof(float)));


            dt.Columns.Add(new DataColumn("Invoice No", typeof(float)));
            dt.Columns.Add(new DataColumn("InvoiceDate", typeof(DateTime)));

            dt.Columns.Add(new DataColumn("Store No", typeof(int)));
            dt.Columns.Add(new DataColumn("SNO", typeof(int)));

            return dt;
        }
    }
}