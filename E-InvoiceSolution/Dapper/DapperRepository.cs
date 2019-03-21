using ClosedXML.Excel;
using Dapper;
using E_InvoiceSolution.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
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
                if (DistributorID == 56)
                {
                    excelData = GenerateEuropaDataSetFromExcel(Path);
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
                    keheUploadResultModel.TotalSkusReceived = InvoiceDetails.TotalProducts;
                    keheUploadResultModel.TotalQuantityReceived = InvoiceDetails.TotalQty;
                    //Retrieves the purchase order level data for the purchase orders
                    keheUploadResultModel.PurchaseOrderDetails = data.Read<PurchaseOrderDetails>().ToList();
                    //Retrieves the items which are shipped but not ordered and displays them
                    keheUploadResultModel.ShippedButNotOrdered = data.Read<ItemDetails>().ToList();
                    //Retrieves the products which are extra shipped 
                    keheUploadResultModel.ExtraShipped = data.Read<ItemDetails>().ToList();
                    //retrieves the proudcts which are not shipped (ordered but not received)
                    keheUploadResultModel.OrderedButNotReceived = data.Read<ItemDetails>().ToList();
                    // Get the Credit report data
                    keheUploadResultModel.CreditReports = data.Read<CreditReport>().ToList();

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

                if (Path.ToString().ToLower().EndsWith(".xlsx"))
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

        private static DataTable ReadExcelFileWithClosedXML(string path)
        {

            //Create a new DataTable.
            DataTable dt = new DataTable();

            try
            {
                //Open the Excel file using ClosedXML.
                using (XLWorkbook workBook = new XLWorkbook(path))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    //Loop through the Worksheet rows.
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        //Use the first row to add columns to DataTable.
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString().Trim());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            //Add rows to DataTable.
                            dt.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static DataTable GenerateMcKeesonDataSetFromExcel(string path)
        {
            DataTable dt = CreateBlankDataSet();
            int i = 0;

            DataTable dtFromExcel;
            if (path.EndsWith(".xls"))
            {
                dtFromExcel = ReadExcelFile(path);
            }
            if (path.EndsWith(".csv"))
            {
                dtFromExcel = ReadCsvFile(path);
            }
            else
            {
                dtFromExcel = ReadExcelFileWithClosedXML(path);
            }

            foreach (DataRow dataRow in dtFromExcel.Rows)
            {
                i = i++;
                string InvoiceNumber = dataRow["InvoiceNumber"]?.ToString().Trim();
                //string PONumber = dataRow["PO#"]?.ToString().Trim();
                string SKU = dataRow["OrderItemNumber"]?.ToString().Trim();
                string ItemQuantity = dataRow["FilledQuantity"]?.ToString().Trim();
                string UnitPrice = dataRow["FilledUnitPrice"]?.ToString().Trim();
                //string ItemExtendedPrice = dataRow["Item Extended Price"]?.ToString().Trim();
                //string InvoiceDate = dataRow["InvoiceDate"]?.ToString().Trim();
                string Description = dataRow["SellDescription"]?.ToString().Trim();
                //string UPC = dataRow["NDC/UPC value"]?.ToString().Trim();
                string Wholesale = dataRow["ProperContractPrice"]?.ToString().Trim();
                string MSRP = dataRow["RetailPrice"]?.ToString().Trim();

                dt.Rows.Add(
                        //Line
                        i.ToString(),
                        //Order Qty
                        null,
                        //Shipped Qty
                        Convert.ToInt32(ItemQuantity),
                        //Shipped Item
                        SKU,
                        //Pack Size
                        "",
                        //Brand
                        "",
                        //Description
                        Description,
                        // UPC Code
                        "",
                        //Retail
                        Convert.ToDouble(UnitPrice),
                        // Suggested Retail
                        Convert.ToDouble(MSRP),
                        //Wholesale
                        Convert.ToDouble(Wholesale),
                        //Adj Wholesale
                        Convert.ToDouble(UnitPrice),
                        //Discount %
                        null,
                        //Discount $
                        null,
                        //Net Each
                        null,
                        //Net Billable
                        null,
                        //UpCharges
                        null,
                        //BottleTax
                        null,
                        //Invoice No
                        Convert.ToDouble(InvoiceNumber),
                        //InvoiceDate
                        null, //Convert.ToDateTime(InvoiceDate),
                              //Store No
                        null,
                        //SNO
                        null);
            }

            return dt;
        }

        private static DataTable GenerateEuropaDataSetFromExcel(string path)
        {
            DataTable dt = CreateBlankDataSet();
            int i = 0;
            DataTable dtFromExcel = ReadExcelFileWithClosedXML(path);
            List<string> duplicateSKUList = new List<string>();

            foreach (DataRow dataRow in dtFromExcel.Rows)
            {
                string SKU = dataRow["Item Number"]?.ToString().Trim();
                if (!duplicateSKUList.Contains(SKU))
                {

                    i = i++;
                    DateTime InvoiceDate;
                    if (!string.IsNullOrEmpty(dataRow["Order Date"]?.ToString().Trim())
                        && DateTime.TryParse(dataRow["Order Date"]?.ToString().Trim(),
                        out InvoiceDate))
                    {
                        string InvoiceNumber = dataRow["Invoice Number"]?.ToString().Trim();
                        string PONumber = dataRow["PO Number"]?.ToString().Trim();

                        string ItemQuantity = dataRow["QTY"]?.ToString().Trim();
                        string UnitPrice = dataRow["Your Price"]?.ToString().Trim().Replace("$", "");
                        string ItemExtendedPrice = dataRow["Ext Price"]?.ToString().Trim().Replace("$", "");
                        //string InvoiceDate = dataRow["Order Date"]?.ToString();
                        string Description = dataRow["Item Description"]?.ToString().Trim();
                        string WholeSale = dataRow["Wholesale"]?.ToString().Trim().Replace("$", "");
                        string DiscPercent = dataRow["Disc %"]?.ToString().Trim();

                        // Check if sheet has duplicate SKUs
                        IEnumerable<DataRow> duplicateSKUs = from tbl in dtFromExcel.AsEnumerable()
                                                             where tbl.Field<string>(3) == SKU
                                                             select tbl;

                        int Qty = Convert.ToInt32(ItemQuantity);
                        double UP = Convert.ToDouble(UnitPrice);
                        if (duplicateSKUs.Count() > 1)
                        {
                            for (int l = 1; l < duplicateSKUs.Count(); l++)
                            {
                                duplicateSKUList.Add(duplicateSKUs.ToList()[l]["Item Number"]?.ToString().Trim());
                                Qty += Convert.ToInt32(duplicateSKUs.ToList()[l]["QTY"]);
                            }                            
                            UP = UP / Qty;
                            duplicateSKUs = null;
                        }

                        dt.Rows.Add(
                            //Line
                            i.ToString(),
                            //Order Qty
                            null,
                            //Shipped Qty
                            Convert.ToInt32(ItemQuantity),
                            //Shipped Item
                            SKU,
                            //Pack Size
                            "",
                            //Brand
                            "",
                            //Description
                            Description,
                            // UPC Code
                            "",
                            //Retail
                            Convert.ToDouble(UP),
                            // Suggested Retail
                            null,
                            //Wholesale
                            Convert.ToDouble(WholeSale),
                            //Adj Wholesale - Received Unit Price
                            Convert.ToDouble(UP),
                            //Discount %
                            Convert.ToDouble(DiscPercent),
                            //Discount $
                            null,
                            //Net Each
                            null,
                            //Net Billable
                            null,
                            //UpCharges
                            null,
                            //BottleTax
                            null,
                            //Invoice No
                            Convert.ToDouble(InvoiceNumber),
                            //InvoiceDate
                            Convert.ToDateTime(InvoiceDate),
                            //Store No
                            null,
                            //SNO
                            null);
                    }

                }
            }
            return dt;
        }

        public static DataTable ReadCsvFile(string path)
        {

            DataTable dtCsv = new DataTable();
            string Fulltext;
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString(); //read full file text  
                    string[] rows = Fulltext.Split('\n'); //split full file text into rows  
                    for (int i = 0; i < rows.Count() - 1; i++)
                    {
                        string[] rowValues = rows[i].Split(','); //split each row with comma to get individual values  
                        {
                            if (i == 0)
                            {
                                //add headers  
                                for (int j = 0; j < rowValues.Count(); j++)
                                {
                                    if (dtCsv.Columns.Contains(rowValues[j]))
                                    {
                                        dtCsv.Columns.Add(rowValues[j] + j);
                                    }
                                    else { dtCsv.Columns.Add(rowValues[j]); }
                                }
                            }
                            else
                            {
                                DataRow dr = dtCsv.NewRow();
                                for (int k = 0; k < rowValues.Count(); k++)
                                {
                                    dr[k] = rowValues[k].ToString();
                                }
                                dtCsv.Rows.Add(dr); //add other rows  
                            }
                        }
                    }
                }
            }

            return dtCsv;
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