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

        public static async Task<KeheUploadResultModel> UploadExcelDataIntoNaturesBestPOTable(int POID, string Path)
        {
            try
            {
                KeheUploadResultModel keheUploadResultModel = new KeheUploadResultModel();

                //Read the excel file into datatable
                DataTable excelData = ReadExcelFile(Path);
                // Add SrNo into datatable 
                excelData.Columns.Add("SNO", typeof(int));
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
                    keheUploadResultModel.RowsDeltedFromNaturesBestSupportingTable = data.Read<int>().FirstOrDefault();
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
                //string excelConnectString = @"Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " + excelFileName + ";" + "Extended Properties = Excel 8.0; HDR=Yes;IMEX=1";

                OleDbConnection objConn = new OleDbConnection(excelConnectString);
                OleDbCommand objCmd = new OleDbCommand("Select * From [Sheet1$]", objConn);

                OleDbDataAdapter objDatAdap = new OleDbDataAdapter();
                objDatAdap.SelectCommand = objCmd;
                DataSet ds = new DataSet();
                objDatAdap.Fill(ds);
                return RemoveEmptyRowsFromDataTable(ds.Tables[0]);
            }
            catch (Exception)
            {
                throw;
            }
        }

        static DataTable RemoveEmptyRowsFromDataTable(DataTable dt)
        {
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                if (dt.Rows[i][1] == DBNull.Value)
                    dt.Rows[i].Delete();
            }
            dt.AcceptChanges();
            return dt;
        }
    }
}