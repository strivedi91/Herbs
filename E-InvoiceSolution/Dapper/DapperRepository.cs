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
                            @POID = 1,
                            @ExcelData = excelData.AsTableValuedParameter()
                        },
                        commandType: CommandType.StoredProcedure
                    );
                }

                //Fill the result object
               

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
                return ds.Tables[0];
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}