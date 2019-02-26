using E_InvoiceSolution.Dapper;
using E_InvoiceSolution.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace E_InvoiceSolution.Controllers
{
    public class HomeController : Controller
    {
        public async Task<ActionResult> Index()
        {
            ViewBag.DistributorsList = new SelectList(await DapperRepository.GetDistributorsList(), "Distributorid", "Distributorname");
            return View();
        }

        public async Task<JsonResult> GetOutstandingPOsByDistributorId(int DistributorId)
        {
            var pOs = await DapperRepository.GetOutstandingPOs(DistributorId);
            return Json(pOs);
        }

        private string UploadExcelFileToFolder(int DistributorID)
        {
            string fname = string.Empty;

            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];


                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            fname = file.FileName;
                        }
                        // Get the complete folder path and store the file inside it.  
                        fname = Path.Combine(GetUploadDirectoryPath(DistributorID), fname);
                        file.SaveAs(fname);
                    }
                }
                catch (Exception ex)
                {
                    return "";
                }
                return fname;
            }
            else
            {
                return "";
            }
        }

        private async Task<ActionResult> UpdatePO(int DistributorID, int POID, string filePath)
        {

            // Read the excel and store it into the database
            var keheUploadResult = await DapperRepository.UploadExcelDataIntoNaturesBestPOTable(DistributorID,
                POID, filePath);

            var html = ExtensionMethods.RenderViewToString(this.ControllerContext, "UploadResult", keheUploadResult);
            System.IO.File.WriteAllText(GetCreditReportDirectoryPath() + POID + "_log.xls", html);

            //Set Download path
            keheUploadResult.DownloadPath = GetCreditReportDirectoryPath() + POID + "_log.xls";
            keheUploadResult.DownloadFileName = POID + "_log.xls";
            return PartialView("UploadResult", keheUploadResult);

        }

        [HttpPost]
        public async Task<ActionResult> UploadFiles()
        {
            int POID = Int32.Parse(Request["POID"]);
            int DistributorID = Int32.Parse(Request["DistributorID"]);

            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    var fname = UploadExcelFileToFolder(DistributorID);

                   return await UpdatePO(DistributorID, POID, fname);

                    // Returns message that successfully uploaded  
                    return Json("File Uploaded Successfully!");
                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }
        }

        [HttpGet]
        public virtual ActionResult Download(string file)
        {
            string fullPath = Path.Combine(GetCreditReportDirectoryPath(), file);
            return File(fullPath, "application/vnd.ms-excel", file);
        }

        private string GetUploadDirectoryPath(int DistributorID)
        {
            DateTime today = DateTime.Now;
            string CurrentMonth = String.Format("{0:MMMM}", today);
            string Year = DateTime.Now.Year.ToString();
            string path = string.Empty;
            if (DistributorID == 39)
            {
                path = Server.MapPath("~/Uploads/NaturesBestPOFiles/" + Year + "/" + CurrentMonth + "/");
            }
            else if (DistributorID == 51)
            {
                path = Server.MapPath("~/Uploads/McKeesonPOFiles/" + Year + "/" + CurrentMonth + "/");
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }

        private string GetCreditReportDirectoryPath()
        {
            DateTime today = DateTime.Now;
            string CurrentMonth = String.Format("{0:MMMM}", today);
            string Year = DateTime.Now.Year.ToString();
            string path = Server.MapPath("~/CreditReports/NaturesBestPOFiles/" + Year + "/" + CurrentMonth + "/");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }
    }
}