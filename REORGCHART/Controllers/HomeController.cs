using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using System.IO;
using System.ComponentModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Web;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using System.Web.Script.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Dynamic;
using Newtonsoft.Json.Converters;
using System.Reflection;

using REORGCHART.Models;
using REORGCHART.Helper;
using REORGCHART.Data;
using OfficeOpenXml;

namespace REORGCHART.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        Models.DBContext db = new Models.DBContext();
        ApplicationUser UserData = System.Web.HttpContext.Current.GetOwinContext().GetUserManager<ApplicationUserManager>().FindById(System.Web.HttpContext.Current.User.Identity.GetUserId());

        Common comClass = new Common();

        public ActionResult Index()
        {
            MyLastAction myla = GetUserCurrentAction();

            var viewModel = new MyModel
            {
                UseDate = DateTime.Now,
                KeyDate = myla.KeyDate,
                ShowLevel = myla.ShowLevel,
                Levels = myla.Levels,
                Version = myla.Version,
                Oper = myla.Oper,
                View = myla.View,
                Country = myla.Country,
                ChartData = GetOrgChartData(myla.Role, myla.Country, myla.ShowLevel, myla.Levels, myla.Oper, myla.Version),
                Role = myla.Role
            };

            return View(viewModel);
        }

        private MyLastAction GetUserCurrentAction()
        {
            var UCA = (from uca in db.UserLastActions
                       where uca.UserId == UserData.UserName
                       select uca).FirstOrDefault();

            MyLastAction viewModel = new MyLastAction();
            if (UCA == null)
            {
                UserLastActions uca = new UserLastActions();
                uca.Company = UserData.CompanyName;
                uca.KeyDate = "2018/07/11";
                uca.UserId = UserData.UserName;
                uca.Version = "1";
                uca.UsedView = "Normal";
                uca.Oper = "OV";
                uca.Levels = "One";
                uca.ShowLevel = "101212";
                uca.Country = "";
                uca.Role = "User";
                db.UserLastActions.Add(uca);

                db.SaveChanges();

                viewModel.ShowLevel = "101212";
                viewModel.KeyDate = "2018/07/11";
                viewModel.Levels = "One";
                viewModel.Version = "1";
                viewModel.Oper = "OV";
                viewModel.View = "Normal";
                viewModel.Country = "";
                viewModel.Role = "User";
            }
            else
            {
                viewModel.KeyDate = UCA.KeyDate;
                viewModel.ShowLevel = UCA.ShowLevel;
                viewModel.Levels = UCA.Levels;
                viewModel.Version = UCA.Version;
                viewModel.Oper = UCA.Oper;
                viewModel.View = UCA.UsedView;
                viewModel.Country = UCA.Country;
                viewModel.Role = UCA.Role;
            }

            return viewModel;
        }

        [HttpPost]
        public string SetSelectedValues(string KeyDate, string UsedView, string Country, string ShowLevel, string Levels,
                                        string Oper, string Version, string Role)
        {
            var UCA = (from uca in db.UserLastActions
                       where uca.UserId == UserData.UserName
                       select uca).FirstOrDefault();
            if (UCA != null)
            {
                UCA.KeyDate = KeyDate;
                UCA.UsedView = UsedView;
                UCA.Country = Country;
                UCA.ShowLevel = ShowLevel;
                UCA.Levels = Levels;
                UCA.Oper = Oper;
                UCA.Version = Version;
                UCA.Role = Role;

                db.SaveChanges();
            }

            return "Sucess";
        }

        [HttpPost]
        public string GetOrgChartData(string UserType, string Country, string ShowLevel, string Levels, string Oper, string Version)
        {
            try
            {
                DataTable orgChartData = null;

                if (ShowLevel == null) ShowLevel = "";
                if (Levels == null || Levels == "" || Levels == " ") Levels = "1";

                if (Oper == "LV")
                {
                    // Legal chart details 
                    LegalInfo LI = new LegalInfo();
                    orgChartData = LI.GetLegalInfo(UserData.UserName, UserData.CompanyName, Country, ShowLevel, Levels, Version);
                }
                else if (Oper == "OV")
                {
                    // Operation chart details              
                    LevelInfo LI = new LevelInfo();
                    orgChartData = LI.GetLevelInfo(UserData.UserName, UserData.CompanyName, UserType, ShowLevel, Levels, Version);
                }

                return JsonConvert.SerializeObject(orgChartData);
            }
            catch (Exception ex)
            {
                Dictionary<string, string> Info = new Dictionary<string, string>();
                Info.Add("WebHTML", "Error");
                Info.Add("Message", ex.Message);
                Info.Add("StackTrace", ex.StackTrace);

                return JsonConvert.SerializeObject(Info);
            }
        }

        private DataTable ConvertToDatatable<T>(List<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    table.Columns.Add(prop.Name, prop.PropertyType.GetGenericArguments()[0]);
                else
                    table.Columns.Add(prop.Name, prop.PropertyType);
            }

            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }

        private void InsertDataToDb(DataTable dtTable)
        {
            // Create the SqlBulkCopy object. 
            // Note that the column positions in the source DataTable 
            // match the column positions in the destination table so 
            // there is no need to map columns. 
            string ConnectionString = ConfigurationManager.ConnectionStrings["ORG_CHART"].ConnectionString;
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(ConnectionString))
            {
                bulkCopy.DestinationTableName = "dbo.LEVEL_INFO";

                try
                {
                    // Write from the source to the destination.
                    bulkCopy.WriteToServer(dtTable);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-excel", fileName);
            }
            else
            {
                // Problem - Log the error, generate a blank file,
                //           redirect to another controller action - whatever fits with your application
                return new EmptyResult();
            }
        }

        public ActionResult VersionControl(string VersionControl, string Oper, string Version, string ChartData)
        {
            if (VersionControl == "DownloadVersion")
            {
                var UploadExcelFile = (from uef in db.UploadFilesHeaders
                                       where uef.CompanyName == UserData.CompanyName && uef.UserId == UserData.UserName
                                       select uef).OrderByDescending(x => x.CreatedDate).FirstOrDefault();

                string ServerMapPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/Upload/" + UploadExcelFile.JSONFileName);
                var filePath = new FileInfo(ServerMapPath);
                using (ExcelPackage workbook = new ExcelPackage(filePath))
                {
                    // Generate a new unique identifier against which the file can be stored
                    string handle = Guid.NewGuid().ToString();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        memoryStream.Position = 0;
                        TempData[handle] = memoryStream.ToArray();
                    }

                    // Note we are returning a filename as well as the handle
                    return new JsonResult()
                    {
                        Data = new { FileGuid = handle, FileName = "TestReportOutput.xlsx" }
                    };
                }
            }
            else if (VersionControl == "SaveVersion")
            {
                var UploadExcelFile = (from uef in db.UploadFilesHeaders
                                       where uef.CompanyName == UserData.CompanyName
                                       select uef).OrderByDescending(x => x.CreatedDate).FirstOrDefault();

                string ServerMapPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/Download/" + UploadExcelFile.JSONFileName);
                using (StreamReader reader = new StreamReader(ServerMapPath))
                {
                    string jsonData = reader.ReadToEnd();
                    dynamic lstDynamic = JsonConvert.DeserializeObject(jsonData);

                    foreach (var item in lstDynamic.data)
                    {
                    }
                }
            }

            MyLastAction myla = GetUserCurrentAction();
            var viewModel = new MyModel
            {
                UseDate = DateTime.Now,
                ShowLevel = myla.ShowLevel,
                Levels = myla.Levels,
                Version = myla.Version,
                Oper = myla.Oper,
                View = myla.View,
                Country = myla.Country,
                ChartData = GetOrgChartData(myla.Role, myla.Country, myla.ShowLevel, myla.Levels, myla.Oper, myla.Version),
                Role = myla.Role
            };

            return View("Index", viewModel);
        }

        private void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
                expandoDict[propertyName] = propertyValue;
            else
                expandoDict.Add(propertyName, propertyValue);
        }

        private ExpandoObject SearchProperty(List<dynamic> lstDyn, string propertyName, string propertyValue)
        {
            foreach (ExpandoObject eo in lstDyn)
            {
                var expandoDict = eo as IDictionary<string, object>;
                if (expandoDict[propertyName].ToString() == propertyValue) return eo;
            }

            return null;
        }

        private void InsertDynamicDataToDB(List<dynamic> lstDynamic)
        {
            string sTableName = UserData.CompanyName.ToString().Trim().ToUpper().Replace(" ", "_") + "_LevelInfos";
            string sQuery = "DROP TABLE [dbo].["+ sTableName + "] ", sFields = "", sValues="";
            comClass.ExecuteQuery(sQuery);

            sQuery = "CREATE TABLE [dbo].["+ sTableName + "] (";
            var expandoField = (IDictionary<string, object>)lstDynamic[0];
            foreach (var pairKV in expandoField)
            {
                if (pairKV.Key == "LEVEL_ID" || pairKV.Key == "USER_ID" || pairKV.Key == "DATE_UPDATED")
                    sFields += ", [" + pairKV.Key.ToString().Trim().ToUpper().Replace(" ", "_") + "][varchar](100)";
                else
                    sFields += ", [" + pairKV.Key.ToString().Trim().ToUpper().Replace(" ", "_") + "][varchar](200) NULL";
            }
            sQuery += sFields.Substring(1) + " CONSTRAINT [PK_"+ sTableName + "] PRIMARY KEY CLUSTERED ( [USER_ID] ASC, [LEVEL_ID] ASC, [DATE_UPDATED] ASC) )  ON [PRIMARY]";
            comClass.ExecuteQuery(sQuery);

            int recCount = 0; sQuery = "";
            foreach (ExpandoObject recObj in lstDynamic)
            {
                var expandoRecord = (IDictionary<string, object>)recObj;
                sQuery += "INSERT INTO [dbo].[" + sTableName + "] VALUES( ";
                foreach (var pairKV in expandoRecord)
                    sValues += ",'"+pairKV.Value.ToString().Replace("'","''")+"'";
                sQuery += sValues.Substring(1) + " );";
                sValues = "";

                recCount++;
                if (recCount >= 100)
                {
                    comClass.ExecuteQuery(sQuery);
                    recCount = 0; sQuery = "";
                }
            }
            if (recCount >= 1) comClass.ExecuteQuery(sQuery);
        }

        public ActionResult UpdateTable()
        {
            int iShowCount = 0, iNullCount = 0, iKey = 0;
            List<string> lstParentName = new List<string>();
            List<dynamic> lstDynamic = new List<dynamic>();

            var UploadExcelFile = (from uef in db.UploadFilesHeaders
                                   where uef.CompanyName == UserData.CompanyName
                                   select uef).OrderByDescending(x => x.CreatedDate).FirstOrDefault();

            string ServerMapPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/Uploads/"+ UploadExcelFile.JSONFileName);
            using (StreamReader reader = new StreamReader(ServerMapPath))
            {
                string SUP_DISPLAY_NAME = UploadExcelFile.ParentField;
                string RNUM = UploadExcelFile.FirstPositionField;
                string SNUM = "";

                string jsonData = reader.ReadToEnd();
                dynamic array = JsonConvert.DeserializeObject(jsonData);
                foreach (var item in array.data)
                {
                    if (item[SUP_DISPLAY_NAME] != null) iShowCount++;
                    if (item[SUP_DISPLAY_NAME] == null) iNullCount++;

                    if (UploadExcelFile.SerialNoFlag=="Y")
                        SNUM = (100000 + Convert.ToInt32(item[RNUM])).ToString();
                    else
                        SNUM = (100000 + iKey++).ToString();
                    if (item[SUP_DISPLAY_NAME] != null)
                    {
                        if (lstParentName.Count() >= 1)
                        {
                            var match = lstParentName.FirstOrDefault(stringToCheck => stringToCheck.Contains(item[SUP_DISPLAY_NAME].ToString().Trim()));
                            if (match == null) lstParentName.Add(item[SUP_DISPLAY_NAME].ToString().Trim());
                        }
                        else lstParentName.Add(item[SUP_DISPLAY_NAME].ToString().Trim());
                    }

                    // Employee Details
                    if (UploadExcelFile.UseFields != "")
                    {
                        string FULL_NAME = "";
                        if (UploadExcelFile.FullNameFields != "")
                        {
                            string[] FN = UploadExcelFile.FullNameFields.Split(',');
                            foreach (string strFN in FN)
                                FULL_NAME += " " + item[strFN];
                        }
                        dynamic DyObj = new ExpandoObject();
                        string[] UF = UploadExcelFile.UseFields.Split(',');
                        foreach (string strUF in UF)
                        {
                            string strField = strUF.Trim().ToUpper().Replace(" ", "_");
                            if (strField=="LEVEL_ID") AddProperty(DyObj, strField, SNUM); 
                            else if (strField == "PARENT_LEVEL_ID") AddProperty(DyObj, strField, "999999");
                            else if (strField == "VERSION") AddProperty(DyObj, strField, "1");
                            else if (strField == "FULL_NAME") AddProperty(DyObj, strField, (FULL_NAME == "") ? "" : FULL_NAME.Substring(1));
                            else if (strField == "DATE_UPDATED") AddProperty(DyObj, strField, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"));
                            else if (strField == "USER_ID") AddProperty(DyObj, strField, UserData.CompanyName);
                            else AddProperty(DyObj, strField, ((item[strUF] == null)?"":item[strUF]));
                        }
                        lstDynamic.Add(DyObj);
                    }
                }

                // Gets the Parent name
                int Index = 0;
                foreach (string pn in lstParentName)
                {
                    ExpandoObject ObjId = lstDynamic.Where(Obj => Obj.FULL_NAME==pn).FirstOrDefault();
                    if (ObjId != null)
                    {
                        var expandoLEVEL_ID = (IDictionary<string, object>)ObjId;

                        Index++;
                        var Objects = lstDynamic.Where(Obj => Obj.SUP_DISPLAY_NAME==pn).ToList();
                        foreach (ExpandoObject md in Objects)
                        {
                            var expandoPARENT_ID = (IDictionary<string, object>)md;
                            expandoPARENT_ID["PARENT_LEVEL_ID"] = expandoLEVEL_ID["LEVEL_ID"].ToString();
                        }
                    }
                }
            }

            // Insert the data into SQL table
            InsertDynamicDataToDB(lstDynamic);

            MyLastAction myla = GetUserCurrentAction();

            var viewModel = new MyModel
            {
                UseDate = DateTime.Now,
                ShowLevel = myla.ShowLevel,
                Levels = myla.Levels,
                Version = myla.Version,
                Oper = myla.Oper,
                View = myla.View,
                Country = myla.Country,
                ChartData = GetOrgChartData(myla.Role, myla.Country, myla.ShowLevel, myla.Levels, myla.Oper, myla.Version),
                Role = myla.Role
            };

            return View("Index", viewModel);
        }

        public ActionResult Settings()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult LogOut()
        {
            return Redirect("~/account/LogOff");
        }

    }
}