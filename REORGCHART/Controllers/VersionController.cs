using System;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Collections.Generic;
using System.Dynamic;

using Newtonsoft.Json;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.AspNet.Identity;

using REORGCHART.Data;
using REORGCHART.Models;

namespace REORGCHART.Controllers
{
    [Authorize]
    public class VersionController : Controller
    {
        Models.DBContext db = new Models.DBContext();
        ApplicationUser UserData = System.Web.HttpContext.Current.GetOwinContext().GetUserManager<ApplicationUserManager>().FindById(System.Web.HttpContext.Current.User.Identity.GetUserId());

        Common comClass = new Common();

        // GET: Version
        public ActionResult Index()
        {
            return View();
        }

        // GET: Version
        public ActionResult UploadVersion()
        {
            return View();
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
            string sQuery = "DROP TABLE [dbo].[" + sTableName + "] ", sFields = "", sValues = "";
            comClass.ExecuteQuery(sQuery);

            sQuery = "CREATE TABLE [dbo].[" + sTableName + "] (";
            var expandoField = (IDictionary<string, object>)lstDynamic[0];
            foreach (var pairKV in expandoField)
            {
                if (pairKV.Key == "LEVEL_ID" || pairKV.Key == "USER_ID" || pairKV.Key == "DATE_UPDATED")
                    sFields += ", [" + pairKV.Key.ToString().Trim().ToUpper().Replace(" ", "_") + "][varchar](100)";
                else
                    sFields += ", [" + pairKV.Key.ToString().Trim().ToUpper().Replace(" ", "_") + "][varchar](200) NULL";
            }
            sQuery += sFields.Substring(1) + " CONSTRAINT [PK_" + sTableName + "] PRIMARY KEY CLUSTERED ( [USER_ID] ASC, [LEVEL_ID] ASC, [DATE_UPDATED] ASC) )  ON [PRIMARY]";
            comClass.ExecuteQuery(sQuery);

            int recCount = 0; sQuery = "";
            foreach (ExpandoObject recObj in lstDynamic)
            {
                var expandoRecord = (IDictionary<string, object>)recObj;
                sQuery += "INSERT INTO [dbo].[" + sTableName + "] VALUES( ";
                foreach (var pairKV in expandoRecord)
                    sValues += ",'" + pairKV.Value.ToString().Replace("'", "''") + "'";
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

        private void InsertPlayerDataToDB(List<dynamic> lstDynamic, string VersionNo)
        {
            string sTableName = UserData.CompanyName.ToString().Trim().ToUpper().Replace(" ", "_") + "_LevelInfos", sValues="";
            string sQuery = "DELETE FROM [dbo].[" + sTableName + "] WHERE USER_ID='"+UserData.UserName+"' AND VERSION='"+ VersionNo + "'";
            comClass.ExecuteQuery(sQuery);

            int recCount = 0; sQuery = "";
            foreach (ExpandoObject recObj in lstDynamic)
            {
                var expandoRecord = (IDictionary<string, object>)recObj;
                sQuery += "INSERT INTO [dbo].[" + sTableName + "] VALUES( ";
                foreach (var pairKV in expandoRecord)
                    sValues += ",'" + pairKV.Value.ToString().Replace("'", "''") + "'";
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

        private void UpdateTableWithJSON(string RoleType, string VersionNo)
        {
            int iShowCount = 0, iNullCount = 0, iKey = 0;
            List<string> lstParentName = new List<string>();
            List<dynamic> lstDynamic = new List<dynamic>();

            var UploadExcelFile = (from uef in db.UploadFilesHeaders
                                   where uef.CompanyName == UserData.CompanyName
                                   select uef).OrderByDescending(x => x.CreatedDate).FirstOrDefault();

            string ServerMapPath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data/Uploads/" + UploadExcelFile.JSONFileName);
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

                    if (UploadExcelFile.SerialNoFlag == "Y")
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
                            if (strField == "LEVEL_ID") AddProperty(DyObj, strField, SNUM);
                            else if (strField == "PARENT_LEVEL_ID") AddProperty(DyObj, strField, "999999");
                            else if (strField == "VERSION") AddProperty(DyObj, strField, VersionNo);
                            else if (strField == "FULL_NAME") AddProperty(DyObj, strField, (FULL_NAME == "") ? "" : FULL_NAME.Substring(1));
                            else if (strField == "DATE_UPDATED") AddProperty(DyObj, strField, DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss"));
                            else if (strField == "USER_ID") AddProperty(DyObj, strField, UserData.CompanyName);
                            else AddProperty(DyObj, strField, ((item[strUF] == null) ? "" : item[strUF]));
                        }
                        lstDynamic.Add(DyObj);
                    }
                }

                // Gets the Parent name
                int Index = 0;
                foreach (string pn in lstParentName)
                {
                    ExpandoObject ObjId = lstDynamic.Where(Obj => Obj.FULL_NAME == pn).FirstOrDefault();
                    if (ObjId != null)
                    {
                        var expandoLEVEL_ID = (IDictionary<string, object>)ObjId;

                        Index++;
                        var Objects = lstDynamic.Where(Obj => Obj.SUP_DISPLAY_NAME == pn).ToList();
                        foreach (ExpandoObject md in Objects)
                        {
                            var expandoPARENT_ID = (IDictionary<string, object>)md;
                            expandoPARENT_ID["PARENT_LEVEL_ID"] = expandoLEVEL_ID["LEVEL_ID"].ToString();
                        }
                    }
                }
            }

            // Insert the data into SQL table
            if (RoleType == "Player")
                InsertPlayerDataToDB(lstDynamic, VersionNo);
            else if (RoleType == "Finalyzer")
                InsertDynamicDataToDB(lstDynamic);
        }

        [HttpPost]
        public JsonResult SaveVersionInfo(string selFP, string txtFP, string chkSN, 
                                          string selNL, string selPL, string txtUN, 
                                          string chkFT, string txtFN)
        {
            var UserLastAction = (from ula in db.UserLastActions
                                  where ula.UserId == UserData.UserName && ula.Company==UserData.CompanyName
                                  select ula).FirstOrDefault();

            if (UserLastAction!=null)
            {
                var UploadFilesHeader = (from ufh in db.UploadFilesHeaders
                                         where ufh.UserId == UserData.UserName && 
                                               ufh.CompanyName == UserData.CompanyName && 
                                               ufh.Role== UserLastAction.Role
                                         select ufh).FirstOrDefault();
                if (UploadFilesHeader!=null)
                {
                    UploadFilesHeader.SerialNoFlag = chkSN;
                    UploadFilesHeader.FirstPositionField = selFP;
                    UploadFilesHeader.FirstPosition = txtFP;
                    UploadFilesHeader.KeyField = selNL;
                    UploadFilesHeader.ParentField = selPL;
                    UploadFilesHeader.FullNameFields = txtUN;
                    UploadFilesHeader.FileType = chkFT;
                    UploadFilesHeader.JSONFileName = txtFN;

                    UploadFilesHeader.VersionNo++;
                    UploadFilesHeader.CurrentVersionNo = UploadFilesHeader.VersionNo;
                    UserLastAction.Version = UploadFilesHeader.VersionNo.ToString();
                }
                else
                {
                    UploadFilesHeaders UFH = new UploadFilesHeaders();

                    UFH.SerialNoFlag = chkSN;
                    UFH.FirstPositionField = selFP;
                    UFH.FirstPosition = txtFP;
                    UFH.KeyField = selNL;
                    UFH.ParentField = selPL;
                    UFH.FullNameFields = txtUN;
                    UFH.FileType = chkFT;
                    UFH.JSONFileName = txtFN;
                    UFH.Role = UserLastAction.Role;
                    UFH.VersionNo = 1;
                    UFH.CurrentVersionNo = 1;
                    UserLastAction.Version = "1";

                    db.UploadFilesHeaders.Add(UFH);
                }

                UploadFileDetails UFD = new UploadFileDetails();

                UFD.JSONFileName = txtFN;
                UFD.KeyDate = DateTime.Now;
                UFD.VersionNo = Convert.ToInt32(UserLastAction.Version);
                UFD.VersionStatus="P";
                UFD.Role= UserLastAction.Role;
                UFD.CompanyName = UserData.CompanyName;
                UFD.UserId = UserData.UserName;

                db.UploadFileDetails.Add(UFD);

                db.SaveChanges();
            }

            // Updates the Table with new Version
            if (chkFT=="JSON")
                UpdateTableWithJSON(UserLastAction.Role, UserLastAction.Role);
            else if (chkFT == "XLSX")
                UpdateTableWithJSON(UserLastAction.Role, UserLastAction.Role);

            return Json(new
            {
                Success = "Yes"
            });
        }

        public string[] CheckFields(string FileName)
        {
            string[] FieldsInf = { "", "" };
            string MissingFields = "", ShowFields = "";
            var UploadExcelFile = (from uef in db.UploadFilesHeaders
                                   where uef.CompanyName == UserData.CompanyName
                                   select uef).OrderByDescending(x => x.CreatedDate).FirstOrDefault();

            using (StreamReader reader = new StreamReader(FileName))
            {
                string jsonData = reader.ReadToEnd();
                dynamic array = JsonConvert.DeserializeObject(jsonData);
                int Index = 0, ErrCount=0;
                foreach (var item in array.data)
                {
                    if (UploadExcelFile.UseFields != "")
                    {
                        Index++;
                        string[] UF = UploadExcelFile.UseFields.Split(',');
                        foreach (string strUF in UF)
                        {
                            string strField = strUF.Trim().ToUpper().Replace(" ", "_");
                            if (!(strField == "LEVEL_ID" || strField == "PARENT_LEVEL_ID" || strField == "VERSION" ||
                                  strField == "FULL_NAME" || strField == "DATE_UPDATED" || strField == "USER_ID"))
                            {
                                try
                                {
                                    if (item[strUF] == null)
                                    { 
                                        if (ErrCount <= 100) MissingFields += "," + Index.ToString() + "~:~" + strUF;
                                        ErrCount++;
                                    }
                                    else
                                    {
                                        if (Index == 1) ShowFields += "," + strUF;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (ErrCount <= 100) MissingFields += "," + Index.ToString() + "~:~" + strUF;
                                    ErrCount++;
                                }
                            }
                        }
                    }

                }
            }

            FieldsInf[0] = (MissingFields == "") ? "" : MissingFields.Substring(1);
            FieldsInf[1] = (ShowFields == "") ? "" : ShowFields.Substring(1);

            return FieldsInf;
        }

        [HttpPost]
        public ActionResult UploadFile()
        {
            string[] Fields = { "", "" };
            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int Idx = 0; Idx < files.Count; Idx++)
                    {
                        //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                        //string filename = Path.GetFileName(Request.Files[i].FileName);  

                        HttpPostedFileBase file = files[Idx];
                        string fname;

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
                        fname = Path.Combine(Server.MapPath("~/App_Data/Uploads/"), fname);
                        file.SaveAs(fname);

                        Fields=CheckFields(fname);
                    }

                    // Returns message that successfully uploaded  
                    return Json(new
                    {
                        Success = (Fields[0] == "") ? "Yes" : "No",
                        Message= (Fields[0] != "") ? "Missing Fields" : "",
                        MF= Fields[0],
                        SF = Fields[1]
                    });

                }
                catch (Exception ex)
                {
                    return Json(new
                    {
                        Success = "No",
                        Message = ex.Message,
                        MF = "",
                        SF = Fields[1]
                    });
                }
            }
            else
            {
                return Json(new
                {
                    Success = "No",
                    Message = "No files selected.",
                    MF = "",
                    SF = Fields[1]
                });
            }
        }
    }
}