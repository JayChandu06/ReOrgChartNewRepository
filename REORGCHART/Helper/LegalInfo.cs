using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.Serialization;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Web.Configuration;
using System.Text;
using System.Web.UI;
using Newtonsoft.Json;

using REORGCHART.Data;

namespace REORGCHART.Helper
{
    public class LegalInfo
    {
        string OPR_LEVEL_ID = WebConfigurationManager.AppSettings["OPR_LEVEL_ID"];
        string OPR_PARENT_ID = WebConfigurationManager.AppSettings["OPR_PARENT_ID"];
        string OPR_SEARCH_FIELD = WebConfigurationManager.AppSettings["OPR_SEARCH_FIELD"];
        string LGL_LEVEL_ID = WebConfigurationManager.AppSettings["LGL_LEVEL_ID"];
        string LGL_PARENT_ID = WebConfigurationManager.AppSettings["LGL_PARENT_ID"];

        public DataTable GetLegalInfo(string UserId, string CompanyName, string Country, string ShowLevel, string Levels, string Version)
        {

            DataTable retDT = null;

            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 0;
            cmd.CommandText = "PROC_GET_POSITION_TREE_LEGALCHART";

            cmd.Parameters.Add("@STARTPOSITION", SqlDbType.VarChar, 15).Value = ShowLevel;
            cmd.Parameters.Add("@DEPTH", SqlDbType.Int).Value = Levels;
            cmd.Parameters.Add("@COUNTRYID", SqlDbType.Int).Value = Country;            
            cmd.Parameters.Add("@VERSIONID", SqlDbType.Int).Value = Version;
            cmd.Parameters.Add("@USERID", SqlDbType.VarChar, 15).Value = UserId;

            Common csobj = new Common();
            retDT = csobj.SPReturnDataTable(cmd);

            return retDT;
        }
    }
}