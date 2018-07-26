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
    // Org chart creation 
    public class LevelInfo
    {
        public DataTable GetLevelInfo(string UserId, string CompanyName, string UserType, 
                                      string ShowLevel, string Levels, string Version)
        {
            DataTable retDT = null;

            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 0;
            cmd.CommandText = "PROC_GET_POSITION_TREE_OPERATIONALCHART";

            cmd.Parameters.Add("@STARTPOSITION", SqlDbType.VarChar, 15).Value = ShowLevel;
            cmd.Parameters.Add("@DEPTH", SqlDbType.VarChar, 15).Value = Levels;
            cmd.Parameters.Add("@VERSION", SqlDbType.VarChar, 150).Value = Version;
            cmd.Parameters.Add("@USERTYPE", SqlDbType.VarChar, 50).Value = UserType;
            cmd.Parameters.Add("@COMPANYNAME", SqlDbType.VarChar, 150).Value = CompanyName;
            cmd.Parameters.Add("@USERID", SqlDbType.VarChar, 150).Value = UserId;

            Common csobj = new Common();
            retDT = csobj.SPReturnDataTable(cmd);

            return retDT;
        }
    }
}


