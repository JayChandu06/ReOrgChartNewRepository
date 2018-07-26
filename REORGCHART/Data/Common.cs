using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using OrgChartModelling;

using OrgChartModelling.Utils;

namespace REORGCHART.Data
{
    public class Common
    {
        static clsSecurity secutiry = new clsSecurity();
        public string g_ConnString = System.Configuration.ConfigurationManager.ConnectionStrings["ORG_CHART"].ConnectionString.ToString();

        public string getConnStr()
        {
            return g_ConnString;
        }

        public DataTable SPReturnDataTable(SqlCommand sqlCmd)
        {
            SqlConnection SqlCon = new SqlConnection(g_ConnString);
            SqlDataAdapter SqlAdapter = new SqlDataAdapter();
            DataSet SqlDataSet = new DataSet();
            DataTable sqlDataTable = new DataTable();
            sqlCmd.Connection = SqlCon;
            sqlCmd.CommandTimeout = 0;
            try
            {
                SqlCon.Open();
                SqlAdapter.SelectCommand = sqlCmd;
                SqlAdapter.Fill(SqlDataSet, "TableName");
                sqlDataTable = SqlDataSet.Tables["TableName"];

                return sqlDataTable;
            }
            catch (Exception ex)
            {
                return GetMessageTable(AppConstants.OUTPUT_STATUS_COLUMN_NAME, AppConstants.OUTPUT_STATUS_FALSE, AppConstants.OUTPUT_MESSAGE_COLUMN_NAME, ex.Message);
            }
            finally
            {
                if (SqlCon.State == ConnectionState.Open)
                    SqlCon.Close();
                SqlCon.Dispose();
                sqlCmd.Dispose();
                SqlAdapter.Dispose();
                SqlDataSet.Dispose();
                sqlDataTable.Dispose();
                sqlCmd = null;
                SqlAdapter = null;
                SqlDataSet = null;
                sqlDataTable = null;
            }

        }

        public DataTable SPReturnDataTable(SqlCommand sqlCmd, SqlConnection SqlCon)
        {
            SqlDataAdapter SqlAdapter = new SqlDataAdapter();
            DataSet SqlDataSet = new DataSet();
            DataTable sqlDataTable = new DataTable();
            sqlCmd.Connection = SqlCon;
            sqlCmd.CommandTimeout = 0;
            try
            {
                //SqlCon.Open();
                SqlAdapter.SelectCommand = sqlCmd;
                SqlAdapter.Fill(SqlDataSet, "TableName");
                sqlDataTable = SqlDataSet.Tables["TableName"];

                return sqlDataTable;
            }
            catch (Exception ex)
            {
                return GetMessageTable(AppConstants.OUTPUT_STATUS_COLUMN_NAME, AppConstants.OUTPUT_STATUS_FALSE, AppConstants.OUTPUT_MESSAGE_COLUMN_NAME, ex.Message);
            }
            finally
            {
                sqlCmd.Dispose();
                SqlAdapter.Dispose();
                SqlDataSet.Dispose();
                sqlDataTable.Dispose();
                sqlCmd = null;
                SqlAdapter = null;
                SqlDataSet = null;
                sqlDataTable = null;
            }

        }

        public DataSet SPReturnDataSet(SqlCommand sqlCmd)
        {
            SqlConnection SqlCon = new SqlConnection(g_ConnString);
            SqlDataAdapter SqlAdapter = new SqlDataAdapter();
            DataSet SqlDataSet = new DataSet();
            DataTable sqlDataTable = new DataTable();
            sqlCmd.Connection = SqlCon;
            sqlCmd.CommandTimeout = 0;
            try
            {
                SqlCon.Open();
                SqlAdapter.SelectCommand = sqlCmd;
                SqlAdapter.Fill(SqlDataSet);

                return SqlDataSet;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (SqlCon.State == ConnectionState.Open)
                    SqlCon.Close();
                SqlCon.Dispose();
                sqlCmd.Dispose();
                SqlAdapter.Dispose();
                SqlDataSet.Dispose();
                sqlDataTable.Dispose();
                sqlCmd = null;
                SqlAdapter = null;
                SqlDataSet = null;
                sqlDataTable = null;
            }

        }

        public DataTable SQLReturnDataTable(string strQuery)
        {
            SqlConnection SqlCon = new SqlConnection(g_ConnString);
            SqlDataAdapter SqlAdapter = new SqlDataAdapter();
            DataSet SqlDataSet = new DataSet();
            DataTable sqlDataTable = new DataTable();
            SqlCommand sqlCmd = new SqlCommand(strQuery, SqlCon);
            sqlCmd.CommandTimeout = 0;
            try
            {
                SqlCon.Open();
                SqlAdapter.SelectCommand = sqlCmd;
                SqlAdapter.Fill(SqlDataSet, "TableName");
                sqlDataTable = SqlDataSet.Tables["TableName"];

                return sqlDataTable;
            }
            catch (Exception ex)
            {
                return GetMessageTable(AppConstants.OUTPUT_STATUS_COLUMN_NAME, AppConstants.OUTPUT_STATUS_FALSE, AppConstants.OUTPUT_MESSAGE_COLUMN_NAME, ex.Message);
            }
            finally
            {
                if (SqlCon.State == ConnectionState.Open)
                    SqlCon.Close();
                SqlCon.Dispose();
                sqlCmd.Dispose();
                SqlAdapter.Dispose();
                SqlDataSet.Dispose();
                sqlDataTable.Dispose();
                sqlCmd = null;
                SqlAdapter = null;
                SqlDataSet = null;
                sqlDataTable = null;
            }
        }

        public DataTable SQLSchemaDataTable(string strTable)
        {
            SqlConnection SqlCon = new SqlConnection(g_ConnString);
            DataTable sqlDataTable = null;
            try
            {
                SqlCon.Open();
                sqlDataTable = SqlCon.GetSchema("Tables");

                return sqlDataTable;
            }
            catch (Exception ex)
            {
                return GetMessageTable(AppConstants.OUTPUT_STATUS_COLUMN_NAME, AppConstants.OUTPUT_STATUS_FALSE, AppConstants.OUTPUT_MESSAGE_COLUMN_NAME, ex.Message);
            }
            finally
            {
                if (SqlCon.State == ConnectionState.Open)
                    SqlCon.Close();
                SqlCon.Dispose();
                sqlDataTable.Dispose();
                sqlDataTable = null;
            }
        }

        public string returnVal(string qry)
        {
            string functionReturnValue = null;
            SqlConnection sqlcon = new SqlConnection(g_ConnString);
            SqlCommand sqlcmd = new SqlCommand(qry, sqlcon);
            sqlcmd.CommandTimeout = 0;
            try
            {
                sqlcon.Open();
                if (sqlcmd.ExecuteScalar() == System.DBNull.Value)
                {
                    return "0";
                }
                else
                {
                    functionReturnValue = sqlcmd.ExecuteScalar().ToString();
                    if (functionReturnValue == null)
                    {
                        return "0";
                    }
                    else
                    {
                        return functionReturnValue;
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (sqlcon.State == ConnectionState.Open)
                    sqlcon.Close();
                sqlcon.Dispose();
                sqlcmd.Dispose();
                sqlcon = null;
                sqlcmd = null;
            }
            // return functionReturnValue;
        }

        public void ExecuteQuery(string strName)
        {
            SqlConnection SqlCon = new SqlConnection(g_ConnString);
            SqlCommand sqlcmd = new SqlCommand(strName, SqlCon);
            sqlcmd.CommandTimeout = 0;
            try
            {
                SqlCon.Open();
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                sqlcmd.Dispose();
                if (SqlCon.State == ConnectionState.Open)
                    SqlCon.Close();
                SqlCon.Dispose();
                SqlCon = null;
                sqlcmd = null;
            }
        }

        public void ExecuteTransactionQuery(string strName, SqlTransaction transaction, SqlConnection SqlCon)
        {
            SqlCommand sqlcmd = new SqlCommand(strName, SqlCon);
            try
            {
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandTimeout = 0;
                sqlcmd.Transaction = transaction;
                sqlcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
                //  File.AppendAllText(g_LogLocation + "ErrorLog.txt", "ExecuteQuery - " + System.DateTime.UtcNow + " - " + ex.Message + System.Environment.NewLine);
                //return false;
            }
            finally
            {
                sqlcmd.Dispose();
                sqlcmd = null;
            }
        }

        public void ExecuteTransactionProcedure(SqlCommand sqlcmd, SqlTransaction transaction)
        {
            try
            {
                sqlcmd.Transaction = transaction;
                sqlcmd.CommandTimeout = 0;
                int retval = sqlcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
                //  File.AppendAllText(g_LogLocation + "ErrorLog.txt", "ExecuteQuery - " + System.DateTime.UtcNow + " - " + ex.Message + System.Environment.NewLine);
                //return false;
            }
            finally
            {
                sqlcmd.Dispose();
                sqlcmd = null;
            }
        }

        public void ExecuteQuery(string strName, SqlConnection SqlCon)
        {
            SqlCommand sqlcmd = new SqlCommand(strName, SqlCon);
            try
            {
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandTimeout = 0;
                sqlcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
                //  File.AppendAllText(g_LogLocation + "ErrorLog.txt", "ExecuteQuery - " + System.DateTime.UtcNow + " - " + ex.Message + System.Environment.NewLine);
                //return false;
            }
            finally
            {
                sqlcmd.Dispose();
                sqlcmd = null;
            }
        }

        public static DataTable GetMessageTable(string StatusColumnName, string StatusColumnValue, string MessageColumnName, string MessageColumnValue)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn(StatusColumnName));
            dt.Columns.Add(new DataColumn(MessageColumnName));

            DataRow row = dt.NewRow();
            row[StatusColumnName] = StatusColumnValue;
            row[MessageColumnName] = MessageColumnValue;

            dt.Rows.Add(row);

            return dt;
        }

        public static string GetUser()
        {

            string userInfo = System.Web.HttpContext.Current.User.Identity.Name.ToString();
            string uName = "SYS";
            if (userInfo != null && userInfo != "" && userInfo != " ")
            {
                if (userInfo.Split('\\').Length > 1)
                    uName = userInfo.Split('\\')[1];
                else
                    uName = userInfo.Split('\\')[0];
            }
            
            uName = "JHAAM2";
           // uName = "VERMARU4";
            //uName = "JJA3";
            return uName;
        }

    }
}