using HPQ.Data;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
    [DataObjectAttribute()]
    public static class Employee
    {
        public static DataTable GetUserInRole(string UserId, string ProductVersionId, string RoleCd)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetUserInRole");

            dw.CreateParameter(cmd, "@p_UserId", SqlDbType.Int, UserId);
            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);
            dw.CreateParameter(cmd, "@p_RoleCd", SqlDbType.VarChar, RoleCd);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable ListEmployees()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListEmployees");

            return dw.ExecuteCommandTable(cmd);
        }
        [DataObjectMethodAttribute(DataObjectMethodType.Update, false)]
        public static int UpdateEmployeeOdmLoginStatus(string EmployeeID, string OdmLoginStatus)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateEmployeeODMLoginStatus");

            dw.CreateParameter(cmd, "@p_EmployeeID", SqlDbType.Int, EmployeeID);
            dw.CreateParameter(cmd, "@p_ODMLoginStatus", SqlDbType.TinyInt, OdmLoginStatus);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static DataTable GetUserInfo(string UserName, string Domain)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetUserInfo");

            dw.CreateParameter(cmd, "@UserName", SqlDbType.VarChar, UserName, 30);
            dw.CreateParameter(cmd, "@Domain", SqlDbType.VarChar, Domain, 30);

            return dw.ExecuteCommandTable(cmd);
        }
        public static int GetUserID(string UserName, string Domain)
        {
            DataTable dt = GetUserInfo(UserName, Domain);
            if (dt.Rows.Count == 1)
            {
                return Convert.ToInt32(dt.Rows[0]["ID"]);
            }
            else
            {
                return 0;
            }
        }

        public static int GetUserID(string CurrentUser)
        {
            string userName = CurrentUser;
            string domain = string.Empty;

            if (CurrentUser.IndexOf("\\") > 0)
            {
                string[] currentUser = CurrentUser.Split('\\');
                userName = currentUser[1];
                domain = currentUser[0];
            }
            return GetUserID(userName, domain);
        }

        public static string GetUserName(string UserName, string Domain)
        {
            DataTable dt = GetUserInfo(UserName, Domain);
            if (dt.Rows.Count == 1)
            {
                return dt.Rows[0]["Name"].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        public static string GetUserName(string CurrentUser)
        {
            string[] currentUser = CurrentUser.Split('\\');
            return GetUserName(currentUser[1], currentUser[0]);
        }

    }
}
