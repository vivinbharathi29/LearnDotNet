using HPQ.Data;
using System;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
    public static class Images
    {
        public static DataTable spGetImageProperties(string ImageID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.spGetImageProperties");

            dw.CreateParameter(cmd, "@ImageID", SqlDbType.Int, ImageID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable usp_ListImageDriveDefinitions(Boolean ShowAll)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.usp_ListImageDriveDefinitions");

            dw.CreateParameter(cmd, "@p_ShowAll", SqlDbType.Bit, ShowAll ? "1" : "0");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static int usp_ImageDriveDefinitionInsert(string DivCd, string SiteCd, string DriveName, string PartNo, string PartNoRev, string IsAssembly, string Active, string LastUpdUser, string LastUpdDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.usp_ImageDriveDefinitionInsert");

            dw.CreateParameter(cmd, "@DivCd", SqlDbType.VarChar, DivCd.ToString());
            dw.CreateParameter(cmd, "@SiteCd", SqlDbType.VarChar, SiteCd.ToString());
            dw.CreateParameter(cmd, "@DriveName", SqlDbType.VarChar, DriveName.ToString());
            dw.CreateParameter(cmd, "@PartNo", SqlDbType.VarChar, PartNo.ToString());
            dw.CreateParameter(cmd, "@PartNoRev", SqlDbType.VarChar, PartNoRev.ToString());
            dw.CreateParameter(cmd, "@IsAssembly", SqlDbType.Bit, IsAssembly.ToString());
            dw.CreateParameter(cmd, "@Active", SqlDbType.Bit, Active.ToString());
            dw.CreateParameter(cmd, "@LastUpdUser", SqlDbType.VarChar, LastUpdUser.ToString());
            dw.CreateParameter(cmd, "@LastUpdDate", SqlDbType.DateTime, LastUpdDate.ToString());

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }

        public static int usp_ImageDriveDefinitionUpdate(string ID, string DivCd, string SiteCd, string DriveName, string PartNo, string PartNoRev, string IsAssembly, string Active, string LastUpdUser, string LastUpdDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.usp_ImageDriveDefinitionUpdate");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ID.ToString());
            dw.CreateParameter(cmd, "@DivCd", SqlDbType.VarChar, DivCd.ToString());
            dw.CreateParameter(cmd, "@SiteCd", SqlDbType.VarChar, SiteCd.ToString());
            dw.CreateParameter(cmd, "@DriveName", SqlDbType.VarChar, DriveName.ToString());
            dw.CreateParameter(cmd, "@PartNo", SqlDbType.VarChar, PartNo.ToString());
            dw.CreateParameter(cmd, "@PartNoRev", SqlDbType.VarChar, PartNoRev.ToString());
            dw.CreateParameter(cmd, "@IsAssembly", SqlDbType.Bit, IsAssembly.ToString());
            dw.CreateParameter(cmd, "@Active", SqlDbType.Bit, Active.ToString());
            dw.CreateParameter(cmd, "@LastUpdUser", SqlDbType.VarChar, LastUpdUser.ToString());
            dw.CreateParameter(cmd, "@LastUpdDate", SqlDbType.DateTime, LastUpdDate.ToString());

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }

        public static int usp_UpdateImagesImageDriveDefinitionId(string ImageId, string ImageDriveDefinitionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.usp_UpdateImagesImageDriveDefinitionId");

            dw.CreateParameter(cmd, "@p_ImageId", SqlDbType.Int, ImageId.ToString());
            dw.CreateParameter(cmd, "@p_ImageDriveDefinitionId", SqlDbType.Int, ImageDriveDefinitionId.ToString());

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }

    }
}
