using HPQ.Data;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
    public static class Product
    {
        public static DataTable ListProductStatuses()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListProductStatuses");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable ListDevCenters()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListDevCenters");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable ListPartners(string ReportType, string PartnerTypeID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListPartners");

            dw.CreateParameter(cmd, "@ReportType", SqlDbType.Int, ReportType.ToString());
            dw.CreateParameter(cmd, "@PartnerTypeID", SqlDbType.Int, PartnerTypeID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

    }
}
