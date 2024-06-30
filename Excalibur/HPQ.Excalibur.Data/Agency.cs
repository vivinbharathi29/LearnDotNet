using HPQ.Data;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
    public static class Agency
    {
        public static DataTable AgencyStatusSelectDocumentsForPmView(string ProductVersionId, string DeliverableVersionId, string WorkflowStatus)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_AgencySelectDocumentsForPmView");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId.ToString());
            dw.CreateParameter(cmd, "@p_DeliverableVersionId", SqlDbType.Int, DeliverableVersionId.ToString());
            dw.CreateParameter(cmd, "@p_WorkflowStatus", SqlDbType.Bit, WorkflowStatus.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public static DataTable AgencyStatusSelectBlockedCountries(string ProductVersionId, string DeliverableVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_AgencyStatusSelectBlockedCountries");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId.ToString());
            dw.CreateParameter(cmd, "@p_DeliverableVersionId", SqlDbType.Int, DeliverableVersionId.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public static DataTable AgencyStatusSelectCompletedCountries(string ProductVersionId, string DeliverableVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_AgencyStatusSelectCompletedCountries");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId.ToString());
            dw.CreateParameter(cmd, "@p_DeliverableVersionId", SqlDbType.Int, DeliverableVersionId.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

    }
}
