using HPQ.Data;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
        /*#########################################################
        #########################################################
        ########################################################
        #########################################################

        ATTENTION!
        this file-salad also called Excalibur has some aspx files WITH THE CODE IN THE ASPX PAGE in a <script> section
        some other files are using SQLdatasource and objectdatasource in the markup;
        The methods in HPQ.Excalibur.Data and HPQ.Excalibur.Service used by those aspx pages will show 0 references if you have CodeLenses active
        Do not remove any method from this wonky project unless you are sure 100% that it is not used by any aspx file;
        In other words, do a text search.

        YOU HAVE BEEN WARNED.


        #########################################################
        #########################################################
        #########################################################
        #########################################################*/

    [DataObjectAttribute()]
    public static class Service
    {
        public static DataTable ListServiceGeos()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListServiceGeos");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        [DataObjectMethodAttribute(DataObjectMethodType.Select, false)]
        public static DataTable SelectServiceFamilyPartnerDetails(string ID, string ServiceFamilyPn, string PartnerID, string ServiceGeoID, string Status, string ServicePartnerTypeCode, out string ReturnCd, out string ReturnDesc)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceFamilyPartnerDetails");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ID);
            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn);
            dw.CreateParameter(cmd, "@PartnerID", SqlDbType.Int, PartnerID);
            dw.CreateParameter(cmd, "@ServiceGeoID", SqlDbType.Int, ServiceGeoID);
            dw.CreateParameter(cmd, "@Status", SqlDbType.Char, Status);
            dw.CreateParameter(cmd, "@ServicePartnerTypeCode", SqlDbType.VarChar, ServicePartnerTypeCode);
            dw.CreateParameter(cmd, "@RETURN_CODE", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_DESC", SqlDbType.NVarChar, string.Empty, 510, ParameterDirection.Output);

            DataTable dt = dw.ExecuteCommandTable(cmd);

            ReturnCd = cmd.Parameters["@RETURN_CODE"].Value.ToString();
            ReturnDesc = cmd.Parameters["@RETURN_DESC"].Value.ToString();

            return dt;

        }

        [DataObjectMethodAttribute(DataObjectMethodType.Insert, false)]
        public static int InsertServiceFamilyPartnerDetails(string ServiceFamilyPn, string PartnerID, string ServiceGeoID, string Status, string ServicePartnerTypeCode, string LastUpdUser, out string ReturnCd, out string ReturnDesc)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertServiceFamilyPartnerDetails");

            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn);
            dw.CreateParameter(cmd, "@PartnerID", SqlDbType.Int, PartnerID);
            dw.CreateParameter(cmd, "@ServiceGeoID", SqlDbType.Int, ServiceGeoID);
            dw.CreateParameter(cmd, "@Status", SqlDbType.Char, Status);
            dw.CreateParameter(cmd, "@ServicePartnerTypeCode", SqlDbType.VarChar, ServicePartnerTypeCode);
            dw.CreateParameter(cmd, "@LastUpdUser", SqlDbType.VarChar, LastUpdUser);
            dw.CreateParameter(cmd, "@NEW_OR_EXISTING_ID", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_CODE", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_DESC", SqlDbType.NVarChar, string.Empty, 510, ParameterDirection.Output);

            int iReturnValue = dw.ExecuteCommandNonQuery(cmd);

            ReturnCd = cmd.Parameters["@RETURN_CODE"].Value.ToString();
            ReturnDesc = cmd.Parameters["@RETURN_DESC"].Value.ToString();

            return Convert.ToInt32(cmd.Parameters["@NEW_OR_EXISTING_ID"].Value);
        }

        [DataObjectMethodAttribute(DataObjectMethodType.Update, false)]
        public static int UpdateServiceFamilyPartnerDetails(string ID, string ServiceFamilyPn, string PartnerID, string ServiceGeoID, string Status, string ServicePartnerTypeCode, string LastUpdUser, out string ReturnCd, out string ReturnDesc)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateServiceFamilyPartnerDetails");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ID);
            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn);
            dw.CreateParameter(cmd, "@PartnerID", SqlDbType.Int, PartnerID);
            dw.CreateParameter(cmd, "@ServiceGeoID", SqlDbType.Int, ServiceGeoID);
            dw.CreateParameter(cmd, "@Status", SqlDbType.Char, Status);
            dw.CreateParameter(cmd, "@ServicePartnerTypeCode", SqlDbType.VarChar, ServicePartnerTypeCode);
            dw.CreateParameter(cmd, "@LastUpdUser", SqlDbType.VarChar, LastUpdUser);
            dw.CreateParameter(cmd, "@NEW_OR_EXISTING_ID", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_CODE", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_DESC", SqlDbType.NVarChar, string.Empty, 510, ParameterDirection.Output);

            int iReturnValue = dw.ExecuteCommandNonQuery(cmd);

            ReturnCd = cmd.Parameters["@RETURN_CODE"].Value.ToString();
            ReturnDesc = cmd.Parameters["@RETURN_DESC"].Value.ToString();

            return Convert.ToInt32(cmd.Parameters["@NEW_OR_EXISTING_ID"].Value);
        }


        public static DataTable ListOsspPartners()
        {
            return Product.ListPartners(string.Empty, "2");
        }

        public static DataTable SelectServiceFamilyOsspAssignments(string UserId, string StatusId, string DevCenter, string BusinessType, string OsspId, out string ReturnCd, out string ReturnDesc)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("rpt_ServiceFamilyOSSPAssignments");

            dw.CreateParameter(cmd, "@USER_ID", SqlDbType.Int, UserId.ToString());
            dw.CreateParameter(cmd, "@STATUS_ID", SqlDbType.Int, StatusId.ToString());
            dw.CreateParameter(cmd, "@DEV_CENTER", SqlDbType.Int, DevCenter.ToString());
            dw.CreateParameter(cmd, "@BUSINESS_TYPE", SqlDbType.Int, BusinessType.ToString());
            dw.CreateParameter(cmd, "@OSSP_ID", SqlDbType.Int, OsspId.ToString());
            dw.CreateParameter(cmd, "@RETURN_CODE", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@RETURN_DESC", SqlDbType.NVarChar, string.Empty, 510, ParameterDirection.Output);


            DataTable dt = dw.ExecuteCommandTable(cmd);

            ReturnCd = cmd.Parameters["@RETURN_CODE"].Value.ToString();
            ReturnDesc = cmd.Parameters["@RETURN_DESC"].Value.ToString();

            return dt;
        }


        public static DataTable GetProductsOnCommodityMatrix(int Division)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListProductsOnCommodityMatrix");

            dw.CreateParameter(cmd, "@Division", SqlDbType.Int, Division.ToString());

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetProductLines()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_ListProductLines");

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetPlatformAssignmentMetrics(string sPlatform, string sODM, string sGPLM, string sBomAnalysis, string sPsm, string sServiceFamilyPn, string sProjectNumber, string sBusiness, string StartDate, string EndDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_ListPlatformAssignmentMetrics");

            //if (sPlatform != string.Empty) dw.CreateParameter(cmd, "@Platform", SqlDbType.Int, sPlatform);
            if (sPlatform != string.Empty) dw.CreateParameter(cmd, "@Platform", SqlDbType.VarChar, sPlatform, 200);
            if (sODM != string.Empty) dw.CreateParameter(cmd, "@Odm", SqlDbType.Int, sODM);
            if (sGPLM != string.Empty) dw.CreateParameter(cmd, "@Gplm", SqlDbType.Int, sGPLM);
            if (sBomAnalysis != string.Empty) dw.CreateParameter(cmd, "@BomAnalysis", SqlDbType.Int, sBomAnalysis);
            if (sPsm != string.Empty) dw.CreateParameter(cmd, "@Psm", SqlDbType.Int, sPsm);
            if (sServiceFamilyPn != string.Empty) dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, sServiceFamilyPn, 10);
            if (sProjectNumber != string.Empty) dw.CreateParameter(cmd, "@ProjectNumber", SqlDbType.NVarChar, sProjectNumber, 50);
            if (sBusiness != string.Empty) dw.CreateParameter(cmd, "@Business", SqlDbType.Int, sBusiness);
            if (StartDate != string.Empty) dw.CreateParameter(cmd, "@StartDate", SqlDbType.DateTime, StartDate);
            if (EndDate != string.Empty) dw.CreateParameter(cmd, "@EndDate", SqlDbType.DateTime, EndDate);


            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetPlatformAssignmentDesktop(string sPlatform, string sODM, string sGPLM, string sProductLine, string sServiceFamilyPn, string sProjectNumber)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_ListPlatformAssignmentDesktop");

            if (sPlatform != string.Empty) dw.CreateParameter(cmd, "@Platform", SqlDbType.VarChar, sPlatform, 200);
            if (sODM != string.Empty) dw.CreateParameter(cmd, "@Odm", SqlDbType.Int, sODM);
            if (sGPLM != string.Empty) dw.CreateParameter(cmd, "@Gplm", SqlDbType.Int, sGPLM);
            if (sProductLine != string.Empty) dw.CreateParameter(cmd, "@ProductLine", SqlDbType.Int, sProductLine);
            if (sServiceFamilyPn != string.Empty) dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, sServiceFamilyPn, 10);
            if (sProjectNumber != string.Empty) dw.CreateParameter(cmd, "@ProjectNumber", SqlDbType.NVarChar, sProjectNumber, 50);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetProductFamilies()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListProductsActive");

            return dw.ExecuteCommandTable(cmd);
        }

        public static int InsertDesktopPlatform(string ProductFamilyId, string ProductLineId, string ServiceFamilyPn, string PlatformName, string PlatformDescription, string PartnerID, string GPLM, string FCSDate, string EndOfService)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceInsertDesktopPlatform");

            dw.CreateParameter(cmd, "@ProductFamilyID", SqlDbType.Int, ProductFamilyId);
            dw.CreateParameter(cmd, "@ProductLineID", SqlDbType.Int, ProductLineId);
            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.NVarChar, ServiceFamilyPn, 10);
            dw.CreateParameter(cmd, "@DOTSNAME", SqlDbType.NVarChar, PlatformName, 30);
            dw.CreateParameter(cmd, "@Description", SqlDbType.NVarChar, PlatformDescription, 200);
            dw.CreateParameter(cmd, "@GPLM", SqlDbType.Int, GPLM);
            dw.CreateParameter(cmd, "@PartnerID", SqlDbType.Int, PartnerID);
            if (FCSDate != string.Empty) dw.CreateParameter(cmd, "@FCSDate", SqlDbType.DateTime, FCSDate);
            if (EndOfService != string.Empty) dw.CreateParameter(cmd, "@EndOfService", SqlDbType.DateTime, EndOfService);


            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static int UpdateDesktopDates(string ProductVersionId, string ServiceFamilyPn, string FCSDate, string EndOfService)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceUpdateDesktopDates");

            dw.CreateParameter(cmd, "@ProductVersion_ID", SqlDbType.Int, ProductVersionId);
            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.NVarChar, ServiceFamilyPn, 10);
            dw.CreateParameter(cmd, "@FCSDate", SqlDbType.DateTime, FCSDate);
            dw.CreateParameter(cmd, "@EndOfService", SqlDbType.DateTime, EndOfService);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static int UpdatePlatformName(string ProductVersionId, string ProductVersionName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateProductVersionName");

            dw.CreateParameter(cmd, "@ProductVersionID", SqlDbType.Int, ProductVersionId);
            dw.CreateParameter(cmd, "@DOTSName", SqlDbType.NVarChar, ProductVersionName, 30);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static DataTable GetServiceDates(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceSelectDesktopDates");

            dw.CreateParameter(cmd, "@ProductVersion_ID", SqlDbType.Int, ProductVersionId);

            return dw.ExecuteCommandTable(cmd);
        }

        public static int InsertProductFamily(string ProductFamilyName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spAddNewProductFamily");

            dw.CreateParameter(cmd, "@FamilyName", SqlDbType.NVarChar, ProductFamilyName, 100);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static int InsertSupplier(string SupplierName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spAddVendor");

            //dw.CreateParameter(cmd, "@Name", SqlDbType.NVarChar, SupplierName, 50);
            //return dw.ExecuteCommandNonQuery(cmd);
            dw.CreateParameter(cmd, "@Name", SqlDbType.NVarChar, SupplierName, 50);
            dw.CreateParameter(cmd, "@NewID", SqlDbType.Int, string.Empty, ParameterDirection.Output);
            dw.CreateParameter(cmd, "@NewSMTID", SqlDbType.Int, string.Empty, ParameterDirection.Output);

            //return dw.ExecuteCommandNonQuery(cmd);
            dw.ExecuteCommandNonQuery(cmd);

            return Convert.ToInt32(cmd.Parameters["@NewID"].Value.ToString());
        }



        public static DataTable GetSparekitsMaxEOSDate()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_ListSparekits_MaxEOSDate");

            return dw.ExecuteCommandTable(cmd);
        }



        public static DataTable GetServiceDesktopFamilyPartNumbers(string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_Desktop_FamilyPartNumbers");

            dw.CreateParameter(cmd, "@ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);

            return dw.ExecuteCommandTable(cmd);

        }

        public static DataTable GetSparekitBom(string SpareKitNumber)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_SelectPartBom");
            dw.CreateParameter(cmd, "@p_PartNumber", SqlDbType.NVarChar, SpareKitNumber, 10);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetCustomerLevel()
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_ListServiceCsrLevels");

            return dw.ExecuteCommandTable(cmd);
        }

        public static string GetDesktopsNewSpareKitMapId(string ProductBrandId, string SpareKitId, string ServiceFamilypn)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_Desktops_InsertServiceSpareKitMap");
            string serviceSpareKitMapId = string.Empty;

            dw.CreateParameter(cmd, "@p_ProductBrandId", SqlDbType.Int, ProductBrandId);
            dw.CreateParameter(cmd, "@p_SpareKitId", SqlDbType.Int, SpareKitId);
            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, ServiceFamilypn, 10);
            dw.CreateParameter(cmd, "@p_ServiceSpareKitMapId", SqlDbType.Int, string.Empty, ParameterDirection.Output);

            dw.ExecuteCommandNonQuery(cmd);

            return cmd.Parameters["@p_ServiceSpareKitMapId"].Value.ToString();
        }




        public static DataTable GetSparekitAV(string SpareKitNumber, string ServiceFamilypn)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_Service_Desktop_GetSparekitAv");

            dw.CreateParameter(cmd, "@SpsPartNumber", SqlDbType.NVarChar, SpareKitNumber, 10);
            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, ServiceFamilypn, 10);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetSparekit(string sSparekitNumber)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_Service_DesktopSparekitInNotebook");

            dw.CreateParameter(cmd, "@SparekitNumber", SqlDbType.NVarChar, sSparekitNumber, 10);

            return dw.ExecuteCommandTable(cmd);
        }

        public static int InsertDesktopSparekitAvNumber(string ServiceFamilypn, string SpsPartNumber, string SpsDescription, string SparekitCategoryID, string CustomerLevel, string Disposition, string Warranty, string LocalStockAdvice, string FirstServiceDate, string RslComments, bool GeoNA, bool GeoLA, bool GeoAPJ, bool GeoEMEA, string Supplier)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceInsertDesktopSparekit");

            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, ServiceFamilypn, 10);
            dw.CreateParameter(cmd, "@SpsPartNumber", SqlDbType.NVarChar, SpsPartNumber, 10);
            dw.CreateParameter(cmd, "@SpsDescription", SqlDbType.NVarChar, SpsDescription, 100);

            if (SparekitCategoryID != string.Empty) dw.CreateParameter(cmd, "@SparekitCategoryID", SqlDbType.Int, SparekitCategoryID);
            if (CustomerLevel != string.Empty) dw.CreateParameter(cmd, "@CustomerLevel", SqlDbType.Int, CustomerLevel);
            if (Disposition != string.Empty) dw.CreateParameter(cmd, "@Disposition", SqlDbType.Int, Disposition);
            if (Warranty != string.Empty) dw.CreateParameter(cmd, "@Warranty", SqlDbType.Char, Warranty, 1);
            if (LocalStockAdvice != string.Empty) dw.CreateParameter(cmd, "@LocalStockAdvice", SqlDbType.Int, LocalStockAdvice);
            if (RslComments != string.Empty) dw.CreateParameter(cmd, "@RslComments", SqlDbType.NVarChar, RslComments, 100);
            if (FirstServiceDate != string.Empty) dw.CreateParameter(cmd, "@FirstServiceDate", SqlDbType.DateTime, FirstServiceDate);

            if (GeoNA == true) { dw.CreateParameter(cmd, "@GeoNA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoNA", SqlDbType.Int, "0"); };
            if (GeoLA == true) { dw.CreateParameter(cmd, "@GeoLA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoLA", SqlDbType.Int, "0"); };
            if (GeoAPJ == true) { dw.CreateParameter(cmd, "@GeoAPJ", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoAPJ", SqlDbType.Int, "0"); };
            if (GeoEMEA == true) { dw.CreateParameter(cmd, "@GeoEMEA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoEMEA", SqlDbType.Int, "0"); };

            if (Supplier != string.Empty) dw.CreateParameter(cmd, "@Supplier", SqlDbType.NVarChar, Supplier, 50);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static int UpdateDesktopSparekit(string ServiceFamilypn, string SpsPartNumber, string SpsDescription, string SparekitCategoryID, string CustomerLevel, string Disposition, string Warranty, string LocalStockAdvice, bool GeoNA, bool GeoLA, bool GeoAPJ, bool GeoEMEA, string Supplier)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_UpdateDesktopSparekits");

            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.NVarChar, ServiceFamilypn, 10);
            dw.CreateParameter(cmd, "@SpsPartNumber", SqlDbType.NVarChar, SpsPartNumber, 10);
            dw.CreateParameter(cmd, "@SpsDescription", SqlDbType.NVarChar, SpsDescription, 100);

            if (SparekitCategoryID != string.Empty) dw.CreateParameter(cmd, "@SparekitCategoryID", SqlDbType.Int, SparekitCategoryID);
            if (CustomerLevel != string.Empty) dw.CreateParameter(cmd, "@CustomerLevel", SqlDbType.Int, CustomerLevel);
            if (Disposition != string.Empty) dw.CreateParameter(cmd, "@Disposition", SqlDbType.Int, Disposition);
            if (Warranty != string.Empty) dw.CreateParameter(cmd, "@Warranty", SqlDbType.Char, Warranty, 1);
            if (LocalStockAdvice != string.Empty) dw.CreateParameter(cmd, "@LocalStockAdvice", SqlDbType.Int, LocalStockAdvice);

            if (GeoNA == true) { dw.CreateParameter(cmd, "@GeoNA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoNA", SqlDbType.Int, "0"); };
            if (GeoLA == true) { dw.CreateParameter(cmd, "@GeoLA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoLA", SqlDbType.Int, "0"); };
            if (GeoAPJ == true) { dw.CreateParameter(cmd, "@GeoAPJ", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoAPJ", SqlDbType.Int, "0"); };
            if (GeoEMEA == true) { dw.CreateParameter(cmd, "@GeoEMEA", SqlDbType.Int, "1"); } else { dw.CreateParameter(cmd, "@GeoEMEA", SqlDbType.Int, "0"); };

            if (Supplier != string.Empty) dw.CreateParameter(cmd, "@Supplier", SqlDbType.NVarChar, Supplier, 50);

            return dw.ExecuteCommandNonQuery(cmd);

        }


        public static int DeleteServiceFamilySparekits(string ServiceFamilypn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_Delete_FamilySparekitsMapping");

            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.Char, ServiceFamilypn, 10);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public static int UpdateLinkAvSparekits(string ServiceFamilypn, string SpareKitNumber, string Category, string AVNumber)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_UpdateLinkAvSparekits");

            dw.CreateParameter(cmd, "@ServiceFamilypn", SqlDbType.Char, ServiceFamilypn, 10);
            dw.CreateParameter(cmd, "@SpareKitNumber", SqlDbType.NVarChar, SpareKitNumber, 10);
            dw.CreateParameter(cmd, "@CategoryID", SqlDbType.Int, Category);
            dw.CreateParameter(cmd, "@AVNumber", SqlDbType.NVarChar, AVNumber, 15);

            return dw.ExecuteCommandNonQuery(cmd);

        }





        public static DataTable GetServiceCommoditiesEOA(string Supplier, string Category, string HpPartNumber, string ServiceEOA)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_Service_Commodity_EOA");


            if (Supplier != string.Empty) dw.CreateParameter(cmd, "@p_Supplier", SqlDbType.NVarChar, Supplier, 150);
            if (Category != string.Empty) dw.CreateParameter(cmd, "@p_Category", SqlDbType.NVarChar, Category, 100);
            if (HpPartNumber != string.Empty) dw.CreateParameter(cmd, "@p_HPPartNumber", SqlDbType.NVarChar, HpPartNumber, 50);
            if (ServiceEOA != string.Empty) dw.CreateParameter(cmd, "@p_ServiceEOA", SqlDbType.Int, ServiceEOA);

            return dw.ExecuteCommandTable(cmd);


        }

        public static int UpdateServiceEOADate(string ExcaliburID, string sServiceEOADate)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Update_Service_EOADate");

            dw.CreateParameter(cmd, "@p_ExcaliburID", SqlDbType.Int, ExcaliburID);
            dw.CreateParameter(cmd, "@p_ServiceEOADate", SqlDbType.NVarChar, sServiceEOADate, 20);

            return dw.ExecuteCommandNonQuery(cmd);

        }

        public static DataTable GetSuppliers(string Category)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_ListSupplier");
            if (Category != string.Empty) dw.CreateParameter(cmd, "@p_Category", SqlDbType.NVarChar, Category, 100);


            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetUserIDInNTUserAndRole(string sNTUserName, string Role)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetUserIDInNTUserAndRole");

            dw.CreateParameter(cmd, "@p_User", SqlDbType.NVarChar, sNTUserName, 80);
            dw.CreateParameter(cmd, "@P_UserRole", SqlDbType.Int, Role);

            return dw.ExecuteCommandTable(cmd);
        }





        public static DataTable GetServiceList_ProductsVersionNames()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_ListProducts");

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_SPS_BOM(string SpareKitNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_Sps_BOM");

            if (SpareKitNumbers != string.Empty) dw.CreateParameter(cmd, "@SpareKitNumbers", SqlDbType.NVarChar, SpareKitNumbers, 2000);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_SPS_By_Category(string SpareCategoryIDs)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_SPS_By_Category");

            if (SpareCategoryIDs != string.Empty) dw.CreateParameter(cmd, "@SPSCategories", SqlDbType.NVarChar, SpareCategoryIDs, 500);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_SPS_From_AvNumbers(string AvNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_SPS_From_AvNumbers");
            if (AvNumbers != string.Empty) dw.CreateParameter(cmd, "@AvNumbers", SqlDbType.NVarChar, AvNumbers, 2000);
            return dw.ExecuteCommandTable(cmd);
        }


        public static DataTable GetServiceAVNumbersBaseUnitCategory(string AvNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_GetAVNumbersBaseUnitCategory");

            if (AvNumbers != string.Empty) dw.CreateParameter(cmd, "@AvNumbers", SqlDbType.NVarChar, AvNumbers, 250);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceAVNumbersKMAT(string AvNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Service_GetAVNumbersKMAT");

            if (AvNumbers != string.Empty) dw.CreateParameter(cmd, "@AvNumbers", SqlDbType.NVarChar, AvNumbers, 250);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_SKU_To_Sparekits(string SkuNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_SKU_To_SPS");

            if (SkuNumbers != string.Empty) dw.CreateParameter(cmd, "@SkuNumbers", SqlDbType.NVarChar, SkuNumbers, 2000);
            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_ProductVersion_To_Sku_Sparekits(string ProductVersionIds, string SkuNumbers)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_Product_To_Sku_Sparekits");
            if (ProductVersionIds != string.Empty) dw.CreateParameter(cmd, "@ProductVersionIds", SqlDbType.NVarChar, ProductVersionIds, 250);
            if (SkuNumbers != string.Empty) dw.CreateParameter(cmd, "@SkuNumbers", SqlDbType.NVarChar, SkuNumbers, 2000);
            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable GetServiceReport_UsedBy(string ProductVersionIds, string SKUs, string SpareKitNumbers, string SubAssemblies, string Components)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceReports_UsedBy");

            if (ProductVersionIds != string.Empty) dw.CreateParameter(cmd, "@ProductVersionIds", SqlDbType.NVarChar, ProductVersionIds, 500);
            if (SKUs != string.Empty) dw.CreateParameter(cmd, "@SKUNumbers", SqlDbType.NVarChar, SKUs, 2000);
            if (SpareKitNumbers != string.Empty) dw.CreateParameter(cmd, "@SpareKitNumbers", SqlDbType.NVarChar, SpareKitNumbers, 2000);
            if (SubAssemblies != string.Empty) dw.CreateParameter(cmd, "@SubAssemblies", SqlDbType.NVarChar, SubAssemblies, 2000);
            if (Components != string.Empty) dw.CreateParameter(cmd, "@Components", SqlDbType.NVarChar, Components, 2000);
            //  if (RegionsGeo != string.Empty) dw.CreateParameter(cmd, "@RegionGeo", SqlDbType.NVarChar, RegionsGeo, 50);


            return dw.ExecuteCommandTable(cmd);
        }



        public static DataTable GetServiceSpareKitBomDetails(string ServiceFamilyPn, string SpareKitNumber)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceProgramBomSS");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 100);
            dw.CreateParameter(cmd, "@p_SpareKitNumber", SqlDbType.Char, SpareKitNumber, 100);

            return dw.ExecuteCommandTable(cmd);
        }

        // DateTime SkuStartDate, DateTime SkuEndDate, DateTime SpsStartDate, DateTime SpsEndDate, 
        public static DataTable getServiceSpareKits(string SKUNumbers, string KMATs, string ProductNameId, string SpareKitCategories, string ServiceFamilyPartNumbers, string SpareKitPartNumbers, string OSSP, string ProductType, string GeoSKU, string GeoSpsNA, string GeoSpsLA, string GeoSpsAPJ, string GeoSpsEMEA, string SpsStartDate, string SpsEndDate, string MaxRows)
        {
            DataWrapper dw = new DataWrapper();

            SqlCommand cmd = dw.CreateCommand("usp_ServiceGetSpareKits");

            if (SKUNumbers != string.Empty) dw.CreateParameter(cmd, "@p_SKUNumber", SqlDbType.NVarChar, SKUNumbers, 250);
            if (KMATs != string.Empty) dw.CreateParameter(cmd, "@p_KMAT", SqlDbType.NVarChar, KMATs, 250);
            if (ProductNameId != string.Empty) dw.CreateParameter(cmd, "@p_ProductNameId", SqlDbType.NVarChar, ProductNameId, 100);
            if (SpareKitCategories != string.Empty) dw.CreateParameter(cmd, "@p_SpareKitCategories", SqlDbType.NVarChar, SpareKitCategories, 100);
            if (ServiceFamilyPartNumbers != string.Empty) dw.CreateParameter(cmd, "@p_ServiceFamilyPartNumber", SqlDbType.NVarChar, ServiceFamilyPartNumbers, 250);
            if (SpareKitPartNumbers != string.Empty) dw.CreateParameter(cmd, "@p_SpareKitPartNumber", SqlDbType.NVarChar, SpareKitPartNumbers, 250);
            if (OSSP != string.Empty) dw.CreateParameter(cmd, "@p_OSSP", SqlDbType.VarChar, OSSP, 100);
            if (ProductType != "0") dw.CreateParameter(cmd, "@p_ProductType", SqlDbType.Int, ProductType);
            if (GeoSKU != string.Empty) dw.CreateParameter(cmd, "@P_GeoSKU", SqlDbType.NVarChar, GeoSKU, 10);
            if (GeoSpsNA != string.Empty) dw.CreateParameter(cmd, "@p_GeoSPSNA", SqlDbType.Int, GeoSpsNA);
            if (GeoSpsLA != string.Empty) dw.CreateParameter(cmd, "@p_GeoSPSLA", SqlDbType.Int, GeoSpsLA);
            if (GeoSpsAPJ != string.Empty) dw.CreateParameter(cmd, "@p_GeoSPSAPJ", SqlDbType.Int, GeoSpsAPJ);
            if (GeoSpsEMEA != string.Empty) dw.CreateParameter(cmd, "@p_GeoSPSEMEA", SqlDbType.VarChar, GeoSpsEMEA);
            //if (SkuStartDate != string.Empty) dw.CreateParameter(cmd, "@p_SkuStartDate", SqlDbType.VarChar, SkuStartDate, 50);
            //if (SkuEndDate != string.Empty) dw.CreateParameter(cmd, "@p_SkuEndDate", SqlDbType.VarChar, SkuEndDate, 50);
            if (SpsStartDate != string.Empty) dw.CreateParameter(cmd, "@p_SpsStartDate", SqlDbType.DateTime, SpsStartDate);
            if (SpsEndDate != string.Empty) dw.CreateParameter(cmd, "@p_SpsEndDate", SqlDbType.DateTime, SpsEndDate);

            if (MaxRows != string.Empty) dw.CreateParameter(cmd, "@p_MaxRows", SqlDbType.Int, MaxRows);

            return dw.ExecuteCommandTable(cmd);
        }


        public static DataTable getServiceSpareKit(string SpareKitNumber)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetServiceSpareKit");

            dw.CreateParameter(cmd, "@p_SpareKitNumber", SqlDbType.Char, SpareKitNumber, 10);

            return dw.ExecuteCommandTable(cmd);
        }




        public static DataTable getAdvancedServiceBomReport(string SKUNumber, string KMAT, string MaxRows)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ServiceKmatSkuToSkuSPS");

            if (SKUNumber != string.Empty) dw.CreateParameter(cmd, "@p_SKU", SqlDbType.NVarChar, SKUNumber, 2000);
            if (KMAT != string.Empty) dw.CreateParameter(cmd, "@p_KMAT", SqlDbType.NVarChar, KMAT, 2000);
            if (MaxRows != string.Empty) dw.CreateParameter(cmd, "@p_MaxRows", SqlDbType.Int, MaxRows);

            return dw.ExecuteCommandTable(cmd);
        }


        public static DataTable getAvsDeletedMappedToSPS()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListAvsDeleted");

            return dw.ExecuteCommandTable(cmd);
        }


        public static int UpdateDeletedAvMappedToSPS(string AvDetailIDs, string UserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateServiceSpareKitMapAV");

            dw.CreateParameter(cmd, "@p_AvDetailIds", SqlDbType.NVarChar, AvDetailIDs, 150);
            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, UserID);

            return dw.ExecuteCommandNonQuery(cmd);
        }


        public static DataTable getAVsNotMappedToSPS(string UserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceUnmappedAvs");
            dw.CreateParameter(cmd, "@p_UserId", SqlDbType.Int, UserID);


            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable getSPSNotMappedToAV(string UserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceSpareKitsNotMapped");

            dw.CreateParameter(cmd, "@p_UserId", SqlDbType.Int, UserID);

            return dw.ExecuteCommandTable(cmd);
        }




    }
}
