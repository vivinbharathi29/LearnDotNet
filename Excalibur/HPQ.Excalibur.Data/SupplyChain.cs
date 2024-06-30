using HPQ.Data;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{
    [DataObjectAttribute()]
    public static class SupplyChain
    {

        public static void InsertAvRegionalDates(string strGeoID, string strAvDetail_ProductBrandID,
                                  string strRegionalCPLBlindDate, string strRegionalRasDiscDate, string strStatus,
                                  string strProductBrandID, string strAvFeatureCategoryID)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_AVRegionalDatesInsert");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@AvDetail_ProductBrandID", SqlDbType.Int, strAvDetail_ProductBrandID);
            dw.CreateParameter(cmd, "@RegionalCPLBlindDate", SqlDbType.Date, strRegionalCPLBlindDate);
            dw.CreateParameter(cmd, "@RegionalRasDiscDate", SqlDbType.Date, strRegionalRasDiscDate);
            dw.CreateParameter(cmd, "@Status", SqlDbType.Int, strStatus);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, strProductBrandID, 8000);
            dw.CreateParameter(cmd, "@AvFeatureCategoryID", SqlDbType.Int, strAvFeatureCategoryID);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static DataTable SelectScmDetail_RegionAndPlatformsView(string ProductVersionID, string ProductBrandID, string Categories, string GeoID)
        {
            //, string GeoID)
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScmDetail_RegionAndPlatformsView");
            cmd.CommandTimeout = 120;

            //GeoID = 1

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, ProductVersionID, 8000);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, ProductBrandID, 8000);
            if (Categories != "-1") //If the Categories variable is empty then don't send it to the sproc parameter collection.
            {
                dw.CreateParameter(cmd, "@p_Categories", SqlDbType.VarChar, Categories, 500);
            }
            dw.CreateParameter(cmd, "@p_GeoID", SqlDbType.Int, GeoID);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable SelectScmDetail_RegionAndPlatformsView_PlantView(string ProductVersionID, string ProductBrandID, string Categories, string GeoID, string strPlants)
        {
            //, string GeoID)
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScmDetail_RegionAndPlatformsView_PlantView");
            cmd.CommandTimeout = 120;

            //GeoID = 1

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, ProductVersionID, 8000);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, ProductBrandID, 8000);
            if (Categories != "-1") //If the Categories variable is empty then don't send it to the sproc parameter collection.
            {
                dw.CreateParameter(cmd, "@p_Categories", SqlDbType.VarChar, Categories, 500);
            }
            dw.CreateParameter(cmd, "@p_GeoID", SqlDbType.Int, GeoID);
            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.VarChar, strPlants, 8000);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable SelectScmDetail_RegionAndPlatformsView_MktCampView(string ProductVersionID, string ProductBrandID, string Categories, string GeoID, string strPlants, string strMktCampID)
        {
            //, string GeoID)
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScmDetail_RegionAndPlatformsView_MktCampView");
            cmd.CommandTimeout = 120;

            //GeoID = 1

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, ProductVersionID, 8000);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, ProductBrandID, 8000);
            if (Categories != "-1") //If the Categories variable is empty then don't send it to the sproc parameter collection.
            {
                dw.CreateParameter(cmd, "@p_Categories", SqlDbType.VarChar, Categories, 500);
            }
            dw.CreateParameter(cmd, "@p_GeoID", SqlDbType.Int, GeoID);
            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.VarChar, strPlants, 8000);
            //dw.CreateParameter(cmd, "@MktCampID", SqlDbType.VarChar, strMktCampID, 8000);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable SelectScmDetail_RegionAndPlatformsView_WithOutCats(string ProductVersionID, string ProductBrandID, string GeoID)
        {
            DataWrapper dw = new DataWrapper();
            //SqlCommand cmd = dw.CreateCommand("usp_SelectScmDetail_RegionAndPlatformsView");
            SqlCommand cmd = dw.CreateCommand("usp_SelectScmDetail_RegionAndPlatformsView_ByGeoID");
            cmd.CommandTimeout = 120;

            //GeoID = 1

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, ProductVersionID, 8000);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, ProductBrandID, 8000);
            dw.CreateParameter(cmd, "@p_GeoID", SqlDbType.Int, GeoID);

            return dw.ExecuteCommandTable(cmd);
        }

        public static DataTable SelectProductBrands(string BusinessID, string ProductStatusID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductBrands");

            dw.CreateParameter(cmd, "@p_BusinessID", SqlDbType.Int, BusinessID);
            dw.CreateParameter(cmd, "@p_ProductStatusID", SqlDbType.Int, ProductStatusID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable SelectRCTOPlants_ByGeoID(string strGeoID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_RCTOPlantsSelect_ByGeoID");

            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable SelectProductBrandNames_ByProductVersionIDAndProductBrandID(string strProdVerID, string strProdBrandID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ProductBrandNames_ByProductVersionIDAndProductBrandIDSelect");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, strProdVerID, 8000);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, strProdBrandID, 8000);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static void UpdateAvRegionalDates(string strAvRegionalDatesID, string strGEOID,
                                                    string strAvDetail_ProductBrandID, string strRegionalCPLBlindDate,
                                                    string strRegionalRasDiscDate, string strStatus,
                                                    string strp_ProductBrandID, string strAvFeatureCategoryID,
                                                    string strCheckedRec, string strRecSelected)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_AvRegionalDatesUpdate");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@AvRegionalDatesID", SqlDbType.Int, strAvRegionalDatesID);
            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGEOID);
            dw.CreateParameter(cmd, "@AvDetail_ProductBrandID", SqlDbType.Int, strAvDetail_ProductBrandID);
            dw.CreateParameter(cmd, "@RegionalCPLBlindDate", SqlDbType.Date, strRegionalCPLBlindDate);
            dw.CreateParameter(cmd, "@RegionalRasDiscDate", SqlDbType.Date, strRegionalRasDiscDate);
            dw.CreateParameter(cmd, "@Status", SqlDbType.Int, strStatus);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.VarChar, strp_ProductBrandID, 8000);
            dw.CreateParameter(cmd, "@AvFeatureCategoryID", SqlDbType.Int, strAvFeatureCategoryID);
            dw.CreateParameter(cmd, "@CheckedRec", SqlDbType.Int, strCheckedRec);
            dw.CreateParameter(cmd, "@RecSelected", SqlDbType.Int, strRecSelected);

            dw.ExecuteCommandNonQuery(cmd);
        }


        public static DataTable Regions_Select_All()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Regions_Select_All");
            cmd.CommandTimeout = 120;

            return dw.ExecuteCommandTable(cmd);
        }


        public static void InsertAvPlantDates_OneRecAtATtime(string strRCTOPlantsID, string strAvRegionalDatesID,
                                  string strGeoID, string strPlantStartDate, string strPlantEndDate)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_RCTOPlants_AVDetailProductBrandInsert");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.Int, strRCTOPlantsID);
            dw.CreateParameter(cmd, "@AvRegionalDatesID", SqlDbType.Int, strAvRegionalDatesID);
            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@PlantStartDate", SqlDbType.Date, strPlantStartDate);
            dw.CreateParameter(cmd, "@PlantEndDate", SqlDbType.Date, strPlantEndDate);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void UpdateAvPlantDates(string strRCTOPlantsID, string strAvRegionalDatesID,
                                  string strGeoID, string strPlantStartDate, string strPlantEndDate,
                                  string strRCTOPlants_AVDetailProductBrand, string strCheckedRec, string strRecSelected)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_RCTOPlants_AVDetailProductBrandUpdate");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@RCTOPlants_AVDetailProductBrand", SqlDbType.Int, strRCTOPlants_AVDetailProductBrand);
            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.Int, strRCTOPlantsID);
            dw.CreateParameter(cmd, "@AvRegionalDatesID", SqlDbType.Int, strAvRegionalDatesID);
            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@PlantStartDate", SqlDbType.Date, strPlantStartDate);
            dw.CreateParameter(cmd, "@PlantEndDate", SqlDbType.Date, strPlantEndDate);
            dw.CreateParameter(cmd, "@CheckedRec", SqlDbType.Int, strCheckedRec);
            dw.CreateParameter(cmd, "@RecSelected", SqlDbType.Int, strRecSelected);

            dw.ExecuteCommandNonQuery(cmd);
        }



        public static DataTable ListAllMktCamps_ByRegion(string strActivebit, string strGeoID)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaignsSelectAll_ByRegion");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@Active", SqlDbType.Bit, strActivebit);
            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);

            //dw.ExecuteCommandNonQuery(cmd);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static DataTable MktCampaigns_GetAllDataForSingleRec(string strMktCampID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaigns_GetAllDataForSingleRec_Select");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampID);

            //dw.ExecuteCommandNonQuery(cmd);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public static void MktCampaigns_Insert(string strGeoID, string strCampaignName, string strStartDate, string strEndDate, string strPlantID, string strPlantName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaignsInsert");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@CampaignName", SqlDbType.VarChar, strCampaignName, 100);
            dw.CreateParameter(cmd, "@StartDate", SqlDbType.Date, strStartDate);
            dw.CreateParameter(cmd, "@EndDate", SqlDbType.Date, strEndDate);
            dw.CreateParameter(cmd, "@PlantID", SqlDbType.VarChar, strPlantID, 800);
            dw.CreateParameter(cmd, "@PlantName", SqlDbType.VarChar, strPlantName, 800);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void MktCampaigns_Update(string strMktCampaigns, string strGeoID, string strCampaignName,
                                               string strStartDate, string strEndDate, string strActive, string strPlantID, string strPlantName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaignsUpdate");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampaigns);
            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@CampaignName", SqlDbType.VarChar, strCampaignName, 100);
            dw.CreateParameter(cmd, "@StartDate", SqlDbType.Date, strStartDate);
            dw.CreateParameter(cmd, "@EndDate", SqlDbType.Date, strEndDate);
            dw.CreateParameter(cmd, "@Active", SqlDbType.Bit, strActive);
            dw.CreateParameter(cmd, "@PlantID", SqlDbType.VarChar, strPlantID, 800);
            dw.CreateParameter(cmd, "@PlantName", SqlDbType.VarChar, strPlantName, 800);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void MktCampaigns_AVDetailProductBrandInsert(string strMktCampID, string strAVDPBID, string strRCTOPlantsID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaigns_AVDetailProductBrandInsert");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampID);
            dw.CreateParameter(cmd, "@AvDetailProductBrandID", SqlDbType.Int, strAVDPBID);
            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.Int, strRCTOPlantsID);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void MktCampaigns_AVDetailProductBrandDeleteByMktCampOnly(string strMktCampiagns)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaigns_AVDetailProductBrandDelete_ByMktCampaignsID");

            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampiagns);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void MktCampaigns_AVDetailProductBrandUpdate(string strMktCampaigns_AVDetailProductBrandID,
                                    string strMktCampaignsID, string strAvDetailProductBrandID, string strRCTOPlantsID,
                                    string strCheckedRec, string strRecSelected)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaigns_AVDetailProductBrandUpdate");
            //int returnVal = 0;

            dw.CreateParameter(cmd, "@MktCampaigns_AVDetailProductBrandID", SqlDbType.Int, strMktCampaigns_AVDetailProductBrandID);
            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampaignsID);
            dw.CreateParameter(cmd, "@AvDetailProductBrandID", SqlDbType.Int, strAvDetailProductBrandID);
            dw.CreateParameter(cmd, "@RCTOPlantsID", SqlDbType.Int, strRCTOPlantsID);
            dw.CreateParameter(cmd, "@CheckedRec", SqlDbType.Bit, strCheckedRec);
            dw.CreateParameter(cmd, "@RecSelected", SqlDbType.Bit, strRecSelected);

            dw.ExecuteCommandNonQuery(cmd);
        }

        public static void MktCampaigns_Delete(string strMktCampiagns)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaignsDelete");

            dw.CreateParameter(cmd, "@MktCampaignsID", SqlDbType.Int, strMktCampiagns);

            dw.ExecuteCommandNonQuery(cmd);
        }

        //This function looks for duplicate Mkt Campaign records.
        public static bool DupRecCheck(string strGeoID, string strCampaignName,
                                               string strStartDate, string strEndDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampDupRecCheck");

            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@CampaignName", SqlDbType.VarChar, strCampaignName, 100);
            dw.CreateParameter(cmd, "@StartDate", SqlDbType.Date, strStartDate);
            dw.CreateParameter(cmd, "@EndDate", SqlDbType.Date, strEndDate);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            bool bolRecFound = false;

            if (dt.Rows.Count == 0)
                bolRecFound = false;
            else if (dt.Rows[0]["Rec"].ToString() == "0")
                bolRecFound = false;
            else
                bolRecFound = true;

            return bolRecFound;
        }

        //This function looks for duplicate Mkt Campaign records.
        public static int MktCampaign_ReturnMktCampID(string strGeoID, string strCampaignName,
                                               string strStartDate, string strEndDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_MktCampaignsSelectNewlyCreatedRecByValuesAndReturnID");

            dw.CreateParameter(cmd, "@GeoID", SqlDbType.Int, strGeoID);
            dw.CreateParameter(cmd, "@CampaignName", SqlDbType.VarChar, strCampaignName, 100);
            dw.CreateParameter(cmd, "@StartDate", SqlDbType.Date, strStartDate);
            dw.CreateParameter(cmd, "@EndDate", SqlDbType.Date, strEndDate);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            int intMktCampID = 0;

            if (dt.Rows.Count == 0)
                intMktCampID = -1;
            else if (dt.Rows[0]["MktCampaignsID"].ToString() == "0")
                intMktCampID = -1;
            else
                intMktCampID = Convert.ToInt16(dt.Rows[0]["MktCampaignsID"].ToString());

            return intMktCampID;
        }

    }
}
