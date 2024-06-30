using HPQ.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace HPQ.Excalibur
{

    /*#########################################################
    #########################################################
    #########################################################
    #########################################################

        ATTENTION!
        this file-salad also called Excalibur has some aspx files WITH THE CODE IN THE ASPX PAGE in a <script section>
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
    public class Data
    {
        public DataTable SelectSkuImageStatus(string ProductVersionID, string OsFamilyID,
          string SigVerifyComplete, string CheckLogo6Complete, string WmiComplete)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectSkuImageStatus");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@p_OsFamilyID", SqlDbType.Int, OsFamilyID);
            dw.CreateParameter(cmd, "@p_SigVerifyComplete", SqlDbType.Bit, SigVerifyComplete);
            dw.CreateParameter(cmd, "@p_CheckLogo6Complete", SqlDbType.Bit, CheckLogo6Complete);
            dw.CreateParameter(cmd, "@p_WmiComplete", SqlDbType.Bit, WmiComplete);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable SelectImagesWithUnsignedDrivers(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectImagesWithUnsignedDrivers");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable SelectProductSkusWithoutWhql(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductSkusWithoutWhql");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable SelectProductWhql(string ProductWhqlID, string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductWhql");

            dw.CreateParameter(cmd, "@p_ProductWhqlID", SqlDbType.Int, ProductWhqlID);
            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable SelectAvDetail(string ProductVersionID, string ProductBrandID, string AvDetailID,
       string AvCategoryID, string AvNo, string GpgDescription, string UPC, string Status, string KMAT)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvDetail");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);
            dw.CreateParameter(cmd, "@p_AvDetailID", SqlDbType.Int, AvDetailID);
            dw.CreateParameter(cmd, "@p_AvCategoryID", SqlDbType.Int, AvCategoryID);
            dw.CreateParameter(cmd, "@p_AvNo", SqlDbType.VarChar, AvNo, 10);
            dw.CreateParameter(cmd, "@p_GpgDescription", SqlDbType.VarChar, GpgDescription, 50);
            dw.CreateParameter(cmd, "@p_UPC", SqlDbType.VarChar, UPC, 12);
            dw.CreateParameter(cmd, "@p_Status", SqlDbType.Char, Status, 1);
            dw.CreateParameter(cmd, "@p_KMAT", SqlDbType.Char, KMAT, 6);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListProductBrandBaseUnitsAvDetail(string ProductBrandID)
        {
            DataTable dtCommercial = SelectAvDetail(string.Empty, ProductBrandID, string.Empty, "1", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            DataTable dtConsumer = SelectAvDetail(string.Empty, ProductBrandID, string.Empty, "86", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

            DataTable dtBaseUnits = dtCommercial.Clone();
            foreach (DataRow dRow in dtCommercial.Rows)
            {
                dtBaseUnits.ImportRow(dRow);
            }
            foreach (DataRow dRow in dtConsumer.Rows)
            {
                dtBaseUnits.ImportRow(dRow);
            }

            return dtBaseUnits;
        }
        public DataTable ListProductBrandCpuAvDetail(string ProductBrandID)
        {
            DataTable dtCommercial = SelectAvDetail(string.Empty, ProductBrandID, string.Empty, "4", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            DataTable dtConsumer = SelectAvDetail(string.Empty, ProductBrandID, string.Empty, "99", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

            DataTable dtProcessors = dtCommercial.Clone();
            foreach (DataRow dRow in dtCommercial.Rows)
            {
                dtProcessors.ImportRow(dRow);
            }
            foreach (DataRow dRow in dtConsumer.Rows)
            {
                dtProcessors.ImportRow(dRow);
            }
            return dtProcessors;
        }
        public DataTable ListOsFamilies()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListOsFamilies");
            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListBrandSeries(string ProductBrandID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListBrandSeries ");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListProductBrands(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListProductBrands");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListWhqlModels(string ProductWhqlID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListWhqlModels");

            dw.CreateParameter(cmd, "@p_ProductWhqlID", SqlDbType.Int, ProductWhqlID);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable SelectProductWHQLWithoutBootVis(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductWHQLWithoutBootVis");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListImagesWithOutWhqlStatus(string ImageSkuNumber)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListImagesWithOutWhqlStatus");

            dw.CreateParameter(cmd, "@p_ImageSkuNumber", SqlDbType.VarChar, ImageSkuNumber, 10);

            return dw.ExecuteCommandTable(cmd);
        }
        public DataTable ListImagesWithWhqlStatus(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListImagesWithWhqlStatus");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable SelectSpareKits(string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectSpareKits");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable ListSpdmUsers()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListSpdmUsers");

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable SelectAvsNotInKmatBom(string ProductBrandId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvsNotInKmatBom");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandId);

            return dw.ExecuteCommandTable(cmd);
        }

        public Int32 GetWhqlAvDetailCount(string WhqlID, string BaseUnitID, string ProcessorID, string OsFamilyID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetWHQLAvDetailCount");

            dw.CreateParameter(cmd, "@p_WHQLID", SqlDbType.Int, WhqlID);
            dw.CreateParameter(cmd, "@p_BUID", SqlDbType.Int, BaseUnitID);
            dw.CreateParameter(cmd, "@p_CPUID", SqlDbType.Int, ProcessorID);
            dw.CreateParameter(cmd, "@p_OsFamilyID", SqlDbType.Int, OsFamilyID);

            return Convert.ToInt32(dw.ExecuteCommandScalar(cmd));
        }

        public DataTable SelectWhqlSubmissions(string ProductWhqlID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectWHQLSubmissions");

            dw.CreateParameter(cmd, "@p_ProductWhqlID", SqlDbType.Int, ProductWhqlID);

            return dw.ExecuteCommandTable(cmd);
        }

        public Int32 GetProductWhqlID(string SubmissionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetProductWHQLID");

            dw.CreateParameter(cmd, "@p_SubmissionID", SqlDbType.VarChar, SubmissionID, 50);

            return Convert.ToInt32(dw.ExecuteCommandScalar(cmd));
        }

        public Int32 InsertWhqlAvDetail(string ProductWhqlID, string BaseUnitID, string ProcessorID, string OsFamilyID, string ProductBrandID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertWHQLAvDetail");

            dw.CreateParameter(cmd, "@p_WHQLID", SqlDbType.Int, ProductWhqlID);
            dw.CreateParameter(cmd, "@p_BUID", SqlDbType.Int, BaseUnitID);
            dw.CreateParameter(cmd, "@p_CPUID", SqlDbType.Int, ProcessorID);
            dw.CreateParameter(cmd, "@p_OsFamilyID", SqlDbType.Int, OsFamilyID);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public Int32 InsertProductWhql(string SubmissionID, string SubmissionDt, string WhqlDt, string Status,
    string ProductVersionID, string Location, string ReleaseDt, string LogoDisplayed, string Milestone3)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertProductWHQL");

            dw.CreateParameter(cmd, "@p_SubmissionID", SqlDbType.VarChar, SubmissionID, 50);
            dw.CreateParameter(cmd, "@p_SubmissionDt", SqlDbType.DateTime, SubmissionDt);
            dw.CreateParameter(cmd, "@p_WhqlDt", SqlDbType.DateTime, WhqlDt);
            dw.CreateParameter(cmd, "@p_Status", SqlDbType.Int, Status);
            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@p_Location", SqlDbType.VarChar, Location, 200);
            dw.CreateParameter(cmd, "@p_DateReleased", SqlDbType.DateTime, ReleaseDt);
            dw.CreateParameter(cmd, "@p_LogoDisplayed", SqlDbType.Bit, LogoDisplayed);
            dw.CreateParameter(cmd, "@p_Milestone3", SqlDbType.Bit, Milestone3);
            dw.CreateParameter(cmd, "@p_ProductWhqlID", SqlDbType.Int, "", ParameterDirection.Output);

            dw.ExecuteCommandNonQuery(cmd);

            return Convert.ToInt32(cmd.Parameters["@p_ProductWhqlID"].Value);
        }

        public int InsertProductWhqlSeries(string ProductWhqlID, string SeriesID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertProductWhqlSeries");

            dw.CreateParameter(cmd, "@p_ProductWhqlID", SqlDbType.Int, ProductWhqlID);
            dw.CreateParameter(cmd, "@p_SeriesID", SqlDbType.Int, SeriesID);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public int LeverageImageWhqlStatus(string ImageSkuNo, string LeveragedSkuNo)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SetImageSigveryfyLeveragedStatus");

            dw.CreateParameter(cmd, "@p_ImageSkuNo", SqlDbType.VarChar, ImageSkuNo, 10);
            dw.CreateParameter(cmd, "@p_LeveragedSkuNo", SqlDbType.VarChar, LeveragedSkuNo, 10);

            return dw.ExecuteCommandNonQuery(cmd);
        }
        public int RollbackScmPublish(string ProductBrandId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_RollbackScmPublish");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandId);

            return dw.ExecuteCommandNonQuery(cmd);
        }
        public DataTable ListBrands4Product(int ProductID, int SelectedOnly)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListBrands4Product");

            dw.CreateParameter(cmd, "@ProductID", SqlDbType.Int, ProductID.ToString());
            dw.CreateParameter(cmd, "@SelectedOnly", SqlDbType.TinyInt, SelectedOnly.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectAvHistory(string ProductBrandID, string AvHistoryID, string ShowOnSCM, string ShowOnPM, string ShowAll, string ShowDays)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvHistory");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);
            dw.CreateParameter(cmd, "@p_AvHistoryID", SqlDbType.Int, AvHistoryID);
            dw.CreateParameter(cmd, "@p_ShowOnSCM", SqlDbType.Bit, ShowOnSCM);
            dw.CreateParameter(cmd, "@p_ShowOnPM", SqlDbType.Bit, ShowOnPM);
            dw.CreateParameter(cmd, "@p_ShowAll", SqlDbType.Bit, ShowAll);
            dw.CreateParameter(cmd, "@p_ShowDays", SqlDbType.Int, ShowDays);

            return dw.ExecuteCommandTable(cmd);
        }

        public int SetAvHistoryShowOnScmStatus(string AvHistoryId, string ShowOnScm, string LastUpdUser)
        {
            return SetAvHistoryShowOnScmPmStatus(AvHistoryId, ShowOnScm, string.Empty, LastUpdUser);
        }

        public int SetAvHistoryShowOnPmStatus(string AvHistoryId, string ShowOnPm, string LastUpdUser)
        {
            return SetAvHistoryShowOnScmPmStatus(AvHistoryId, string.Empty, ShowOnPm, LastUpdUser);
        }

        public int SetAvHistoryShowOnScmPmStatus(string AvHistoryId, string ShowOnScm, string ShowOnPm, string LastUpdUser)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SetAvHistoryShowOnScmStatus");

            dw.CreateParameter(cmd, "@p_AvHistoryID", SqlDbType.Int, AvHistoryId);
            dw.CreateParameter(cmd, "@p_ShowOnScm", SqlDbType.Bit, ShowOnScm);
            dw.CreateParameter(cmd, "@p_ShowOnPM", SqlDbType.Bit, ShowOnPm);
            dw.CreateParameter(cmd, "@p_LastUpdUser", SqlDbType.VarChar, LastUpdUser, 50);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public DataTable GetProgramCoordinatorStatus(string EmployeeID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetProgramCoordinatorStatus");

            dw.CreateParameter(cmd, "@p_EmployeeID", SqlDbType.Int, EmployeeID);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable ListFeatureCategoy_Localized()
        {
            //, string GeoID)
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListLocalizedFeatureCategory");
            cmd.CommandTimeout = 120;

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable SelectSpbDetails(string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectSpbDetails");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);

            return dw.ExecuteCommandTable(cmd);
        }

        [System.ComponentModel.DataObjectMethodAttribute(System.ComponentModel.DataObjectMethodType.Update, true)]
        public int UpdateServiceFamilyDetails(string ServiceFamilyPn, string SpdmContactID, string GplmContactID, bool Active,
            string SharePointPath, string SharedDrivePath, string SelfRepairDoc, bool AutoPublishRsl, int BusinessUnit)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateServiceFamilyDetails");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);
            dw.CreateParameter(cmd, "@p_SpdmContactId", SqlDbType.Int, SpdmContactID);
            dw.CreateParameter(cmd, "@p_GplmContactId", SqlDbType.Int, GplmContactID);
            dw.CreateParameter(cmd, "@p_Active", SqlDbType.Bit, Active.ToString());
            dw.CreateParameter(cmd, "@p_SharePointPath", SqlDbType.VarChar, SharePointPath);
            dw.CreateParameter(cmd, "@p_SharedDrivePath", SqlDbType.VarChar, SharedDrivePath);
            dw.CreateParameter(cmd, "@p_SelfRepairDoc", SqlDbType.VarChar, SelfRepairDoc);
            dw.CreateParameter(cmd, "@p_AutoPublishRsl", SqlDbType.Bit, AutoPublishRsl.ToString());
            dw.CreateParameter(cmd, "@p_BusinessUnit", SqlDbType.Int, BusinessUnit.ToString());

            return dw.ExecuteCommandNonQuery(cmd);

        }


        public int UpdateServiceSpareDetail(string ServiceFamilyPn, string HpPartNo, string SpareCategory,
            bool OsspOrderable, string OdmPartNo, string OdmPartDesc, string OdmBulkPartNo, string OdmProdMoq,
            string OdmPostProdMoq, string Comments, string Supplier)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateServiceSpareDetailByUI");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);
            dw.CreateParameter(cmd, "@p_HpPartNo", SqlDbType.Char, HpPartNo, 10);
            dw.CreateParameter(cmd, "@p_SpareCategory", SqlDbType.VarChar, SpareCategory, 50);
            dw.CreateParameter(cmd, "@p_OsspOrderable", SqlDbType.Bit, OsspOrderable.ToString());
            dw.CreateParameter(cmd, "@p_OdmPartNo", SqlDbType.VarChar, OdmPartNo, 50);
            dw.CreateParameter(cmd, "@p_OdmPartDesc", SqlDbType.VarChar, OdmPartDesc, 100);
            dw.CreateParameter(cmd, "@p_OdmBulkPartNo", SqlDbType.VarChar, OdmBulkPartNo, 50);
            dw.CreateParameter(cmd, "@p_OdmProdMoq", SqlDbType.VarChar, OdmProdMoq, 50);
            dw.CreateParameter(cmd, "@p_OdmPostProdMoq", SqlDbType.VarChar, OdmPostProdMoq, 50);
            dw.CreateParameter(cmd, "@p_Comments", SqlDbType.VarChar, Comments, 800);
            dw.CreateParameter(cmd, "@p_Supplier", SqlDbType.VarChar, Supplier, 50);

            return dw.ExecuteCommandNonQuery(cmd);
        }



        public int SetServiceFamilyPn(string ProductVersionId, string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SetServiceFamilyPn");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);
            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);

            return dw.ExecuteCommandNonQuery(cmd);
        }



        public string GetServiceFamilyPn(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_GetServiceFamilyPn");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);

            return dw.ExecuteCommandScalar(cmd).ToString();
        }


        public DataTable ListSpbPublishDates(string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListSpbPublishDates");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);

            return dw.ExecuteCommandTable(cmd);
        }



        public DataTable SelectServiceSpareDetails(string ServiceFamilyPn, string HpPartNo)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceSpareDetails");

            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.Char, ServiceFamilyPn, 10);
            dw.CreateParameter(cmd, "@p_HpPartNo", SqlDbType.Char, HpPartNo, 10);

            return dw.ExecuteCommandTable(cmd);

        }



        public DataTable ListSpdms()
        {
            return ListSpdms(string.Empty);
        }
        public DataTable ListSpdms(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListSpdms");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable ListGplms()
        {
            return ListGplms(string.Empty);
        }
        public DataTable ListGplms(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListGplms");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable ListSvcManagers()
        {
            return ListSvcManagers(string.Empty);
        }
        public DataTable ListSvcManagers(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListSvcManagers");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable ListServiceSpareCategories()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListServiceSpareCategories");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListDeliverablesCategories()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListDeliverablesCategories");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable GetProductVersion(string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetProductVersion");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, PVID);

            return dw.ExecuteCommandTable(cmd);
        }

        public string GetProductVersionIdsByGroupIds(string CommaSeparatedGIds)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetProgramTree");

            dw.CreateParameter(cmd, "@p_ProductGroupIDs", SqlDbType.VarChar, CommaSeparatedGIds);

            DataTable dtProductVerion = dw.ExecuteCommandTable(cmd);
            List<string> prodIdList = new List<string>();
            foreach (DataRow dRow in dtProductVerion.Rows)
            {
                prodIdList.Add(dRow["ProdID"].ToString());
            }
            return string.Join(",", prodIdList.ToArray());
        }

        public DataTable ListSystemTeam(string ProductVersionID)
        {
            return ListSystemTeam(ProductVersionID, string.Empty, string.Empty);
        }
        public DataTable ListSystemTeam(string ProductVersionID, string AllEmployees, string AddBios)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListSystemTeam");

            dw.CreateParameter(cmd, "@ProdID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@AllEmployees", SqlDbType.Bit, AllEmployees);
            dw.CreateParameter(cmd, "@AddBios", SqlDbType.Bit, AddBios);

            return dw.ExecuteCommandTable(cmd);
        }


        public DataTable GetUserInfo(string UserName, string Domain)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetUserInfo");

            dw.CreateParameter(cmd, "@UserName", SqlDbType.VarChar, UserName, 30);
            dw.CreateParameter(cmd, "@Domain", SqlDbType.VarChar, Domain, 30);

            return dw.ExecuteCommandTable(cmd);
        }



        public DataTable GetUserRoles(string UserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListUserInRoles");

            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, UserID);

            return dw.ExecuteCommandTable(cmd);
        }



        public DataTable SelectEmployees(string EmployeeID, string IsAdmin, string NTName, string Domain, string PartnerID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectEmployees");

            dw.CreateParameter(cmd, "@p_EmployeeID", SqlDbType.Int, EmployeeID);
            dw.CreateParameter(cmd, "@p_IsAdmin", SqlDbType.Bit, IsAdmin);
            dw.CreateParameter(cmd, "@p_NTName", SqlDbType.VarChar, NTName, 30);
            dw.CreateParameter(cmd, "@p_Domain", SqlDbType.VarChar, Domain, 30);
            dw.CreateParameter(cmd, "@p_PartnerID", SqlDbType.Int, PartnerID);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable ListProducts(string MinimumStatusID, string MaximumStatusID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListProducts");

            dw.CreateParameter(cmd, "@p_MinimumStatusID", SqlDbType.Int, MinimumStatusID);
            dw.CreateParameter(cmd, "@p_MaximumStatusID", SqlDbType.Int, MaximumStatusID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable ListProductsByDivision(int Division)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetProductsByDivision");

            dw.CreateParameter(cmd, "@Div", SqlDbType.TinyInt, Division.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListPrograms(string BusinessID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListPrograms2");

            dw.CreateParameter(cmd, "@BusinessID", SqlDbType.Int, BusinessID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListDevCenters()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListDevCenters");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListProductStatuses()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListProductStatuses");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListScheduleMilestones(string ReportProfileID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScheduleDefinitionTriggers");

            dw.CreateParameter(cmd, "@p_ReportProfileID", SqlDbType.Int, ReportProfileID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListScheduleMilestones(string ReportProfileID, string SelectedProductVerIds)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScheduleDefinitionTriggers");

            dw.CreateParameter(cmd, "@p_ReportProfileID", SqlDbType.Int, ReportProfileID);
            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.VarChar, SelectedProductVerIds);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListReportProfiles(int EmployeeID, int Type)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListReportProfiles");

            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID.ToString());
            dw.CreateParameter(cmd, "@Type", SqlDbType.Int, Type.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListReportProfilesShared(int EmployeeID, int Type)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListReportProfilesShared");

            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID.ToString());
            dw.CreateParameter(cmd, "@Type", SqlDbType.Int, Type.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable ListReportProfilesGroupShared(int EmployeeID, int Type)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListReportProfilesGroupShared");

            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID.ToString());
            dw.CreateParameter(cmd, "@Type", SqlDbType.Int, Type.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        [DataObjectMethodAttribute(DataObjectMethodType.Select, true)]
        public DataTable ListPartners(string ReportType, string PartnerTypeId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spListPartners");

            dw.CreateParameter(cmd, "@ReportType", SqlDbType.Int, ReportType);
            dw.CreateParameter(cmd, "@PartnerTypeId", SqlDbType.Int, PartnerTypeId);

            return dw.ExecuteCommandTable(cmd);
        }
        [DataObjectMethodAttribute(DataObjectMethodType.Select, true)]

        public DataTable ListPartners(int ReportType)
        {
            DataTable dt = ListPartners(ReportType.ToString(), string.Empty);
            return dt;
        }


        public long AddReportProfile(string ProfileName, string ProfileType, string EmployeeID, string Value15, string value45, string value46, string value47, string value52)
        {
            return AddReportProfile(ProfileName, ProfileType, EmployeeID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, Value15, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, value45, value46, value47, string.Empty, string.Empty, string.Empty,
                string.Empty, value52, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        public long AddReportProfile(string ProfileName, string ProfileType, string EmployeeID, string Value15, string Value45, string Value17)
        {
            return AddReportProfile(ProfileName, ProfileType, EmployeeID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, Value15, string.Empty, Value17, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, Value45, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
        }
        public long AddReportProfile(string ProfileName, string ProfileType, string EmployeeID, string Value5, string Value6, string Value7, string Value8, string Value9, string Value10, string Value11, string Value12, string Value13, string Value14, string Value15, string Value16, string Value17, string Value18, string Value19, string Value20, string Value21, string Value22, string Value23, string Value24, string Value25, string Value27, string Value28, string Value29, string Value30, string Value31, string Value32, string Value33, string Value34, string Value35, string Value36, string Value37, string Value38, string Value41, string Value42, string Value44, string Value45, string Value46, string Value1, string Value2, string Value3, string Value4, string Value26, string Value39, string Value40, string Value43, string DefaultSQL, string ReportFilters)
        {

            return AddReportProfile(ProfileName, ProfileType, EmployeeID, Value5, Value6, Value7, Value8, Value9, Value10, Value11,
                Value12, Value13, Value14, Value15, Value16, Value17, Value18, Value19, Value20, Value21, Value22,
                Value23, Value24, Value25, Value27, Value28, Value29, Value30, Value31, Value32, Value33, Value34,
                Value35, Value36, Value37, Value38, Value41, Value42, Value44, Value45, Value46, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, Value1, Value2, Value3, Value4, Value26,
                Value39, Value40, Value43, DefaultSQL, ReportFilters);
        }

        public long AddReportProfile(string ProfileName, string ProfileType, string EmployeeID, string Value5, string Value6, string Value7, string Value8, string Value9, string Value10, string Value11,
            string Value12, string Value13, string Value14, string Value15, string Value16, string Value17, string Value18, string Value19, string Value20, string Value21, string Value22,
            string Value23, string Value24, string Value25, string Value27, string Value28, string Value29, string Value30, string Value31, string Value32, string Value33, string Value34,
            string Value35, string Value36, string Value37, string Value38, string Value41, string Value42, string Value44, string Value45, string Value46, string Value47, string Value48,
            string Value49, string Value50, string Value51, string Value52, string Value53, string Value54, string Value1, string Value2, string Value3, string Value4, string Value26,
            string Value39, string Value40, string Value43, string DefaultSQL, string ReportFilters)
        {

            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spAddProfile");

            dw.CreateParameter(cmd, "@ProfileName", SqlDbType.VarChar, ProfileName.ToString());
            dw.CreateParameter(cmd, "@ProfileType", SqlDbType.Int, ProfileType.ToString());
            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID.ToString());
            dw.CreateParameter(cmd, "@Value5", SqlDbType.Int, Value5.ToString());
            dw.CreateParameter(cmd, "@Value6", SqlDbType.VarChar, Value6.ToString());
            dw.CreateParameter(cmd, "@Value7", SqlDbType.Int, Value7.ToString());
            dw.CreateParameter(cmd, "@Value8", SqlDbType.Int, Value8.ToString());
            dw.CreateParameter(cmd, "@Value9", SqlDbType.VarChar, Value9.ToString());
            dw.CreateParameter(cmd, "@Value10", SqlDbType.VarChar, Value10.ToString());
            dw.CreateParameter(cmd, "@Value11", SqlDbType.VarChar, Value11.ToString());
            dw.CreateParameter(cmd, "@Value12", SqlDbType.VarChar, Value12.ToString());
            dw.CreateParameter(cmd, "@Value13", SqlDbType.VarChar, Value13.ToString());
            dw.CreateParameter(cmd, "@Value14", SqlDbType.VarChar, Value14.ToString());
            dw.CreateParameter(cmd, "@Value15", SqlDbType.VarChar, Value15.ToString());
            dw.CreateParameter(cmd, "@Value16", SqlDbType.Int, Value16.ToString());
            dw.CreateParameter(cmd, "@Value17", SqlDbType.Bit, Value17.ToString());
            dw.CreateParameter(cmd, "@Value18", SqlDbType.Bit, Value18.ToString());
            dw.CreateParameter(cmd, "@Value19", SqlDbType.Bit, Value19.ToString());
            dw.CreateParameter(cmd, "@Value20", SqlDbType.Bit, Value20.ToString());
            dw.CreateParameter(cmd, "@Value21", SqlDbType.Bit, Value21.ToString());
            dw.CreateParameter(cmd, "@Value22", SqlDbType.Int, Value22.ToString());
            dw.CreateParameter(cmd, "@Value23", SqlDbType.VarChar, Value23.ToString());
            dw.CreateParameter(cmd, "@Value24", SqlDbType.Int, Value24.ToString());
            dw.CreateParameter(cmd, "@Value25", SqlDbType.VarChar, Value25.ToString());
            dw.CreateParameter(cmd, "@Value27", SqlDbType.VarChar, Value27.ToString());
            dw.CreateParameter(cmd, "@Value28", SqlDbType.VarChar, Value28.ToString());
            dw.CreateParameter(cmd, "@Value29", SqlDbType.VarChar, Value29.ToString());
            dw.CreateParameter(cmd, "@Value30", SqlDbType.VarChar, Value30.ToString());
            dw.CreateParameter(cmd, "@Value31", SqlDbType.VarChar, Value31.ToString());
            dw.CreateParameter(cmd, "@Value32", SqlDbType.VarChar, Value32.ToString());
            dw.CreateParameter(cmd, "@Value33", SqlDbType.Bit, Value33.ToString());
            dw.CreateParameter(cmd, "@Value34", SqlDbType.Bit, Value34.ToString());
            dw.CreateParameter(cmd, "@Value35", SqlDbType.Bit, Value35.ToString());
            dw.CreateParameter(cmd, "@Value36", SqlDbType.Bit, Value36.ToString());
            dw.CreateParameter(cmd, "@Value37", SqlDbType.Bit, Value37.ToString());
            dw.CreateParameter(cmd, "@Value38", SqlDbType.Bit, Value38.ToString());
            dw.CreateParameter(cmd, "@Value41", SqlDbType.VarChar, Value41.ToString());
            dw.CreateParameter(cmd, "@Value42", SqlDbType.VarChar, Value42.ToString());
            dw.CreateParameter(cmd, "@Value44", SqlDbType.Bit, Value44.ToString());
            dw.CreateParameter(cmd, "@Value45", SqlDbType.VarChar, Value45.ToString());
            dw.CreateParameter(cmd, "@Value46", SqlDbType.VarChar, Value46.ToString());
            dw.CreateParameter(cmd, "@Value47", SqlDbType.VarChar, Value47.ToString());
            dw.CreateParameter(cmd, "@Value48", SqlDbType.Bit, Value48.ToString());
            dw.CreateParameter(cmd, "@Value49", SqlDbType.Bit, Value49.ToString());
            dw.CreateParameter(cmd, "@Value50", SqlDbType.Bit, Value50.ToString());
            dw.CreateParameter(cmd, "@Value51", SqlDbType.Bit, Value51.ToString());
            dw.CreateParameter(cmd, "@Value52", SqlDbType.VarChar, Value52.ToString());
            dw.CreateParameter(cmd, "@Value53", SqlDbType.VarChar, Value53.ToString());
            dw.CreateParameter(cmd, "@Value54", SqlDbType.VarChar, Value54.ToString());
            dw.CreateParameter(cmd, "@Value1", SqlDbType.VarChar, Value1.ToString());
            dw.CreateParameter(cmd, "@Value2", SqlDbType.VarChar, Value2.ToString());
            dw.CreateParameter(cmd, "@Value3", SqlDbType.VarChar, Value3.ToString());
            dw.CreateParameter(cmd, "@Value4", SqlDbType.VarChar, Value4.ToString());
            dw.CreateParameter(cmd, "@Value26", SqlDbType.VarChar, Value26.ToString());
            dw.CreateParameter(cmd, "@Value39", SqlDbType.VarChar, Value39.ToString());
            dw.CreateParameter(cmd, "@Value40", SqlDbType.VarChar, Value40.ToString());
            dw.CreateParameter(cmd, "@Value43", SqlDbType.VarChar, Value43.ToString());
            dw.CreateParameter(cmd, "@DefaultSQL", SqlDbType.VarChar, DefaultSQL.ToString());
            dw.CreateParameter(cmd, "@ReportFilters", SqlDbType.VarChar, ReportFilters.ToString());
            dw.CreateParameter(cmd, "@NewID", SqlDbType.Int, string.Empty, ParameterDirection.Output);

            long returnValue = dw.ExecuteCommandNonQuery(cmd);
            return Convert.ToInt64(cmd.Parameters["@NewID"].Value);
        }



        public DataTable GetReportProfileShared(string ProfileID, string EmployeeID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetReportProfileShared");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ProfileID);
            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public DataTable GetReportProfile(string ProfileID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetReportProfile");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ProfileID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable GetReportProfileGroup(string ProfileID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetReportProfileGroups");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ProfileID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public int DeleteReportProfile(string ProfileID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spDeleteProfile");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ProfileID);

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }


        public int UpdateProfile(string ProfileID, string Value15, string Value45, string Value46, string Value47, string Value52)
        {
            return UpdateProfile(ProfileID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, Value15, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, Value45, Value46, Value47, string.Empty, string.Empty,
                string.Empty, string.Empty, Value52, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        public int UpdateProfile(string ProfileID, string Value15, string Value45, string Value17)
        {
            return UpdateProfile(ProfileID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, Value15, string.Empty, Value17, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, Value45, string.Empty, string.Empty, string.Empty, string.Empty,
                string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

        }
        public int UpdateProfile(string ID, string Value5, string Value6, string Value7, string Value8, string Value9, string Value10, string Value11, string Value12, string Value13, string Value14, string Value15, string Value16, string Value17, string Value18, string Value19, string Value20, string Value21, string Value22, string Value23, string Value24, string Value25, string Value27, string Value28, string Value29, string Value30, string Value31, string Value32, string Value33, string Value34, string Value35, string Value36, string Value37, string Value38, string Value41, string Value42, string Value44, string Value45, string Value46, string Value47, string Value48, string Value49, string Value50, string Value51, string Value52, string Value53, string Value54, string Value1, string Value2, string Value3, string Value4, string Value26, string Value39, string Value40, string Value43, string DefaultSQL, string ReportFilters)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spUpdateProfile");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ID.ToString());
            dw.CreateParameter(cmd, "@Value5", SqlDbType.Int, Value5.ToString());
            dw.CreateParameter(cmd, "@Value6", SqlDbType.VarChar, Value6.ToString());
            dw.CreateParameter(cmd, "@Value7", SqlDbType.Int, Value7.ToString());
            dw.CreateParameter(cmd, "@Value8", SqlDbType.Int, Value8.ToString());
            dw.CreateParameter(cmd, "@Value9", SqlDbType.VarChar, Value9.ToString());
            dw.CreateParameter(cmd, "@Value10", SqlDbType.VarChar, Value10.ToString());
            dw.CreateParameter(cmd, "@Value11", SqlDbType.VarChar, Value11.ToString());
            dw.CreateParameter(cmd, "@Value12", SqlDbType.VarChar, Value12.ToString());
            dw.CreateParameter(cmd, "@Value13", SqlDbType.VarChar, Value13.ToString());
            dw.CreateParameter(cmd, "@Value14", SqlDbType.VarChar, Value14.ToString());
            dw.CreateParameter(cmd, "@Value15", SqlDbType.VarChar, Value15.ToString());
            dw.CreateParameter(cmd, "@Value16", SqlDbType.Int, Value16.ToString());
            dw.CreateParameter(cmd, "@Value17", SqlDbType.Bit, Value17.ToString());
            dw.CreateParameter(cmd, "@Value18", SqlDbType.Bit, Value18.ToString());
            dw.CreateParameter(cmd, "@Value19", SqlDbType.Bit, Value19.ToString());
            dw.CreateParameter(cmd, "@Value20", SqlDbType.Bit, Value20.ToString());
            dw.CreateParameter(cmd, "@Value21", SqlDbType.Bit, Value21.ToString());
            dw.CreateParameter(cmd, "@Value22", SqlDbType.Int, Value22.ToString());
            dw.CreateParameter(cmd, "@Value23", SqlDbType.VarChar, Value23.ToString());
            dw.CreateParameter(cmd, "@Value24", SqlDbType.Int, Value24.ToString());
            dw.CreateParameter(cmd, "@Value25", SqlDbType.VarChar, Value25.ToString());
            dw.CreateParameter(cmd, "@Value27", SqlDbType.VarChar, Value27.ToString());
            dw.CreateParameter(cmd, "@Value28", SqlDbType.VarChar, Value28.ToString());
            dw.CreateParameter(cmd, "@Value29", SqlDbType.VarChar, Value29.ToString());
            dw.CreateParameter(cmd, "@Value30", SqlDbType.VarChar, Value30.ToString());
            dw.CreateParameter(cmd, "@Value31", SqlDbType.VarChar, Value31.ToString());
            dw.CreateParameter(cmd, "@Value32", SqlDbType.VarChar, Value32.ToString());
            dw.CreateParameter(cmd, "@Value33", SqlDbType.Bit, Value33.ToString());
            dw.CreateParameter(cmd, "@Value34", SqlDbType.Bit, Value34.ToString());
            dw.CreateParameter(cmd, "@Value35", SqlDbType.Bit, Value35.ToString());
            dw.CreateParameter(cmd, "@Value36", SqlDbType.Bit, Value36.ToString());
            dw.CreateParameter(cmd, "@Value37", SqlDbType.Bit, Value37.ToString());
            dw.CreateParameter(cmd, "@Value38", SqlDbType.Bit, Value38.ToString());
            dw.CreateParameter(cmd, "@Value41", SqlDbType.VarChar, Value41.ToString());
            dw.CreateParameter(cmd, "@Value42", SqlDbType.VarChar, Value42.ToString());
            dw.CreateParameter(cmd, "@Value44", SqlDbType.Bit, Value44.ToString());
            dw.CreateParameter(cmd, "@Value45", SqlDbType.VarChar, Value45.ToString());
            dw.CreateParameter(cmd, "@Value46", SqlDbType.VarChar, Value46.ToString());
            dw.CreateParameter(cmd, "@Value47", SqlDbType.VarChar, Value47.ToString());
            dw.CreateParameter(cmd, "@Value48", SqlDbType.Bit, Value48.ToString());
            dw.CreateParameter(cmd, "@Value49", SqlDbType.Bit, Value49.ToString());
            dw.CreateParameter(cmd, "@Value50", SqlDbType.Bit, Value50.ToString());
            dw.CreateParameter(cmd, "@Value51", SqlDbType.Bit, Value50.ToString());
            dw.CreateParameter(cmd, "@Value52", SqlDbType.VarChar, Value52.ToString());
            dw.CreateParameter(cmd, "@Value53", SqlDbType.VarChar, Value53.ToString());
            dw.CreateParameter(cmd, "@Value54", SqlDbType.VarChar, Value54.ToString());
            dw.CreateParameter(cmd, "@Value1", SqlDbType.VarChar, Value1.ToString());
            dw.CreateParameter(cmd, "@Value2", SqlDbType.VarChar, Value2.ToString());
            dw.CreateParameter(cmd, "@Value3", SqlDbType.VarChar, Value3.ToString());
            dw.CreateParameter(cmd, "@Value4", SqlDbType.VarChar, Value4.ToString());
            dw.CreateParameter(cmd, "@Value26", SqlDbType.VarChar, Value26.ToString());
            dw.CreateParameter(cmd, "@Value39", SqlDbType.VarChar, Value39.ToString());
            dw.CreateParameter(cmd, "@Value40", SqlDbType.VarChar, Value40.ToString());
            dw.CreateParameter(cmd, "@Value43", SqlDbType.VarChar, Value43.ToString());
            dw.CreateParameter(cmd, "@DefaultSQL", SqlDbType.VarChar, DefaultSQL.ToString());
            dw.CreateParameter(cmd, "@ReportFilters", SqlDbType.VarChar, ReportFilters.ToString());

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }


        public DataTable RenameProfile(string ProfileID, string Name)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spRenameProfile");

            dw.CreateParameter(cmd, "@ID", SqlDbType.Int, ProfileID);
            dw.CreateParameter(cmd, "@Name", SqlDbType.VarChar, Name);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public long UpdateReportTrigger(string ObjectTypeID, string ObjectID, string ReportProfileID, string DaysDiff, string SendEmail, string CreateActionItem, string NoteToSelf, string LastUpdUser)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateReportTrigger");

            dw.CreateParameter(cmd, "@p_ObjectTypeID", SqlDbType.Int, ObjectTypeID);
            dw.CreateParameter(cmd, "@p_ObjectID", SqlDbType.Int, ObjectID);
            dw.CreateParameter(cmd, "@p_ReportProfileID", SqlDbType.Int, ReportProfileID);
            dw.CreateParameter(cmd, "@p_DaysDiff", SqlDbType.Int, DaysDiff);
            dw.CreateParameter(cmd, "@p_SendEmail", SqlDbType.Bit, SendEmail);
            dw.CreateParameter(cmd, "@p_CreateActionItem", SqlDbType.Bit, CreateActionItem);
            dw.CreateParameter(cmd, "@p_NoteToSelf", SqlDbType.VarChar, NoteToSelf);
            dw.CreateParameter(cmd, "@p_LastUpdUser", SqlDbType.VarChar, LastUpdUser);
            dw.CreateParameter(cmd, "@p_ReportTriggerID", SqlDbType.Int, string.Empty, ParameterDirection.Output);

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            long reportTriggerID = Convert.ToInt64(cmd.Parameters["@p_ReportTriggerID"].Value);
            return reportTriggerID;
        }

        public int DeleteReportTrigger(string ReportTriggerID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_DeleteReportTrigger");

            dw.CreateParameter(cmd, "@p_ReportTriggerID", SqlDbType.Int, ReportTriggerID.ToString());

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }



        public int DeleteReportTriggerByProfileID(string ReportProfileID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_DeleteReportTriggerByProfileID");

            dw.CreateParameter(cmd, "@p_ReportProfileID", SqlDbType.Int, ReportProfileID);

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }


        public int RemoveReportProfile(string ProfileID, string EmployeeID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spRemoveSharedProfile2");

            dw.CreateParameter(cmd, "@ProfileID", SqlDbType.Int, ProfileID);
            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID);

            int returnValue = dw.ExecuteCommandNonQuery(cmd);
            return returnValue;
        }


        public int InsertReportTriggerProductVersion(string ReportTriggerID, string ProductVersionID, string PartnerID, string ProgramID,
            string DevCenterID, string ProductStatusID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertReportTriggerProductVersion");

            dw.CreateParameter(cmd, "@p_ReportTriggerID", SqlDbType.Int, ReportTriggerID);
            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@p_PartnerID", SqlDbType.Int, PartnerID);
            dw.CreateParameter(cmd, "@p_ProgramID", SqlDbType.Int, ProgramID);
            dw.CreateParameter(cmd, "@p_DevCenterID", SqlDbType.Int, DevCenterID);
            dw.CreateParameter(cmd, "@p_ProductStatusID", SqlDbType.Int, ProductStatusID);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public int DeleteReportTriggerProductVersionByTriggerID(string ReportTriggerID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_DeleteReportTriggerProductVersionByTrigger");

            dw.CreateParameter(cmd, "@p_ReportTriggerID", SqlDbType.Int, ReportTriggerID);

            return dw.ExecuteCommandNonQuery(cmd);
        }


        public DataTable SelectScheduleSummaryData(string ProductVersions, string Partners, string DevCenters, string Programs, string Status, string Milestones)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScheduleSummaryData");

            dw.CreateParameter(cmd, "@p_ProductVersionIdList", SqlDbType.VarChar, ProductVersions.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_PartnerIdList", SqlDbType.VarChar, Partners.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_DevCenterIdList", SqlDbType.VarChar, DevCenters.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_ProgramIdList", SqlDbType.VarChar, Programs.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_Status", SqlDbType.VarChar, Status, 8000);
            dw.CreateParameter(cmd, "@p_Milestones", SqlDbType.VarChar, Milestones, 8000);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectScheduleHistoryData(string ProductVersions, string Partners, string DevCenters, string Programs, string Status, string Milestones)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectScheduleHistoryData");

            dw.CreateParameter(cmd, "@p_ProductVersionIdList", SqlDbType.VarChar, ProductVersions.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_PartnerIdList", SqlDbType.VarChar, Partners.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_DevCenterIdList", SqlDbType.VarChar, DevCenters.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_ProgramIdList", SqlDbType.VarChar, Programs.ToString(), 8000);
            dw.CreateParameter(cmd, "@p_Status", SqlDbType.VarChar, Status, 8000);
            dw.CreateParameter(cmd, "@p_Milestones", SqlDbType.VarChar, Milestones, 8000);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectRslChangeLog(string ServiceFamilyPn)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectRslHistory");

            // robert.richardson@hp.com - Task 11260 - Allow multiple partnumbers to be entered, so changed limit of 10 char to 1000
            dw.CreateParameter(cmd, "@p_ServiceFamilyPn", SqlDbType.VarChar, ServiceFamilyPn, 1000);

            return dw.ExecuteCommandTable(cmd);
        }

        public DataTable usp_SelectServiceSpareKitsForProduct(string ProductVersionId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("dbo.usp_SelectServiceSpareKitsForProduct");

            dw.CreateParameter(cmd, "@p_ProductVersionId", SqlDbType.Int, ProductVersionId.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectServiceSpareKitAvMap(string ProductBrandId, string ServiceSpareKitMapId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectServiceSpareKitMapAv");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandId);
            dw.CreateParameter(cmd, "@p_ServiceSpareKitMapId", SqlDbType.Int, ServiceSpareKitMapId);

            return dw.ExecuteCommandTable(cmd);
        }

        public int InsertServiceSpareKitAvMap(string SpareKitMapId, string AvCategoryId, string AvNo)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertServiceSpareKitMapAv");
            cmd.CommandTimeout = 120;

            dw.CreateParameter(cmd, "@p_ServiceSpareKitMapId", SqlDbType.Int, SpareKitMapId);
            dw.CreateParameter(cmd, "@p_AvFeatureCategoryId", SqlDbType.Int, AvCategoryId);
            dw.CreateParameter(cmd, "@p_AvNo", SqlDbType.VarChar, AvNo, 15);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public int DeleteServiceSpareKitAvMap(string ServiceSpareKitMapId, string AvNo)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_DeleteServiceSpareKitMapAv");

            dw.CreateParameter(cmd, "@p_ServiceSpareKitMapid", SqlDbType.Int, ServiceSpareKitMapId);
            dw.CreateParameter(cmd, "@p_AvNo", SqlDbType.VarChar, AvNo, 25);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public string GetNewSpareKitMapId(string ProductBrandId, string SpareKitId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertServiceSpareKitMap");
            string serviceSpareKitMapId = string.Empty;

            dw.CreateParameter(cmd, "@p_ProductBrandId", SqlDbType.Int, ProductBrandId);
            dw.CreateParameter(cmd, "@p_SpareKitId", SqlDbType.Int, SpareKitId);
            dw.CreateParameter(cmd, "@p_ServiceSpareKitMapId", SqlDbType.Int, string.Empty, ParameterDirection.Output);

            dw.ExecuteCommandNonQuery(cmd);

            return cmd.Parameters["@p_ServiceSpareKitMapId"].Value.ToString();
        }

        public DataTable SelectAvByBrandCategory(string ProductBrandId, string AvFeatureCategoryId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvByBrandCategory");

            dw.CreateParameter(cmd, "@p_ProductBrandId", SqlDbType.Int, ProductBrandId.ToString());
            dw.CreateParameter(cmd, "@p_AvFeatureCategoryId", SqlDbType.Int, AvFeatureCategoryId.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectAvFeatureCategoriesForService(string ProductBrandId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_ListAvFeatureCategoriesForService");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandId);

            return dw.ExecuteCommandTable(cmd);
        }


        public string UpdateAvDetailViaUpload(string AVNo, string GPGDescription, string ProductVersionID, string ProductBrandID, string UserName)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateAvDetailViaUpload");

            dw.CreateParameter(cmd, "@AVNo", SqlDbType.VarChar, AVNo, 50);
            dw.CreateParameter(cmd, "@GPGDescription", SqlDbType.VarChar, GPGDescription, 50);
            dw.CreateParameter(cmd, "@PVID", SqlDbType.Int, ProductVersionID.ToString());
            dw.CreateParameter(cmd, "@BID", SqlDbType.Int, ProductBrandID.ToString());
            dw.CreateParameter(cmd, "@UserName", SqlDbType.VarChar, UserName, 50);
            dw.CreateParameter(cmd, "@ReturnValue", SqlDbType.Int, string.Empty, ParameterDirection.Output);


            dw.ExecuteCommandNonQuery(cmd);

            return cmd.Parameters["@ReturnValue"].Value.ToString();
        }

        public DataTable SelectDocKitsByKMAT(string KMAT)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDocKitsByKMAT");

            dw.CreateParameter(cmd, "@p_KMAT", SqlDbType.VarChar, KMAT.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectDCRWorkflows()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflows");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectDCRWorkflowsDefinitions()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflowDefinitions");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public int InsertDCRWorkflowHistory(string DCRID, string WorkflowID, string UserID, string PVID, string RTPDate, string EMDate)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_InsertDCRWorkflowHistory");

            dw.CreateParameter(cmd, "@DCRID", SqlDbType.Int, DCRID);
            dw.CreateParameter(cmd, "@WorkflowID", SqlDbType.Int, WorkflowID);
            dw.CreateParameter(cmd, "@UserID", SqlDbType.Int, UserID);
            dw.CreateParameter(cmd, "@PVID", SqlDbType.Int, PVID);
            dw.CreateParameter(cmd, "@RTPDate", SqlDbType.VarChar, RTPDate);
            dw.CreateParameter(cmd, "@EMDate", SqlDbType.VarChar, EMDate);

            return dw.ExecuteCommandNonQuery(cmd);
        }
        public int IsPulsarProduct(string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_IsPulsarProduct");

            dw.CreateParameter(cmd, "@p_intPVID", SqlDbType.Int, PVID);

            int retVal = Convert.ToInt32(dw.ExecuteCommandScalar(cmd));
            return retVal;
        }

        public DataTable SelectDCRWorkflowStatus(string DCRID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflowStatus");

            dw.CreateParameter(cmd, "@DCRID", SqlDbType.Int, DCRID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public int UpdateDCRWorkflowComments(string HistoryID, string Comments)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateDCRWorkflowComments");

            dw.CreateParameter(cmd, "@p_HistoryID", SqlDbType.Int, HistoryID);
            dw.CreateParameter(cmd, "@p_Comments", SqlDbType.VarChar, Comments);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public int UpdateDCRWorkflowComplete(string HistoryID, string DCRID, string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateDCRWorkflowComplete");

            dw.CreateParameter(cmd, "@p_HistoryID", SqlDbType.Int, HistoryID);
            dw.CreateParameter(cmd, "@p_DCRID", SqlDbType.Int, DCRID);
            dw.CreateParameter(cmd, "@p_PVID", SqlDbType.Int, PVID);

            return dw.ExecuteCommandNonQuery(cmd);
        }

        public DataTable SelectDCRWorkflowEmailList(string DCRID, string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflowEmailList");

            dw.CreateParameter(cmd, "@DCRID", SqlDbType.Int, DCRID);
            dw.CreateParameter(cmd, "@PVID", SqlDbType.Int, PVID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectAvWithMissingData(string BID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvWithMissingData");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, BID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectHiddenAvs(string BID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectHiddenAvs");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, BID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable UpdateAvStatus(string AVID, string Hide)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateAvStatus");

            dw.CreateParameter(cmd, "@p_AVDetailID", SqlDbType.Int, AVID.ToString());
            dw.CreateParameter(cmd, "@p_HideAv", SqlDbType.Int, Hide.ToString());


            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectProductsByCycle()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductsByCycle");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectAvActionItems(string AvId)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvActionItems");

            dw.CreateParameter(cmd, "@p_AvId", SqlDbType.Int, AvId);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable GetEmployeeUserSettings(string EmployeeID, string UserSettingsID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spGetEmployeeUserSettings");

            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID);
            dw.CreateParameter(cmd, "@UserSettingsID", SqlDbType.Int, UserSettingsID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable UpdateEmployeeUserSetting(string EmployeeID, string UserSettingsID, string Value)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("spUpdateEmployeeUserSetting2");

            dw.CreateParameter(cmd, "@EmployeeID", SqlDbType.Int, EmployeeID);
            dw.CreateParameter(cmd, "@UserSettingsID", SqlDbType.Int, UserSettingsID);
            dw.CreateParameter(cmd, "@Value", SqlDbType.VarChar, Value.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectDCRWorkflowsByProduct(string PVIDs, string WorkflowID, string FromDate, string ToDate, string Days, string DateRangeType)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflowsByProduct");

            dw.CreateParameter(cmd, "@p_PVIDs", SqlDbType.VarChar, PVIDs);
            dw.CreateParameter(cmd, "@p_WorkflowID", SqlDbType.Int, WorkflowID);
            dw.CreateParameter(cmd, "@p_FromDate", SqlDbType.VarChar, FromDate);
            dw.CreateParameter(cmd, "@p_ToDate", SqlDbType.VarChar, ToDate);
            dw.CreateParameter(cmd, "@p_Days", SqlDbType.Int, Days);
            dw.CreateParameter(cmd, "@p_DateRangeType", SqlDbType.Int, DateRangeType);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectDCRWorkflowsDates(string DCRIDs)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRWorkflowsDates");

            dw.CreateParameter(cmd, "@p_DCRIDs", SqlDbType.VarChar, DCRIDs);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectDCRSummaries(string DCRIDs)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRSummaries");

            dw.CreateParameter(cmd, "@p_DCRIDs", SqlDbType.VarChar, DCRIDs);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectTerminatedDCRSummaries(string PVIDs, string WorkflowID, string FromDate, string ToDate, string Days, string DateRangeType)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectDCRSummariesTerminated");

            dw.CreateParameter(cmd, "@p_PVIDs", SqlDbType.VarChar, PVIDs);
            dw.CreateParameter(cmd, "@p_WorkflowID", SqlDbType.Int, WorkflowID);
            dw.CreateParameter(cmd, "@p_FromDate", SqlDbType.VarChar, FromDate);
            dw.CreateParameter(cmd, "@p_ToDate", SqlDbType.VarChar, ToDate);
            dw.CreateParameter(cmd, "@p_Days", SqlDbType.Int, Days);
            dw.CreateParameter(cmd, "@p_DateRangeType", SqlDbType.Int, DateRangeType);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectSuperUsersByProduct(string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectSuperUsersByProduct");

            dw.CreateParameter(cmd, "@p_PVID", SqlDbType.Int, PVID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectAvFeatureCategoriesFilter(string BusinessID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvFeatureCategories");

            dw.CreateParameter(cmd, "@p_BusinessID", SqlDbType.Int, BusinessID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectSCMCategoriesFilter()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SCM_GetSCMCategories");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public DataTable Product_GetProductReleases(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Product_GetProductReleases");
            dw.CreateParameter(cmd, "@p_intProductVersionID", SqlDbType.Int, ProductVersionID);
            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public DataTable Product_GetDefaultDates(string ProductVersionID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_Product_GetProductDefaultDates");
            dw.CreateParameter(cmd, "@p_intProductVersionID", SqlDbType.Int, ProductVersionID);
            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }
        public DataTable SelectInitialOfferingCategories()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingCategories");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectInitialOfferingData(string BusinessID, string CategoryID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingData");

            dw.CreateParameter(cmd, "@p_BusinessID", SqlDbType.Int, BusinessID);
            dw.CreateParameter(cmd, "@p_CategoryID", SqlDbType.Int, CategoryID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectInitialOfferingDeliverables(string CategoryID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingDeliverables");

            dw.CreateParameter(cmd, "@p_CategoryID", SqlDbType.Int, CategoryID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectInitialOfferingProducts(string BusinessID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingProducts");

            dw.CreateParameter(cmd, "@p_BusinessID", SqlDbType.Int, BusinessID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectInitialOfferingRoleStatus(string UserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingRoleStatus");

            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, UserID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable UpdateInitialOfferingAVChanges(string Business, string CurrentUser)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateInitialOfferingAVChanges");

            dw.CreateParameter(cmd, "@p_BusinessID", SqlDbType.Int, Business);
            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, CurrentUser);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable UpdateInitialOfferingLTFAVSAs(string DeliverableRootID, string LTFAV, string LTFSA, string ActionItemID, string CurrentUserID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateInitialOfferingLTFAVSAs");

            dw.CreateParameter(cmd, "@p_DeliverableRootID", SqlDbType.Int, DeliverableRootID);
            dw.CreateParameter(cmd, "@p_LTFAV", SqlDbType.VarChar, LTFAV);
            dw.CreateParameter(cmd, "@p_LTFSA", SqlDbType.VarChar, LTFSA);
            dw.CreateParameter(cmd, "@p_ActionItemID", SqlDbType.Int, ActionItemID);
            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, CurrentUserID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectInitialOfferingMarketingReq(string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectInitialOfferingMarketingReq");

            dw.CreateParameter(cmd, "@p_PVID", SqlDbType.Int, PVID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectProductsByDeliverable(string RootID, string SAType, string Assign)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductsByDeliverable");

            dw.CreateParameter(cmd, "@p_RootID", SqlDbType.Int, RootID);
            dw.CreateParameter(cmd, "@p_SAType", SqlDbType.Int, SAType);
            dw.CreateParameter(cmd, "@p_Assign", SqlDbType.Int, Assign);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectProductsByDeliverableRelease(string RootID, string SAType, string Assign)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectProductsByDeliverableRelease");

            dw.CreateParameter(cmd, "@p_RootID", SqlDbType.Int, RootID);
            dw.CreateParameter(cmd, "@p_SAType", SqlDbType.Int, SAType);
            dw.CreateParameter(cmd, "@p_Assign", SqlDbType.Int, Assign);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectCommodityGuidanceProductsByProgram(string ProductProgram)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectCommodityGuidanceProductsByProgram");

            dw.CreateParameter(cmd, "@p_ProductProgram", SqlDbType.Int, ProductProgram);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectCommodityGuidanceData(string ProductProgram, string Category)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectCommodityGuidanceData");

            dw.CreateParameter(cmd, "@p_ProductProgram", SqlDbType.Int, ProductProgram);
            dw.CreateParameter(cmd, "@p_CategoryID", SqlDbType.Int, Category);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectCommodityGuidanceCategories()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectCommodityGuidanceCategories");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectCommodityGuidanceProductPrograms()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectCommodityGuidanceProductPrograms");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable SelectAVNoDescriptions(string ProductVersionID, string ProductBrandID, string SCMCategoryID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SCM_GetAVNoDescriptions");

            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, ProductVersionID);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);
            dw.CreateParameter(cmd, "@p_SCMCategoryID", SqlDbType.Int, SCMCategoryID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable UpdateAvDetailFeatureID(string AvCreateID, string FeatureID, string UserID, string AvNo)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateAvDetail_FeatureID");

            dw.CreateParameter(cmd, "@p_AvCreateID", SqlDbType.Int, AvCreateID);
            dw.CreateParameter(cmd, "@p_FeatureID", SqlDbType.Int, FeatureID);
            dw.CreateParameter(cmd, "@p_AvNo", SqlDbType.VarChar, AvNo);
            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, UserID);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;

        }


        public DataTable UpdateAvDetailDeliverableRootID(string AvCreateID, string DeliverableRootID, string ProductBrandID, string UserID, string AvNo, string UpdateDescriptions)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateAvDetail_DeliverableRootID");

            dw.CreateParameter(cmd, "@p_AvCreateID", SqlDbType.Int, AvCreateID);
            dw.CreateParameter(cmd, "@p_DeliverableRootID", SqlDbType.Int, DeliverableRootID);
            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, ProductBrandID);
            dw.CreateParameter(cmd, "@p_UserID", SqlDbType.Int, UserID);
            dw.CreateParameter(cmd, "@p_AvNo", SqlDbType.VarChar, AvNo);
            dw.CreateParameter(cmd, "@p_UpdateDescriptions", SqlDbType.TinyInt, UpdateDescriptions);

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectAvActionScorecardProducts()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvActionScorecardProducts");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectAvsMissingDeliverableRoot(string BID, string PVID)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectAvsMissingDeliverableRoot");

            dw.CreateParameter(cmd, "@p_ProductBrandID", SqlDbType.Int, BID.ToString());
            dw.CreateParameter(cmd, "@p_ProductVersionID", SqlDbType.Int, PVID.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectPMGandGPSyChangesProducts()
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectPMGandGPSyChangesProducts");

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }

        public DataTable SelectPMG100CharChangesFileCount(string PVIDs)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_SelectPMG100CharChangesFileCount");

            dw.CreateParameter(cmd, "@p_PVIDs", SqlDbType.VarChar, PVIDs.ToString());

            DataTable dt = dw.ExecuteCommandTable(cmd);
            return dt;
        }


        public DataTable UpdateAddEditMarketingName(string BID, string Name, string NameType, string PBID, string Series)
        {
            DataWrapper dw = new DataWrapper();
            SqlCommand cmd = dw.CreateCommand("usp_UpdateAddEditMarketingName");

            dw.CreateParameter(cmd, "@p_BID", SqlDbType.VarChar, BID.ToString());
            dw.CreateParameter(cmd, "@p_Name", SqlDbType.VarChar, Name.ToString());
            dw.CreateParameter(cmd, "@p_NameType", SqlDbType.VarChar, NameType.ToString());
            dw.CreateParameter(cmd, "@p_PBID", SqlDbType.VarChar, PBID.ToString());
            dw.CreateParameter(cmd, "@p_Series", SqlDbType.VarChar, Series.ToString());

            return dw.ExecuteCommandTable(cmd);
        }
    }
}
