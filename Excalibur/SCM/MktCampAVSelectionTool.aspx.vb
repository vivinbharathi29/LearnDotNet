Imports System.Data

Partial Class MktCampAVSelectionTool
    Inherits System.Web.UI.Page

    Public Shared Function GetSessionStateValue(ByVal id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByVal id As String, ByVal obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public Shared Property sProdVerIDs() As String
        Get
            Return (GetSessionStateValue("sProdVerIDs"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sProdVerIDs", value)
        End Set
    End Property

    Public Shared Property sProdBrandIDs() As String
        Get
            Return (GetSessionStateValue("sProdBrandIDs"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sProdBrandIDs", value)
        End Set
    End Property

    Public Shared Property sCatIDs() As String
        Get
            Return (GetSessionStateValue("sCatIDs"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sCatIDs", value)
        End Set
    End Property

    Public Shared Property sSelRegion() As String
        Get
            Return (GetSessionStateValue("sSelRegion"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sSelRegion", value)
        End Set
    End Property

    Public Shared Property sPlantID() As String
        Get
            Return (GetSessionStateValue("sPlantID"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sPlantID", value)
        End Set
    End Property

    Public Shared Property sPlantsName() As String
        Get
            Return (GetSessionStateValue("sPlantsName"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sPlantsName", value)
        End Set
    End Property

    Public Shared Property sMktCampID() As String
        Get
            Return (GetSessionStateValue("sMktCampID"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sMktCampID", value)
        End Set
    End Property

    Public Shared Property sMktCampName() As String
        Get
            Return (GetSessionStateValue("sMktCampName"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sMktCampName", value)
        End Set
    End Property

    Public Shared Property sMktCampStartDate() As String
        Get
            Return (GetSessionStateValue("sMktCampStartDate"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sMktCampStartDate", value)
        End Set
    End Property

    Public Shared Property sMktCampEndDate() As String
        Get
            Return (GetSessionStateValue("sMktCampEndDate"))
        End Get
        Set(ByVal value As String)
            AddSessionStateValue("sMktCampEndDate", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            MessageLabel.Visible = True
            lblErrorMessage.Visible = False
            lblRecCountMsg.Visible = True
            lblRecCountMsg1.Visible = True
            lblRecCountMsg2.Visible = True

            lblRecCountMsg1.Visible = True
            'lblRecCountMsg1.Text = "Test 1"

            If Not Me.Page.IsPostBack Then
                Dim dtProdBrands As DataTable
                Dim sIDQSPara As String = "" 'This will be populated by the ID Query String Parameter.
                Dim intPos As Integer = "-1"
                Dim intTotalRecs As Integer = 0
                Dim strRegion As String = ""

                sIDQSPara = Request.QueryString("ID").ToString()
                'sIDQSPara = "763:1050:Americas:1:Houston Campus:6,51:1:CTO:02-01-2011:10-31-2012"
                'sIDQSPara = "763:1050:Americas:1,2:Houston Campus,Cupertino:,6:1:Test:10-20-2012:10-31-2012"
                'sIDQSPara = "100:1462:Americas:21:ODM Inventec - Shanghai (0865)::63:Date Range Test Campaign:2/1/2012:1/27/2014"

                'Get the Values passed to this page and deal with them appropriately.
                Dim strSplitUpIDQSPara() As String = sIDQSPara.Split(":")
                For i As Integer = 0 To UBound(strSplitUpIDQSPara)
                    Select Case i
                        Case 0
                            sProdVerIDs = strSplitUpIDQSPara(i)
                        Case 1
                            sProdBrandIDs = strSplitUpIDQSPara(i)
                        Case 2
                            sSelRegion = strSplitUpIDQSPara(i)
                        Case 3
                            sPlantID = strSplitUpIDQSPara(i)
                        Case 4
                            sPlantsName = strSplitUpIDQSPara(i)
                        Case 5
                            sCatIDs = strSplitUpIDQSPara(i)
                        Case 6
                            sMktCampID = strSplitUpIDQSPara(i)
                        Case 7
                            sMktCampName = strSplitUpIDQSPara(i)
                        Case 8
                            sMktCampStartDate = strSplitUpIDQSPara(i)
                        Case 9
                            sMktCampEndDate = strSplitUpIDQSPara(i)
                    End Select
                Next i

                If (sCatIDs = "-1") Then
                    sCatIDs = ""
                End If

                Select Case sSelRegion
                    Case "Americas"
                        strRegion = "1"
                    Case "EMEA"
                        strRegion = "2"
                    Case "APJ"
                        strRegion = "3"
                End Select

                lblRegion.Text = sSelRegion

                'Now populate all of the Product Brand Names in the Header section of this web page.
                'If Not Me.Page.IsPostBack() Then
                dtProdBrands = HPQ.Excalibur.SupplyChain.SelectProductBrandNames_ByProductVersionIDAndProductBrandID(sProdVerIDs, sProdBrandIDs)

                CurrPlats.Text = ""

                If (dtProdBrands.Rows.Count > 0) Then
                    For Each dr As DataRow In dtProdBrands.Rows
                        If (CurrPlats.Text = "") Then
                            CurrPlats.Text = dr.Item("ShortProdName").ToString()
                        Else
                            CurrPlats.Text = CurrPlats.Text & "  |  " & dr.Item("ShortProdName").ToString()
                        End If
                    Next
                End If

                If sPlantsName = "" Then
                    lblRecCountMsg.Visible = False
                    lblRecCountMsg1.Visible = True
                    lblRecCountMsg1.Text = "No AV Items found for these brands."
                    lblRecCountMsg2.Visible = False
                Else
                    lblMktCampaign.Text = sMktCampName
                    lblStartDate.Text = sMktCampStartDate
                    lblEndDate.Text = sMktCampEndDate

                    lblPlants.Text = ""
                    Dim sSplitPlantsNames() As String = sPlantsName.Split(",")
                    For i As Integer = 0 To UBound(sSplitPlantsNames)
                        If (lblPlants.Text = "") Then
                            lblPlants.Text = sSplitPlantsNames(i)
                        Else
                            lblPlants.Text = lblPlants.Text & " | " & sSplitPlantsNames(i)
                        End If
                    Next i

                    Dim dtProdBrandsCustom As DataTable
                    dtProdBrands = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView_MktCampView(sProdVerIDs, sProdBrandIDs, sCatIDs, strRegion, sPlantID, sMktCampID)
                    dtProdBrandsCustom = SetupDataTableForGridView(CreateTableWithOnlyTheCorrecAVRecs(dtProdBrands, strRegion, intTotalRecs), strRegion, intTotalRecs)
                    'dtProdBrandsCustom = SetupDataTableForGridView(dtProdBrands, strRegion, intTotalRecs)

                    If dtProdBrandsCustom.Rows.Count = 0 Then
                        lblRecCountMsg.Visible = False
                        lblRecCountMsg1.Visible = True
                        lblRecCountMsg1.Text = "No AV Items found for these brands."
                        lblRecCountMsg2.Visible = False
                    Else
                        lblRecCountMsg.Visible = True
                        lblRecCountMsg1.Visible = True
                        lblRecCountMsg1.Text = ("( " & intTotalRecs & " AVs Displayed ) - ")
                        lblRecCountMsg2.Visible = True
                        lblRecCountMsg2.Text = sSelRegion & " Marketing Campaign View"
                        'lblRecCountMsg2.Text = "sProdVerIDs: " + sProdVerIDs + "\n\rsProdBrandIDs: " + sProdBrandIDs + "\n\rsCatIDs: " + sCatIDs + "\n\rstrRegion: " + strRegion + "\n\rsPlantID: " + sPlantID
                        'lblErrorMessage.Visible = True
                        'lblErrorMessage.Text = "sProdVerIDs: " + sProdVerIDs + "\n\rsProdBrandIDs: " + sProdBrandIDs + "\n\rsCatIDs: " + sCatIDs + "\n\rstrRegion: " + strRegion + "\n\rsPlantID: " + sPlantID

                        gvRegAVSelToolGrid.DataSource = dtProdBrandsCustom
                        gvRegAVSelToolGrid.DataBind()
                    End If
                End If
            End If

            MessageLabel.Visible = False
            'End If
        Catch ex As Exception
            Response.Write("<br /><br />" + ex.ToString)
        End Try
    End Sub

    Private Function AddFeatureCategorysToProdBrandDataTable(ByVal dt As DataTable) As DataTable

        Dim i As Integer = 0
        Dim strCellValue As String = ""

        For Each row As DataRow In dt.Rows
            'Column Index 0 is the Feature Category column.
            If (strCellValue <> row(0).ToString()) Then
                AddRow(row, dt)
            End If
        Next

        AddFeatureCategorysToProdBrandDataTable = dt
    End Function

    Private Sub AddRow(ByVal row As DataRow, ByVal dt As DataTable)
        Dim workRow As DataRow = dt.NewRow()

        workRow(0) = row(0).ToString() 'Feature Category
        workRow(10) = row(10).ToString()

        dt.Rows.Add(workRow)
    End Sub

    Public Enum RowColor
        LightOrange = 1
        LightBlue = 2
    End Enum

    Private Function ProcessFilter(ByVal sProdBrands As String, ByVal sCategory As String, ByVal strSelRegion As String) As Boolean
        'sProdVerIDs
        'sProdBrandIDs
        'sCatIDs
        'sSelRegion
        Try
            Dim URL As String = Nothing
            'Dim applicationRoot As String = Session("ApplicationRoot")
            'If sCategories = "" Then
            'Response.Write("<script language='javascript'> {window.close();}</script>")
            'Else
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sProdBrands & ":" & sCategory & ":" & strSelRegion))
            'Response.Write("<script language='javascript'> { window.close();}</script>")

            'End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.InnerException.ToString
        End Try
    End Function

    Private Function CreateTableWithOnlyTheCorrecAVRecs(ByVal dtProdBrands As DataTable, ByVal strRegionID As String, ByRef intTotalRecs As Integer) As DataTable

        Dim strTestText As String = ""
        Dim strFeatCat As String = ""
        Dim strAvDetail_ProdBrandID As String = ""
        Dim dr2 As Integer = 0
        Dim drRowOfFirstPlantRec As Integer
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim w As Integer = 0
        Dim m As Integer = 0
        Dim intRowsToSkipTo As Integer = 0
        Dim bolAddRec As Boolean = False
        Dim bolUnSelectedRec As Boolean = False
        Dim bolFoundPID As Boolean = False
        Dim bolAVDIDFoundInAllFilteredPlants = True
        Dim bolMktCampID As Boolean = False
        Dim strCurrAVDetalRecID As String = ""
        Dim strPlantIDsFound As String = ""
        Dim intPlantCount As Integer = 0

        Dim dt As New DataTable()
        Dim dtTestData As New DataTable()
        Dim dr As DataRow
        Dim dcMtkCampaignsAVDetailProductBrandID As New DataColumn("MktCampaigns_AVDetailProductBrandID") 'lblMktCampaigns_AVDetailProductBrandID
        Dim dcMtkCampaignsAVDetailProductBrandID2 As New DataColumn("MktCampaigns_AVDetailProductBrandID") 'lblMktCampaigns_AVDetailProductBrandID
        dt.Columns.Add(dcMtkCampaignsAVDetailProductBrandID)
        dtTestData.Columns.Add(dcMtkCampaignsAVDetailProductBrandID2)
        Dim dcMktCampaignsID As New DataColumn("MktCampaignsID") 'lblMktCampaignsID
        Dim dcMktCampaignsID2 As New DataColumn("MktCampaignsID") 'lblMktCampaignsID
        dt.Columns.Add(dcMktCampaignsID)
        dtTestData.Columns.Add(dcMktCampaignsID2)
        Dim dcFeatCat As New DataColumn("AvFeatureCategory") 'lblFeatCat
        Dim dcFeatCat2 As New DataColumn("AvFeatureCategory") 'lblFeatCat
        dt.Columns.Add(dcFeatCat)
        dtTestData.Columns.Add(dcFeatCat2)
        Dim dcSel As New DataColumn("Select") 'chkSelect
        Dim dcSel2 As New DataColumn("Select") 'chkSelect
        dcSel.DataType = GetType(Boolean)
        dt.Columns.Add(dcSel)
        dtTestData.Columns.Add(dcSel2)
        Dim dcCatField As New DataColumn("CatField") 'chkCatField
        Dim dcCatField2 As New DataColumn("CatField") 'chkCatField
        'dcCatField.DataType = GetType(Boolean)
        dt.Columns.Add(dcCatField)
        dtTestData.Columns.Add(dcCatField2)
        Dim dcGPGDesc As New DataColumn("GPGDescription") 'lblGPGDesc
        Dim dcGPGDesc2 As New DataColumn("GPGDescription") 'lblGPGDesc
        dt.Columns.Add(dcGPGDesc)
        dtTestData.Columns.Add(dcGPGDesc2)
        Dim dcAVNo As New DataColumn("AVNo") 'lblAVNo
        Dim dcAVNo2 As New DataColumn("AVNo") 'lblAVNo
        dt.Columns.Add(dcAVNo)
        dtTestData.Columns.Add(dcAVNo2)
        'Dim dcRegCPLBlind As New DataColumn("RegionalCPLBlindDate") 'lblRegCPLBlind
        Dim dcRegCPLBlind As New DataColumn("PlantStartDate") 'lblRegCPLBlind
        Dim dcRegCPLBlind2 As New DataColumn("PlantStartDate") 'lblRegCPLBlind
        dt.Columns.Add(dcRegCPLBlind)
        dtTestData.Columns.Add(dcRegCPLBlind2)
        'Dim dcRegRASDics As New DataColumn("RegionalRasDiscDate") 'lblRegRASDics
        Dim dcRegRASDics As New DataColumn("PlantEndDate") 'lblRegRASDics
        Dim dcRegRASDics2 As New DataColumn("PlantEndDate") 'lblRegRASDics
        dt.Columns.Add(dcRegRASDics)
        dtTestData.Columns.Add(dcRegRASDics2)
        Dim dcConfigRules As New DataColumn("ConfigRules") 'lblConfigRules
        Dim dcConfigRules2 As New DataColumn("ConfigRules") 'lblConfigRules
        dt.Columns.Add(dcConfigRules)
        dtTestData.Columns.Add(dcConfigRules2)
        Dim dcCategoryRules As New DataColumn("CategoryRules") 'lblCategoryRules
        Dim dcCategoryRules2 As New DataColumn("CategoryRules") 'lblCategoryRules
        dt.Columns.Add(dcCategoryRules)
        dtTestData.Columns.Add(dcCategoryRules2)
        Dim dcIDS_SKUS As New DataColumn("IdsSkus_YN") 'lblIDS_SKUS
        Dim dcIDS_SKUS2 As New DataColumn("IdsSkus_YN") 'lblIDS_SKUS
        dt.Columns.Add(dcIDS_SKUS)
        dtTestData.Columns.Add(dcIDS_SKUS2)
        Dim dcIDS_CTO As New DataColumn("IdsCto_YN") 'lblIDS_CTO
        Dim dcIDS_CTO2 As New DataColumn("IdsCto_YN") 'lblIDS_CTO
        dt.Columns.Add(dcIDS_CTO)
        dtTestData.Columns.Add(dcIDS_CTO2)
        Dim dcRCTO_SKUS As New DataColumn("RctoSkus_YN") 'lblRCTO_SKUS
        Dim dcRCTO_SKUS2 As New DataColumn("RctoSkus_YN") 'lblRCTO_SKUS
        dt.Columns.Add(dcRCTO_SKUS)
        dtTestData.Columns.Add(dcRCTO_SKUS2)
        Dim dcRCTO_CTO As New DataColumn("RctoCto_YN") 'lblRCTO_CTO
        Dim dcRCTO_CTO2 As New DataColumn("RctoCto_YN") 'lblRCTO_CTO
        dt.Columns.Add(dcRCTO_CTO)
        dtTestData.Columns.Add(dcRCTO_CTO2)
        Dim dcp_ProductBrandID As New DataColumn("ProductBrandID") 'lblProductBrandID
        Dim dcp_ProductBrandID2 As New DataColumn("ProductBrandID") 'lblProductBrandID
        dt.Columns.Add(dcp_ProductBrandID)
        dtTestData.Columns.Add(dcp_ProductBrandID2)
        Dim dcFeatureCategoryID As New DataColumn("FeatureCategoryID") 'lblAvFeatureCatID
        Dim dcFeatureCategoryID2 As New DataColumn("FeatureCategoryID") 'lblAvFeatureCatID
        dt.Columns.Add(dcFeatureCategoryID)
        dtTestData.Columns.Add(dcFeatureCategoryID2)
        Dim dcGeoID As New DataColumn("GeoID") 'lblGeoID
        Dim dcGeoID2 As New DataColumn("GeoID") 'lblGeoID
        dt.Columns.Add(dcGeoID)
        dtTestData.Columns.Add(dcGeoID2)
        Dim dcAvDetail_ProdBrandID As New DataColumn("MainAvDetProdBrandID") 'lblAvDetail_ProductBrandID
        Dim dcAvDetail_ProdBrandID2 As New DataColumn("MainAvDetProdBrandID") 'lblAvDetail_ProductBrandID
        dt.Columns.Add(dcAvDetail_ProdBrandID)
        dtTestData.Columns.Add(dcAvDetail_ProdBrandID2)
        Dim dcGSEndDt As New DataColumn("GSEndDate") 'lblGlobalSeriesConfigEOL
        Dim dcGSEndDt2 As New DataColumn("GSEndDate") 'lblGlobalSeriesConfigEOL
        dt.Columns.Add(dcGSEndDt)
        dtTestData.Columns.Add(dcGSEndDt2)
        Dim dcAvRegionalDatesID As New DataColumn("AvRegionalDatesID") 'lblAVRegionalDetailID
        Dim dcAvRegionalDatesID2 As New DataColumn("AvRegionalDatesID") 'lblAVRegionalDetailID
        dt.Columns.Add(dcAvRegionalDatesID)
        dtTestData.Columns.Add(dcAvRegionalDatesID2)
        Dim dcRCTOPlantsID As New DataColumn("RCTOPlantsID") 'lblRCTOPlantsID
        Dim dcRCTOPlantsID2 As New DataColumn("RCTOPlantsID") 'lblRCTOPlantsID
        dt.Columns.Add(dcRCTOPlantsID)
        dtTestData.Columns.Add(dcRCTOPlantsID2)
        Dim dcRCTOGEOID As New DataColumn("RCTOP_GEOID") 'lblRCTOGEOID
        Dim dcRCTOGEOID2 As New DataColumn("RCTOP_GEOID") 'lblRCTOGEOID
        dt.Columns.Add(dcRCTOGEOID)
        dtTestData.Columns.Add(dcRCTOGEOID2)
        Dim dcCheckedRec As New DataColumn("CheckedRecFlag") 'chkCheckedRec
        Dim dcCheckedRec2 As New DataColumn("CheckedRecFlag") 'chkCheckedRec
        dt.Columns.Add(dcCheckedRec)
        dtTestData.Columns.Add(dcCheckedRec2)
        Dim dcRecSelected As New DataColumn("RecSelectedFlag") 'chkRecSelected
        Dim dcRecSelected2 As New DataColumn("RecSelectedFlag") 'chkRecSelected
        dt.Columns.Add(dcRecSelected)
        dtTestData.Columns.Add(dcRecSelected2)
        Dim dcPlantCheckedRec As New DataColumn("PlantCheckedRecFlag") 'chkPlantCheckedRec
        Dim dcPlantCheckedRec2 As New DataColumn("PlantCheckedRecFlag") 'chkPlantCheckedRec
        dt.Columns.Add(dcPlantCheckedRec)
        dtTestData.Columns.Add(dcPlantCheckedRec2)
        Dim dcPlantRecSelected As New DataColumn("PlantRecSelectedFlag") 'chkPlantRecSelected
        Dim dcPlantRecSelected2 As New DataColumn("PlantRecSelectedFlag") 'chkPlantRecSelected
        dt.Columns.Add(dcPlantRecSelected)
        dtTestData.Columns.Add(dcPlantRecSelected2)
        Dim dcRegCheckedRec As New DataColumn("RegCheckedRecFlag") 'chkRegCheckedRec
        Dim dcRegCheckedRec2 As New DataColumn("RegCheckedRecFlag") 'chkRegCheckedRec
        dt.Columns.Add(dcRegCheckedRec)
        dtTestData.Columns.Add(dcRegCheckedRec2)
        Dim dcRegRecSelected As New DataColumn("RegRecSelectedFlag") 'chkRegRecSelected
        Dim dcRegRecSelected2 As New DataColumn("RegRecSelectedFlag") 'chkRegRecSelected
        dt.Columns.Add(dcRegRecSelected)
        dtTestData.Columns.Add(dcRegRecSelected2)

        Dim strPrevAVRegID As String = ""

        If (dtProdBrands.Rows.Count > 0) Then
            'strPrevAVRegID = dtProdBrands.Rows(0).Item("AvRegionalDatesID").ToString()

            For dr2 = 0 To (dtProdBrands.Rows.Count - 1)
                If (dr2 >= dtProdBrands.Rows.Count) Then
                    Exit For
                End If

                If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString()) Then
                    strTestText = dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString()

                    'If The current Av record is the same as the previous AvRegionalDatesID record then
                    'If (dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString() = strPrevAVRegID) Then
                    If ((dr2 + 1) < dtProdBrands.Rows.Count) Then
                        If (dtProdBrands.Rows((dr2 + 1)).Item("AvRegionalDatesID").ToString() = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()) Then
                            If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                                bolMktCampID = True

                                'Check if the next row is the same AVRegionalDatesID and keep looping to the next row until a different one is found.
                                For m = dr2 To (dtProdBrands.Rows.Count - 1)
                                    If (dtProdBrands.Rows(m).Item("AvRegionalDatesID").ToString() = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()) Then
                                    Else
                                        'dr2 = (m - 1) 'I used "m - 1" here instead of "m" because dr2 gets incremented at the end of this loop.
                                        Exit For
                                    End If
                                Next m

                                strPrevAVRegID = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                            Else
                                bolMktCampID = False
                                m = dr2
                                strPrevAVRegID = ""
                            End If
                        Else
                            bolMktCampID = True
                            m = dr2
                            strPrevAVRegID = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                        End If
                    Else
                        If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                            bolMktCampID = True
                            strPrevAVRegID = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                        Else
                            bolMktCampID = False
                            strPrevAVRegID = ""
                        End If
                        m = dr2
                    End If
                End If

                If (bolMktCampID = True) Then
                    dr = dtTestData.NewRow()
                    If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                        dr("MktCampaignsID") = sMktCampID
                    Else
                        dr("MktCampaignsID") = "-1"
                    End If

                    dr("MktCampaigns_AVDetailProductBrandID") = dtProdBrands.Rows(dr2).Item("MktCampaigns_AVDetailProductBrandID") 'lblMktCampaigns_AVDetailProductBrandID
                    dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                    dr("GPGDescription") = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()
                    dr("AVNo") = dtProdBrands.Rows(dr2).Item("AVNo").ToString()
                    dr("PlantStartDate") = dtProdBrands.Rows(dr2).Item("PlantStartDate").ToString()
                    dr("PlantEndDate") = dtProdBrands.Rows(dr2).Item("PlantEndDate").ToString()
                    dr("GSEndDate") = dtProdBrands.Rows(dr2).Item("GSEndDate").ToString()
                    dr("ConfigRules") = dtProdBrands.Rows(dr2).Item("ConfigRules").ToString()
                    dr("CategoryRules") = dtProdBrands.Rows(dr2).Item("CategoryRules").ToString()
                    dr("IdsSkus_YN") = dtProdBrands.Rows(dr2).Item("IdsSkus_YN").ToString()
                    dr("IdsCto_YN") = dtProdBrands.Rows(dr2).Item("IdsCto_YN").ToString()
                    dr("RctoSkus_YN") = dtProdBrands.Rows(dr2).Item("RctoSkus_YN").ToString()
                    dr("RctoCto_YN") = dtProdBrands.Rows(dr2).Item("RctoCto_YN").ToString()
                    dr("ProductBrandID") = dtProdBrands.Rows(dr2).Item("ProductBrandID").ToString()
                    dr("FeatureCategoryID") = dtProdBrands.Rows(dr2).Item("FeatureCategoryID").ToString()
                    dr("GeoID") = dtProdBrands.Rows(dr2).Item("GeoID").ToString()
                    dr("MainAvDetProdBrandID") = dtProdBrands.Rows(dr2).Item("MainAvDetProdBrandID").ToString()
                    dr("AvRegionalDatesID") = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                    dr("RCTOPlantsID") = dtProdBrands.Rows(dr2).Item("RCTOPlantsID").ToString()
                    strTestText = dtProdBrands.Rows(dr2).Item("RCTOPlantsID").ToString()
                    dr("RCTOP_GEOID") = dtProdBrands.Rows(dr2).Item("RCTOP_GEOID").ToString()
                    dr("CheckedRecFlag") = dtProdBrands.Rows(dr2).Item("CheckedRecFlag")
                    dr("RecSelectedFlag") = dtProdBrands.Rows(dr2).Item("RecSelectedFlag")
                    dr("PlantCheckedRecFlag") = dtProdBrands.Rows(dr2).Item("PlantCheckedRecFlag")
                    dr("PlantRecSelectedFlag") = dtProdBrands.Rows(dr2).Item("PlantRecSelectedFlag")
                    dr("RegCheckedRecFlag") = dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag")
                    dr("RegRecSelectedFlag") = dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag")
                    dtTestData.Rows.Add(dr)

                    If (m = dr2) Then
                    Else
                        dr2 = m - 1
                    End If
                End If
            Next dr2
        End If

        If (dtTestData.Rows.Count > 0) Then
            'drRowOfFirstPlantRec = dr2 'This is the first row number that the first plant rec was found.
            For dr2 = 0 To (dtTestData.Rows.Count - 1)
                dr2 = drRowOfFirstPlantRec 'This seems stupid but it is needed here to keep the 2 variables insync, do not remove it.

                If (dr2 >= dtTestData.Rows.Count) Then
                    Exit For
                End If

                bolAddRec = False

                bolAVDIDFoundInAllFilteredPlants = True
                'If the user is searching for marketing campaign records by one or more plants then check if each record
                'exist in the region table and all of the plant(s) the user is searching by.  It has to exist in all plants
                'the user searches by and the region, not just one of the plants and/or region.
                If (dtTestData.Rows(dr2).Item("AvRegionalDatesID").ToString() <> "") Then 'Rec exists in the region.
                    strCurrAVDetalRecID = dtTestData.Rows(dr2).Item("AvRegionalDatesID").ToString()

                    Dim sPlantsFilteredBy() As String = sPlantID.Split(",")
                    'Start looping thru the records until the CurrAVDetalRecID changes.
                    strPlantIDsFound = ""
                    bolMktCampID = False
                    Do While (dtTestData.Rows(drRowOfFirstPlantRec).Item("AvRegionalDatesID").ToString() = strCurrAVDetalRecID)
                        bolFoundPID = False
                        For j = 0 To UBound(sPlantsFilteredBy)
                            If (sPlantsFilteredBy(j) = dtTestData.Rows(drRowOfFirstPlantRec).Item("RCTOPlantsID").ToString()) Then
                                If (dtTestData.Rows(drRowOfFirstPlantRec).Item("MktCampaignsID").ToString() <> "") And Not IsDBNull(dtTestData.Rows(drRowOfFirstPlantRec).Item("MktCampaignsID").ToString()) Then
                                    If (dtTestData.Rows(drRowOfFirstPlantRec).Item("MktCampaignsID").ToString() = sMktCampID) Then
                                        bolMktCampID = True
                                    End If
                                End If

                                If (strPlantIDsFound = "") Then
                                    strPlantIDsFound = sPlantsFilteredBy(j)
                                Else
                                    strPlantIDsFound = strPlantIDsFound & "," & sPlantsFilteredBy(j)
                                End If

                                'Else
                                '    drRowOfFirstPlantRec = dr2 'This means that the current row is no longer the same AV Regional Number and the "drRowOfFirstPlantRec" variable needs to be set to the value of the "dr2" variable.
                                '    bolAVDIDFoundInAllFilteredPlants = False

                                Exit For
                            End If
                        Next j

                        drRowOfFirstPlantRec = drRowOfFirstPlantRec + 1

                        If (drRowOfFirstPlantRec >= dtTestData.Rows.Count) Then
                            Exit Do
                        End If
                        'strCurrAVDetalRecID = dtTestData.Rows(drRowOfFirstPlantRec).Item("AvRegionalDatesID").ToString()
                    Loop

                    'Compare the PlantIDs found with the PlantID Filters, if they all match, then include this AV record
                    'but only one record, not all of them.
                    Dim sPlantIDsFound() As String = strPlantIDsFound.Split(",")

                    bolAddRec = True

                    'Get the first Filtered by Plant.
                    For j = 0 To UBound(sPlantsFilteredBy)
                        bolAVDIDFoundInAllFilteredPlants = False
                        'Now check if the first Filtered by Plant has been found in the sPlantIDsFound variable.
                        For w = 0 To UBound(sPlantIDsFound)
                            If (sPlantsFilteredBy(j) = sPlantIDsFound(w)) Then
                                bolAVDIDFoundInAllFilteredPlants = True

                                Exit For
                            End If
                        Next w

                        If (bolAVDIDFoundInAllFilteredPlants = False) Then
                            bolAddRec = False
                        End If
                    Next j
                Else
                    bolAddRec = False
                    drRowOfFirstPlantRec = drRowOfFirstPlantRec + 1
                End If

                If (bolAddRec = True) Then
                    dr = dt.NewRow()

                    'dr("MktCampaignsID") = dtTestData.Rows(dr2).Item("MktCampaignsID").ToString()
                    If (bolMktCampID = True) Then
                        dr("MktCampaignsID") = sMktCampID
                    Else
                        dr("MktCampaignsID") = DBNull.Value
                    End If

                    dr("MktCampaigns_AVDetailProductBrandID") = dtTestData.Rows(dr2).Item("MktCampaigns_AVDetailProductBrandID") 'lblMktCampaigns_AVDetailProductBrandID
                    dr("AvFeatureCategory") = dtTestData.Rows(dr2).Item("AvFeatureCategory").ToString()
                    dr("GPGDescription") = dtTestData.Rows(dr2).Item("GPGDescription").ToString()
                    dr("AVNo") = dtTestData.Rows(dr2).Item("AVNo").ToString()
                    dr("PlantStartDate") = dtTestData.Rows(dr2).Item("PlantStartDate").ToString()
                    dr("PlantEndDate") = dtTestData.Rows(dr2).Item("PlantEndDate").ToString()
                    dr("GSEndDate") = dtTestData.Rows(dr2).Item("GSEndDate").ToString()
                    dr("ConfigRules") = dtTestData.Rows(dr2).Item("ConfigRules").ToString()
                    dr("CategoryRules") = dtTestData.Rows(dr2).Item("CategoryRules").ToString()
                    dr("IdsSkus_YN") = dtTestData.Rows(dr2).Item("IdsSkus_YN").ToString()
                    dr("IdsCto_YN") = dtTestData.Rows(dr2).Item("IdsCto_YN").ToString()
                    dr("RctoSkus_YN") = dtTestData.Rows(dr2).Item("RctoSkus_YN").ToString()
                    dr("RctoCto_YN") = dtTestData.Rows(dr2).Item("RctoCto_YN").ToString()
                    dr("ProductBrandID") = dtTestData.Rows(dr2).Item("ProductBrandID").ToString()
                    dr("FeatureCategoryID") = dtTestData.Rows(dr2).Item("FeatureCategoryID").ToString()
                    dr("GeoID") = dtTestData.Rows(dr2).Item("GeoID").ToString()
                    dr("MainAvDetProdBrandID") = dtTestData.Rows(dr2).Item("MainAvDetProdBrandID").ToString()
                    dr("AvRegionalDatesID") = dtTestData.Rows(dr2).Item("AvRegionalDatesID").ToString()
                    dr("RCTOPlantsID") = dtTestData.Rows(dr2).Item("RCTOPlantsID").ToString()
                    strTestText = dtTestData.Rows(dr2).Item("RCTOPlantsID").ToString()
                    dr("RCTOP_GEOID") = dtTestData.Rows(dr2).Item("RCTOP_GEOID").ToString()

                    If (dtTestData.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                        strTestText = dtTestData.Rows(dr2).Item("CheckedRecFlag")
                        If IsDBNull(dtTestData.Rows(dr2).Item("CheckedRecFlag")) Then
                            dr("CheckedRecFlag") = "N"
                        Else
                            If (dtTestData.Rows(dr2).Item("CheckedRecFlag") = "1") Then
                                dr("CheckedRecFlag") = "Y"
                            Else
                                dr("CheckedRecFlag") = "N"
                            End If
                        End If

                        strTestText = dtTestData.Rows(dr2).Item("RecSelectedFlag")
                        If IsDBNull(dtTestData.Rows(dr2).Item("RecSelectedFlag")) Then
                            dr("RecSelectedFlag") = "N"
                        Else
                            If (dtTestData.Rows(dr2).Item("RecSelectedFlag") = "1") Then
                                dr("RecSelectedFlag") = "Y"
                            Else
                                dr("RecSelectedFlag") = "N"
                            End If
                        End If
                    Else
                        dr("CheckedRecFlag") = "N"
                        dr("RecSelectedFlag") = "N"
                    End If

                    strTestText = dtTestData.Rows(dr2).Item("PlantCheckedRecFlag")
                    If IsDBNull(dtTestData.Rows(dr2).Item("PlantCheckedRecFlag")) Then
                        dr("PlantCheckedRecFlag") = "N"
                    Else
                        If (dtTestData.Rows(dr2).Item("PlantCheckedRecFlag") = "1") Then
                            dr("PlantCheckedRecFlag") = "Y"
                        Else
                            dr("PlantCheckedRecFlag") = "N"
                        End If
                    End If

                    strTestText = dtTestData.Rows(dr2).Item("PlantRecSelectedFlag")
                    If IsDBNull(dtTestData.Rows(dr2).Item("PlantRecSelectedFlag")) Then
                        dr("PlantRecSelectedFlag") = "N"
                    Else
                        If (dtTestData.Rows(dr2).Item("PlantRecSelectedFlag") = "1") Then
                            dr("PlantRecSelectedFlag") = "Y"
                        Else
                            dr("PlantRecSelectedFlag") = "N"
                        End If
                    End If

                    strTestText = dtTestData.Rows(dr2).Item("RegCheckedRecFlag")
                    If IsDBNull(dtTestData.Rows(dr2).Item("RegCheckedRecFlag")) Then
                        dr("RegCheckedRecFlag") = "N"
                    Else
                        If (dtTestData.Rows(dr2).Item("RegCheckedRecFlag") = "1") Then
                            dr("RegCheckedRecFlag") = "Y"
                        Else
                            dr("RegCheckedRecFlag") = "N"
                        End If
                    End If

                    strTestText = dtTestData.Rows(dr2).Item("RegRecSelectedFlag")
                    If IsDBNull(dtTestData.Rows(dr2).Item("RegRecSelectedFlag")) Then
                        dr("RegRecSelectedFlag") = "N"
                    Else
                        If (dtTestData.Rows(dr2).Item("RegRecSelectedFlag") = "1") Then
                            dr("RegRecSelectedFlag") = "Y"
                        Else
                            dr("RegRecSelectedFlag") = "N"
                        End If
                    End If

                    dt.Rows.Add(dr)
                End If
            Next dr2
        End If

        CreateTableWithOnlyTheCorrecAVRecs = dt

    End Function

    Private Function SetupDataTableForGridView(ByVal dtProdBrands As DataTable, ByVal strRegionID As String, ByRef intTotalRecs As Integer) As DataTable

        Dim strTestText As String = ""
        Dim strFeatCat As String = ""
        Dim strAvDetail_ProdBrandID As String = ""
        Dim dr2 As Integer = 0
        Dim drRowOfFirstPlantRec As Integer
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim intRowsToSkipTo As Integer = 0
        Dim bolAddRec As Boolean = False
        Dim bolUnSelectedRec As Boolean = False
        Dim bolAVDIDFoundInAllFilteredPlants = True
        Dim strCurrAVDetalRecID As String = ""
        Dim intPlantCount As Integer = 0
        Dim bolSelect As Boolean = False

        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim dcMtkCampaignsAVDetailProductBrandID As New DataColumn("MktCampaigns_AVDetailProductBrandID") 'lblMktCampaigns_AVDetailProductBrandID
        dt.Columns.Add(dcMtkCampaignsAVDetailProductBrandID)
        Dim dcMktCampaignsID As New DataColumn("MktCampaignsID") 'lblMktCampaignsID
        dt.Columns.Add(dcMktCampaignsID)
        Dim dcFeatCat As New DataColumn("AvFeatureCategory") 'lblFeatCat
        dt.Columns.Add(dcFeatCat)
        Dim dcSel As New DataColumn("Select") 'chkSelect
        dcSel.DataType = GetType(Boolean)
        dt.Columns.Add(dcSel)
        Dim dcCatField As New DataColumn("CatField") 'chkCatField
        'dcCatField.DataType = GetType(Boolean)
        dt.Columns.Add(dcCatField)
        Dim dcGPGDesc As New DataColumn("GPGDescription") 'lblGPGDesc
        dt.Columns.Add(dcGPGDesc)
        Dim dcAVNo As New DataColumn("AVNo") 'lblAVNo
        dt.Columns.Add(dcAVNo)
        'Dim dcRegCPLBlind As New DataColumn("RegionalCPLBlindDate") 'lblRegCPLBlind
        Dim dcRegCPLBlind As New DataColumn("PlantStartDate") 'lblRegCPLBlind
        dt.Columns.Add(dcRegCPLBlind)
        'Dim dcRegRASDics As New DataColumn("RegionalRasDiscDate") 'lblRegRASDics
        Dim dcRegRASDics As New DataColumn("PlantEndDate") 'lblRegRASDics
        dt.Columns.Add(dcRegRASDics)
        Dim dcConfigRules As New DataColumn("ConfigRules") 'lblConfigRules
        dt.Columns.Add(dcConfigRules)
        Dim dcCategoryRules As New DataColumn("CategoryRules") 'lblCategoryRules
        dt.Columns.Add(dcCategoryRules)
        Dim dcIDS_SKUS As New DataColumn("IdsSkus_YN") 'lblIDS_SKUS
        dt.Columns.Add(dcIDS_SKUS)
        Dim dcIDS_CTO As New DataColumn("IdsCto_YN") 'lblIDS_CTO
        dt.Columns.Add(dcIDS_CTO)
        Dim dcRCTO_SKUS As New DataColumn("RctoSkus_YN") 'lblRCTO_SKUS
        dt.Columns.Add(dcRCTO_SKUS)
        Dim dcRCTO_CTO As New DataColumn("RctoCto_YN") 'lblRCTO_CTO
        dt.Columns.Add(dcRCTO_CTO)
        Dim dcp_ProductBrandID As New DataColumn("ProductBrandID") 'lblProductBrandID
        dt.Columns.Add(dcp_ProductBrandID)
        Dim dcFeatureCategoryID As New DataColumn("FeatureCategoryID") 'lblAvFeatureCatID
        dt.Columns.Add(dcFeatureCategoryID)
        Dim dcGeoID As New DataColumn("GeoID") 'lblGeoID
        dt.Columns.Add(dcGeoID)
        Dim dcAvDetail_ProdBrandID As New DataColumn("MainAvDetProdBrandID") 'lblAvDetail_ProductBrandID
        dt.Columns.Add(dcAvDetail_ProdBrandID)
        Dim dcGSEndDt As New DataColumn("GSEndDate") 'lblGlobalSeriesConfigEOL
        dt.Columns.Add(dcGSEndDt)
        Dim dcAvRegionalDatesID As New DataColumn("AvRegionalDatesID") 'lblAVRegionalDetailID
        dt.Columns.Add(dcAvRegionalDatesID)
        Dim dcRCTOPlantsID As New DataColumn("RCTOPlantsID") 'lblRCTOPlantsID
        dt.Columns.Add(dcRCTOPlantsID)
        Dim dcRCTOGEOID As New DataColumn("RCTOP_GEOID") 'lblRCTOGEOID
        dt.Columns.Add(dcRCTOGEOID)
        Dim dcCheckedRec As New DataColumn("CheckedRecFlag") 'chkCheckedRec
        dt.Columns.Add(dcCheckedRec)
        Dim dcRecSelected As New DataColumn("RecSelectedFlag") 'chkRecSelected
        dt.Columns.Add(dcRecSelected)
        Dim dcPlantCheckedRec As New DataColumn("PlantCheckedRecFlag") 'chkPlantCheckedRec
        dt.Columns.Add(dcPlantCheckedRec)
        Dim dcPlantRecSelected As New DataColumn("PlantRecSelectedFlag") 'chkPlantRecSelected
        dt.Columns.Add(dcPlantRecSelected)
        Dim dcRegCheckedRec As New DataColumn("RegCheckedRecFlag") 'chkRegCheckedRec
        dt.Columns.Add(dcRegCheckedRec)
        Dim dcRegRecSelected As New DataColumn("RegRecSelectedFlag") 'chkRegRecSelected
        dt.Columns.Add(dcRegRecSelected)

        If (dtProdBrands.Rows.Count > 0) Then
            drRowOfFirstPlantRec = dr2 'This is the first row number that the first plant rec was found.
            For dr2 = 0 To (dtProdBrands.Rows.Count - 1)
                'bolAVDIDFoundInAllFilteredPlants = True
                ''If the user is searching for marketing campaign records by one or more plants then check if each record
                ''exist in the region table and all of the plant(s) the user is searching by.  It has to exist in all plants
                ''the user searches by and the region, not just one of the plants and/or region.
                'If (dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString() <> "") Then 'Rec exists in the region.
                '    strCurrAVDetalRecID = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                '    'Start looping thru the records until the CurrAVDetalRecID changes.
                '    Do Until (dtProdBrands.Rows(drRowOfFirstPlantRec).Item("AvRegionalDatesID").ToString() = strCurrAVDetalRecID)
                '        Dim sPlantsFilteredBy() As String = sPlantID.Split(",")
                '        For j = 0 To UBound(sPlantsFilteredBy)
                '            If (sPlantsFilteredBy(j) <> dtProdBrands.Rows(drRowOfFirstPlantRec).Item("AvRegionalDatesID").ToString()) Then
                '                drRowOfFirstPlantRec = dr2 'This means that the current row is no longer the same AV Regional Number and the "drRowOfFirstPlantRec" variable needs to be set to the value of the "dr2" variable.
                '                bolAVDIDFoundInAllFilteredPlants = False

                '                Exit For
                '            End If
                '        Next j

                '        drRowOfFirstPlantRec = drRowOfFirstPlantRec + 1
                '    Loop
                'Else
                '    bolAVDIDFoundInAllFilteredPlants = False
                'End If

                'Row = gvRegAVSelToolGrid.Rows(i)
                'For Each dr2 As DataRow In dtProdBrands.Rows
                'Check if the region in this current record is the same as the region the user is filtering by.
                If (dtProdBrands.Rows(dr2).Item("RCTOP_GEOID").ToString() = strRegionID) Then
                    'If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                    'If (bolLastRecAdded = True) Then
                    '    If (strAvDetail_ProdBrandID = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()) Then
                    '    End If
                    '    strAvDetail_ProdBrandID = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()
                    '    bolLastRecAdded = True

                    '    bolAddRec = True
                    'Else
                    'If (dtProdBrands.Rows(dr2).Item("GPGDescription").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()) Then
                    bolAddRec = True
                    bolUnSelectedRec = False 'This is false because it has been selected.
                    'Else
                    '    bolAddRec = False
                    'End If
                    'End If

                    'Now loop thru the next rows to see if any of them have the same "GPGDescription" field
                    'as the current record.  If they do then Add a value to the intRowsToSkipTo integer variable to
                    'skip all matching records and go straight to the next record.
                    intRowsToSkipTo = 0
                    i = dr2
                    If ((dr2 + 1) < dtProdBrands.Rows.Count) Then
                        Do Until (dtProdBrands.Rows(i + 1).Item("GPGDescription").ToString() <> dtProdBrands.Rows(i).Item("GPGDescription").ToString())
                            'Check if the next record is still the same "p_ProductBrandID".  If it is then skip it because
                            'it is a different region.
                            'If (dtProdBrands.Rows(i + 1).Item("p_ProductBrandID").ToString() = dtProdBrands.Rows(i).Item("p_ProductBrandID").ToString()) Then
                            '    intRowsToSkipTo = intRowsToSkipTo + 1
                            'Else 'Else, it is the not the same "p_ProductBrandID" and it needs to be processed next.
                            '    Exit Do
                            'End If

                            i = i + 1
                            If ((i + 1) >= dtProdBrands.Rows.Count) Then
                                Exit Do
                            End If
                        Loop
                    End If
                Else 'The current record is not the same as the region the user filtered by. 'None of the records in this section of the if statement have been selected.
                    bolUnSelectedRec = True 'This is true because it hasn't been selected by the user yet.

                    'An empty "AvDetail_ProductBrandID" field can only occur once and therefore needs to be added.
                    If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = "") Or IsDBNull(dtProdBrands.Rows(dr2).Item("GeoID").ToString()) Then
                        bolAddRec = True
                    Else
                        bolAddRec = True
                        'Check if the next record is still the same "GPGDescription", if not add it tither forthe.
                        If ((dr2 + 1) < dtProdBrands.Rows.Count) Then
                            If (dtProdBrands.Rows(dr2 + 1).Item("GPGDescription").ToString() = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()) Then

                                'Now loop thru the next rows to see if any of them have the same "GPGDescription" field
                                'as the current record.  If they do then Add a value to the intRowsToSkipTo integer variable to
                                'skip all matching records and go straight to the next record.
                                intRowsToSkipTo = 0
                                i = dr2
                                If ((dr2 + 1) < dtProdBrands.Rows.Count) Then
                                    Do Until (dtProdBrands.Rows(i + 1).Item("GPGDescription").ToString() <> dtProdBrands.Rows(i).Item("GPGDescription").ToString())
                                        'Check if the next record is still the same "p_ProductBrandID".  If it is then skip it because
                                        'it is a different region.
                                        If (dtProdBrands.Rows(i + 1).Item("p_ProductBrandID").ToString() = dtProdBrands.Rows(i).Item("p_ProductBrandID").ToString()) Then
                                            'Check if this next record is the same region that the user is filtering by.  If it is then
                                            'set the dr2 variable to the "i + 1" value and keep looping to see if the next
                                            '"GPGDescription" is the same.
                                            If (dtProdBrands.Rows(i + 1).Item("GeoID").ToString() = strRegionID) Then
                                                bolUnSelectedRec = False
                                                dr2 = (i + 1) 'Add this row instead of the current one.
                                            Else
                                                intRowsToSkipTo = intRowsToSkipTo + 1
                                            End If
                                        Else 'Else, it is the not the same "p_ProductBrandID" and it needs to be processed next.
                                            Exit Do
                                        End If

                                        i = i + 1
                                        If ((i + 1) >= dtProdBrands.Rows.Count) Then
                                            Exit Do
                                        End If
                                    Loop
                                End If
                            Else
                                intRowsToSkipTo = 0
                            End If
                        End If
                    End If
                End If

                If (bolAddRec = True) Then
                    intTotalRecs = intTotalRecs + 1
                    dr = dt.NewRow()

                    If (strFeatCat <> dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()) Then
                        'dr("MktCampaignsID") = ""
                        dr("MktCampaignsID") = sMktCampID
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                        dr("ConfigRules") = dtProdBrands.Rows(dr2).Item("CategoryRules").ToString()
                        dr("Select") = False
                        dr("CatField") = True

                        strFeatCat = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        dt.Rows.Add(dr)
                    End If

                    'If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString()) Then
                    dr = dt.NewRow()
                    'dr("MktCampaignsID") = dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString()
                    strTestText = dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString()
                    dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                    ''If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                    'If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                    '    'If (dtProdBrands.Rows(dr2).Item("GPGDescription").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()) Then
                    '    dr("Select") = True
                    '    'Else
                    '    '    dr("Select") = False
                    '    'End If
                    'Else
                    '    dr("Select") = False
                    'End If
                    'strTestText = dtProdBrands.Rows(dr2).Item("MktCampaignsID")
                    If (dtProdBrands.Rows(dr2).Item("MktCampaignsID").ToString() = sMktCampID) Then
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("RecSelectedFlag")) Then
                            dr("Select") = False
                            bolSelect = False
                        Else
                            If (dtProdBrands.Rows(dr2).Item("RecSelectedFlag").ToString() = "N") Or (dtProdBrands.Rows(dr2).Item("RecSelectedFlag").ToString() = "0") Then
                                dr("Select") = False
                                bolSelect = False
                            ElseIf (dtProdBrands.Rows(dr2).Item("RecSelectedFlag").ToString() = "Y") Or (dtProdBrands.Rows(dr2).Item("RecSelectedFlag").ToString() = "1") Then
                                dr("Select") = True
                                bolSelect = True
                            End If
                        End If
                    Else
                        dr("Select") = False
                        bolSelect = False
                    End If

                    dr("CatField") = False

                    dr("MktCampaignsID") = sMktCampID
                    dr("MktCampaigns_AVDetailProductBrandID") = dtProdBrands.Rows(dr2).Item("MktCampaigns_AVDetailProductBrandID")
                    dr("GPGDescription") = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()
                    dr("AVNo") = dtProdBrands.Rows(dr2).Item("AVNo").ToString()
                    dr("PlantStartDate") = dtProdBrands.Rows(dr2).Item("PlantStartDate").ToString()
                    dr("PlantEndDate") = dtProdBrands.Rows(dr2).Item("PlantEndDate").ToString()
                    dr("GSEndDate") = dtProdBrands.Rows(dr2).Item("GSEndDate").ToString()
                    dr("ConfigRules") = dtProdBrands.Rows(dr2).Item("ConfigRules").ToString()
                    dr("CategoryRules") = dtProdBrands.Rows(dr2).Item("CategoryRules").ToString()
                    dr("IdsSkus_YN") = dtProdBrands.Rows(dr2).Item("IdsSkus_YN").ToString()
                    dr("IdsCto_YN") = dtProdBrands.Rows(dr2).Item("IdsCto_YN").ToString()
                    dr("RctoSkus_YN") = dtProdBrands.Rows(dr2).Item("RctoSkus_YN").ToString()
                    dr("RctoCto_YN") = dtProdBrands.Rows(dr2).Item("RctoCto_YN").ToString()
                    dr("ProductBrandID") = dtProdBrands.Rows(dr2).Item("ProductBrandID").ToString()
                    dr("FeatureCategoryID") = dtProdBrands.Rows(dr2).Item("FeatureCategoryID").ToString()
                    dr("GeoID") = dtProdBrands.Rows(dr2).Item("GeoID").ToString()
                    dr("MainAvDetProdBrandID") = dtProdBrands.Rows(dr2).Item("MainAvDetProdBrandID").ToString()
                    dr("AvRegionalDatesID") = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                    strTestText = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                    dr("RCTOPlantsID") = dtProdBrands.Rows(dr2).Item("RCTOPlantsID").ToString()
                    strTestText = dtProdBrands.Rows(dr2).Item("RCTOPlantsID").ToString()
                    dr("RCTOP_GEOID") = dtProdBrands.Rows(dr2).Item("RCTOP_GEOID").ToString()
                    'If (bolSelect = True) Then
                    dr("CheckedRecFlag") = dtProdBrands.Rows(dr2).Item("CheckedRecFlag").ToString()
                    dr("RecSelectedFlag") = dtProdBrands.Rows(dr2).Item("RecSelectedFlag").ToString()
                    'Else
                    '    dr("CheckedRecFlag") = "N"
                    '    dr("RecSelectedFlag") = "N"
                    'End If
                    dr("PlantCheckedRecFlag") = dtProdBrands.Rows(dr2).Item("PlantCheckedRecFlag").ToString()
                    dr("PlantRecSelectedFlag") = dtProdBrands.Rows(dr2).Item("PlantRecSelectedFlag").ToString()
                    dr("RegCheckedRecFlag") = dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag").ToString()
                    dr("RegRecSelectedFlag") = dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag").ToString()



                    'strTestText = dtProdBrands.Rows(dr2).Item("CheckedRecFlag")
                    'If (bolSelect = True) Then
                    '    If (dtProdBrands.Rows(dr2).Item("CheckedRecFlag") = "1") Then
                    '        dr("CheckedRecFlag") = "Y"
                    '    Else
                    '        dr("CheckedRecFlag") = "N"
                    '    End If
                    'Else
                    '    dr("CheckedRecFlag") = "N"
                    'End If

                    'strTestText = dtProdBrands.Rows(dr2).Item("RecSelectedFlag")
                    'If (bolSelect = True) Then
                    '    If (dtProdBrands.Rows(dr2).Item("RecSelectedFlag") = "1") Then
                    '        dr("RecSelectedFlag") = "Y"
                    '    Else
                    '        dr("RecSelectedFlag") = "N"
                    '    End If
                    'Else
                    '    dr("RecSelectedFlag") = "N"
                    'End If

                    'strTestText = dtProdBrands.Rows(dr2).Item("PlantCheckedRecFlag")
                    'If (dtProdBrands.Rows(dr2).Item("PlantCheckedRecFlag") = "1") Then
                    '    dr("PlantCheckedRecFlag") = "Y"
                    'Else
                    '    dr("PlantCheckedRecFlag") = "N"
                    'End If

                    'strTestText = dtProdBrands.Rows(dr2).Item("PlantRecSelectedFlag")
                    'If (dtProdBrands.Rows(dr2).Item("PlantRecSelectedFlag") = "1") Then
                    '    dr("PlantRecSelectedFlag") = "Y"
                    'Else
                    '    dr("PlantRecSelectedFlag") = "N"
                    'End If

                    'strTestText = dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag")
                    'If (dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag") = "1") Then
                    '    dr("RegCheckedRecFlag") = "Y"
                    'Else
                    '    dr("RegCheckedRecFlag") = "N"
                    'End If

                    'strTestText = dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag")
                    'If (dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag") = "1") Then
                    '    dr("RegRecSelectedFlag") = "Y"
                    'Else
                    '    dr("RegRecSelectedFlag") = "N"
                    'End If


                    dt.Rows.Add(dr)
                    'End If

                    dr2 = dr2 + intRowsToSkipTo
                End If
            Next dr2
        End If

        SetupDataTableForGridView = dt

    End Function

    Protected Sub gvRegAVSelToolGrid_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvRegAVSelToolGrid.DataBound
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        If (intRowCount = 0) Then Exit Sub
        Dim i As Integer
        Dim iCurrRowColor As Integer
        Dim strCurrFeatCat As String
        strCurrFeatCat = ""
        Dim strCurrGPGDesc As String
        strCurrGPGDesc = ""
        'strCurrFeatCat = gvRegAVSelToolGrid.Rows(0).FindControl("lblFeatCat").ToString()

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            'Dim k As Integer
            'k = gvRegAVSelToolGrid.DataKeys(i).Value

            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim MtkCampaignsAVDetailProductBrandID As Label = Row.FindControl("lblMktCampaigns_AVDetailProductBrandID")
            Dim MktCampaignsID As Label = Row.FindControl("lblMktCampaignsID") '-0
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
            Dim MktCampCPLBlindLabel As Label = Row.FindControl("lblMktCampCPLBlind") '-2
            Dim MktCampCPLBlindText As TextBox = Row.FindControl("txtMktCampCPLBlind") '-2
            Dim MktCampRASDiscLabel As Label = Row.FindControl("lblMktCampRASDisc") '-3
            Dim MktCampRASDiscText As TextBox = Row.FindControl("txtMktCampRASDisc") '-3
            Dim GPGDesc As Label = Row.FindControl("lblGPGDesc") '-4
            Dim ProdName As Label = Row.FindControl("lblProdName") '-5
            Dim AVNo As Label = Row.FindControl("lblAVNo") '-6
            Dim GlobalCPLBlind As Label = Row.FindControl("lblGlobalCPLBlind") '-7
            Dim GlobalRASDisc As Label = Row.FindControl("lblGlobalRASDisc") '-8
            Dim ConfigRules As Label = Row.FindControl("lblConfigRules") '-9
            Dim IDS_SKUS As Label = Row.FindControl("lblIDS_SKUS") '-10
            Dim IDS_CTO As Label = Row.FindControl("lblIDS_CTO") '-11
            Dim RCTO_SKUS As Label = Row.FindControl("lblRCTO_SKUS") '-12
            Dim RCTO_CTO As Label = Row.FindControl("lblRCTO_CTO") '-13
            Dim p_ProductBrandID As Label = Row.FindControl("p_ProductBrandID") '-14
            Dim FeatureCategoryID As Label = Row.FindControl("FeatureCategoryID") '-15
            Dim GeoID As Label = Row.FindControl("GeoID") '16
            Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
            Dim AVRegionalDetailID As Label = Row.FindControl("lblAVRegionalDetailID") '18
            Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
            Dim PlantCheckedRec As Label = Row.FindControl("chkPlantCheckedRec") '21
            Dim PlantRecSelected As Label = Row.FindControl("chkPlantRecSelected") '22
            Dim RegCheckedRec As Label = Row.FindControl("chkRegCheckedRec") '23
            Dim RegRecSelected As Label = Row.FindControl("chkRegRecSelected") '24

            If (FeatureCat.Text <> strCurrFeatCat) Then
                'Row.Controls.Add(
                strCurrFeatCat = FeatureCat.Text
                ChangeRowColor(iCurrRowColor, i, True)
                chkSel.Visible = True
                CheckedRec.Visible = False
                RecSelected.Visible = False
                PlantCheckedRec.Visible = False
                PlantRecSelected.Visible = False
                RegCheckedRec.Visible = False
                RegRecSelected.Visible = False
                'For j = 5 To 6
                '    Row.Cells.Remove(Row.Cells(j))
                'Next j
                'Row.Cells(2).ColumnSpan = 3
                chkCatField.Text = "True"
            Else
                chkCatField.Text = "False"

                If (strCurrGPGDesc = "") Then
                    strCurrGPGDesc = GPGDesc.Text

                    iCurrRowColor = RowColor.LightOrange
                    ChangeRowColor(iCurrRowColor, i, False)
                Else
                    If (strCurrGPGDesc.Trim <> GPGDesc.Text.Trim) Then
                        If (iCurrRowColor = RowColor.LightBlue) Then
                            iCurrRowColor = RowColor.LightOrange
                            ChangeRowColor(iCurrRowColor, i, False)
                        ElseIf (iCurrRowColor = RowColor.LightOrange) Then
                            iCurrRowColor = RowColor.LightBlue
                            ChangeRowColor(iCurrRowColor, i, False)
                        End If

                        strCurrGPGDesc = GPGDesc.Text
                    Else
                        ChangeRowColor(iCurrRowColor, i, False)
                    End If
                End If

                FeatureCat.Visible = False

                If (chkSel.Checked = False) Then
                    If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                        Row.CssClass = "O"
                    End If
                    'ElseIf (chkSel.Checked = True) Then
                End If

                If (RegRecSelected.Text = "N") Or (PlantRecSelected.Text = "N") Then
                    Row.CssClass = "O"
                    chkSel.Visible = False
                End If
            End If
        Next

    End Sub

    Protected Sub ChangeRowColor(ByVal intColor As Integer, ByVal i As Integer, _
                                 ByVal bolAllColors As Boolean)

        Dim j As Integer
        Dim k As Integer

        If (bolAllColors = False) Then
            j = 6
        Else
            j = 0
        End If

        For k = j To 17
            If (bolAllColors = True) Then
                gvRegAVSelToolGrid.Rows(i).Cells(k).BackColor = Drawing.Color.Yellow
            Else
                If (intColor = 1) Then 'Light Orange
                    gvRegAVSelToolGrid.Rows(i).Cells(k).BackColor = Drawing.Color.FromArgb(218, 238, 243)
                ElseIf (intColor = 2) Then 'Light Blue
                    gvRegAVSelToolGrid.Rows(i).Cells(k).BackColor = Drawing.Color.FromArgb(253, 233, 217)
                End If
            End If
        Next

        If (bolAllColors = False) Then
            'Make the Feature Category Text Appear Invisible.
            gvRegAVSelToolGrid.Rows(i).Cells(2).ForeColor = Drawing.Color.White
        End If
    End Sub

    Protected Sub chkSelect_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strTest As String = "" 'sender.ToString()
        Dim chk As CheckBox = sender
        strTest = chk.Parent.Parent.ToString()
        Dim Row As GridViewRow = chk.NamingContainer
        Dim bolFeatCat As Boolean = False

        Dim i As Integer
        Dim j As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim strRegion As String = ""

        Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
        Dim MtkCampaignsAVDetailProductBrandID As Label = Row.FindControl("lblMktCampaigns_AVDetailProductBrandID")
        Dim MktCampaignsID As Label = Row.FindControl("lblMktCampaignsID") '-0
        Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
        Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
        Dim GPGDesc As Label = Row.FindControl("lblGPGDesc") '-4
        Dim ProdName As Label = Row.FindControl("lblProdName") '-5
        Dim AVNo As Label = Row.FindControl("lblAVNo") '-6
        Dim RegCPLBlind As Label = Row.FindControl("lblRegCPLBlind") '-7
        Dim RegRASDisc As Label = Row.FindControl("lblRegRASDics") '-8
        Dim ConfigRules As Label = Row.FindControl("lblConfigRules") '-9
        Dim IDS_SKUS As Label = Row.FindControl("lblIDS_SKUS") '-10
        Dim IDS_CTO As Label = Row.FindControl("lblIDS_CTO") '-11
        Dim RCTO_SKUS As Label = Row.FindControl("lblRCTO_SKUS") '-12
        Dim RCTO_CTO As Label = Row.FindControl("lblRCTO_CTO") '-13
        Dim p_ProductBrandID As Label = Row.FindControl("p_ProductBrandID") '-14
        Dim FeatureCategoryID As Label = Row.FindControl("FeatureCategoryID") '-15
        Dim GeoID As Label = Row.FindControl("GeoID") '16
        Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
        Dim AVRegionalDetailID As Label = Row.FindControl("lblAVRegionalDetailID") '18
        Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
        Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
        Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
        Dim PlantCheckedRec As Label = Row.FindControl("chkPlantCheckedRec") '21
        Dim PlantRecSelected As Label = Row.FindControl("chkPlantRecSelected") '22
        Dim RegCheckedRec As Label = Row.FindControl("chkRegCheckedRec") '23
        Dim RegRecSelected As Label = Row.FindControl("chkRegRecSelected") '24

        Select Case sSelRegion
            Case "Americas"
                strRegion = "1"
            Case "EMEA"
                strRegion = "2"
            Case "APJ"
                strRegion = "3"
        End Select

        strCurrFeatCat = FeatureCat.Text
        If (chkCatField.Text = "True") Then
            bolFeatCat = True 'This is a feature category row.


            'Dim Row As GridViewRow = chk.NamingContainer

            Dim m As Integer = (Row.RowIndex + 1)
            Dim l As Integer = 0
            Dim bolCatFeatChecked As Boolean = chkSel.Checked

            For l = m To (intRowCount - 1)
                Dim Row2 As GridViewRow
                Row2 = gvRegAVSelToolGrid.Rows(l)

                Dim FeatureCat2 As Label = Row2.FindControl("lblFeatCat")
                Dim MtkCampaignsAVDetailProductBrandID2 As Label = Row2.FindControl("lblMktCampaigns_AVDetailProductBrandID")
                Dim MktCampaignsID2 As Label = Row2.FindControl("lblMktCampaignsID") '-0
                Dim chkSel2 As CheckBox = Row2.FindControl("chkSelect") '-1
                Dim chkCatField2 As Label = Row2.FindControl("chkCatField") '-1.5
                Dim GPGDesc2 As Label = Row2.FindControl("lblGPGDesc") '-4
                Dim ProdName2 As Label = Row2.FindControl("lblProdName") '-5
                Dim AVNo2 As Label = Row2.FindControl("lblAVNo") '-6
                strTest = AVNo.Text
                Dim RegCPLBlindLabel2 As Label = Row2.FindControl("lblRegCPLBlind") '-7
                Dim RegRASDiscLabel2 As Label = Row2.FindControl("lblRegRASDics") '-8
                Dim ConfigRules2 As Label = Row2.FindControl("lblConfigRules") '-9
                Dim IDS_SKUS2 As Label = Row2.FindControl("lblIDS_SKUS") '-10
                Dim IDS_CTO2 As Label = Row2.FindControl("lblIDS_CTO") '-11
                Dim RCTO_SKUS2 As Label = Row2.FindControl("lblRCTO_SKUS") '-12
                Dim RCTO_CTO2 As Label = Row2.FindControl("lblRCTO_CTO") '-13
                Dim p_ProductBrandID2 As Label = Row2.FindControl("lblProductBrandID") '-14
                Dim FeatureCategoryID2 As Label = Row2.FindControl("lblAvFeatureCatID") '-15
                Dim GeoID2 As Label = Row2.FindControl("lblGeoID") '16
                Dim AvDetail_ProdBrandID2 As Label = Row2.FindControl("lblAvDetail_ProductBrandID") '17
                Dim AVRegionalDetailID2 As Label = Row2.FindControl("lblAVRegionalDetailID") '18
                Dim RCTOPlantsID2 As Label = Row2.FindControl("lblRCTOPlantsID")
                Dim CheckedRec2 As Label = Row2.FindControl("chkCheckedRec") '19
                Dim RecSelected2 As Label = Row2.FindControl("chkRecSelected") '20
                Dim PlantCheckedRec2 As Label = Row2.FindControl("chkPlantCheckedRec") '21
                Dim PlantRecSelected2 As Label = Row2.FindControl("chkPlantRecSelected") '22
                Dim RegCheckedRec2 As Label = Row2.FindControl("chkRegCheckedRec") '23
                Dim RegRecSelected2 As Label = Row2.FindControl("chkRegRecSelected")

                strTest = chk.Parent.Parent.ToString()

                If (strCurrFeatCat = FeatureCat2.Text) Then
                    If (chkSel.Checked = False) Then
                        If (chkSel2.Visible = True) Then
                            chkSel2.Checked = False

                            If (CheckedRec2.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                                Row2.CssClass = "O"

                                'UpdateAvRegionalDates
                                HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID2.Text, _
                                                                             MktCampaignsID2.Text, AVRegionalDetailID2.Text, _
                                                                             RCTOPlantsID2.Text, "1", "0")

                                RecSelected2.Text = "N"
                            End If

                            RegCPLBlindLabel2.BackColor = GPGDesc.BackColor
                            RegRASDiscLabel2.BackColor = GPGDesc.BackColor
                        End If
                    Else
                        If (chkSel2.Visible = True) Then
                            chkSel2.Checked = True

                            If (CheckedRec2.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                                HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID2.Text, _
                                                                             MktCampaignsID2.Text, AVRegionalDetailID2.Text, _
                                                                             RCTOPlantsID2.Text, "1", "1")

                                RecSelected2.Text = "Y"
                                Row2.CssClass = ""
                            Else
                                HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandInsert(MktCampaignsID2.Text, AVRegionalDetailID2.Text, _
                                                                                            RCTOPlantsID2.Text)

                                CheckedRec2.Text = "Y"
                                RecSelected2.Text = "Y"
                            End If
                        End If
                    End If
                End If

                'chkSel2.Visible = True
            Next l
        Else
            If (chkSel.Checked = False) Then
                If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                    Row.CssClass = "O"

                    'UpdateAvRegionalDates
                    HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID.Text, _
                                                                 MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                 RCTOPlantsID.Text, "1", "0")

                    RecSelected.Text = "N"
                End If

                RegCPLBlind.BackColor = GPGDesc.BackColor
                RegRASDisc.BackColor = GPGDesc.BackColor
            Else
                If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                    HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID.Text, _
                                                                 MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                 RCTOPlantsID.Text, "1", "1")

                    RecSelected.Text = "Y"
                    Row.CssClass = ""
                Else
                    Dim strTestText As String
                    strTestText = MktCampaignsID.Text
                    strTestText = AvDetail_ProdBrandID.Text
                    strTestText = RCTOPlantsID.Text
                    HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandInsert(MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                                RCTOPlantsID.Text)

                    CheckedRec.Text = "Y"
                    RecSelected.Text = "Y"
                End If
            End If
        End If
        'Next i
    End Sub

    Protected Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged

        Dim i As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim MtkCampaignsAVDetailProductBrandID As Label = Row.FindControl("lblMktCampaigns_AVDetailProductBrandID")
            Dim MktCampaignsID As Label = Row.FindControl("lblMktCampaignsID") '-0
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
            Dim GPGDesc As Label = Row.FindControl("lblGPGDesc") '-4
            Dim ProdName As Label = Row.FindControl("lblProdName") '-5
            Dim AVNo As Label = Row.FindControl("lblAVNo") '-6
            Dim RegCPLBlind As Label = Row.FindControl("lblRegCPLBlind") '-7
            Dim RegRASDisc As Label = Row.FindControl("lblRegRASDics") '-8
            Dim ConfigRules As Label = Row.FindControl("lblConfigRules") '-9
            Dim IDS_SKUS As Label = Row.FindControl("lblIDS_SKUS") '-10
            Dim IDS_CTO As Label = Row.FindControl("lblIDS_CTO") '-11
            Dim RCTO_SKUS As Label = Row.FindControl("lblRCTO_SKUS") '-12
            Dim RCTO_CTO As Label = Row.FindControl("lblRCTO_CTO") '-13
            Dim p_ProductBrandID As Label = Row.FindControl("p_ProductBrandID") '-14
            Dim FeatureCategoryID As Label = Row.FindControl("FeatureCategoryID") '-15
            Dim GeoID As Label = Row.FindControl("GeoID") '16
            Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
            Dim AVRegionalDetailID As Label = Row.FindControl("lblAVRegionalDetailID") '18
            Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20

            If (strCurrFeatCat <> FeatureCat.Text) Then
                strCurrFeatCat = FeatureCat.Text
                chkSel.Checked = chkAll.Checked
            Else
                If (chkSel.Visible = True) Then
                    If (chkAll.Checked = False) Then
                        chkSel.Checked = False

                        If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                            Row.CssClass = "O"

                            HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID.Text, _
                                                                         MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                         RCTOPlantsID.Text, "1", "0")

                            RecSelected.Text = "N"
                        End If
                    Else
                        chkSel.Checked = True

                        If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                            HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandUpdate(MtkCampaignsAVDetailProductBrandID.Text, _
                                                                         MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                         RCTOPlantsID.Text, "1", "1")

                            RecSelected.Text = "Y"
                            Row.CssClass = ""
                        Else
                            HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandInsert(MktCampaignsID.Text, AVRegionalDetailID.Text, _
                                                                                        RCTOPlantsID.Text)

                            CheckedRec.Text = "Y"
                            RecSelected.Text = "Y"
                        End If
                    End If
                End If
            End If
        Next i
    End Sub
End Class
