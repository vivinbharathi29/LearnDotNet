Imports System.Data

Partial Class SCM_PlantAVSelectorTool
    Inherits System.Web.UI.Page
    'Dim dw As HPQ.Excalibur.SupplyChain = New HPQ.Excalibur.SupplyChain()
    'Public sProdVerIDs As String = ""
    'Public sProdBrandIDs As String = ""
    'Public sCatIDs As String = ""
    'Public sSelRegion As String = "1"

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
                'sIDQSPara = "763:1050:Americas:1:Houston Campus:-1,6,51"
                'sIDQSPara = "763:1050:Americas:1,2,3:Houston Campus,Cupertino,Oregon:,6"

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
                    dtProdBrands = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView_PlantView(sProdVerIDs, sProdBrandIDs, sCatIDs, strRegion, sPlantID)
                    dtProdBrandsCustom = SetupDataTableForGridView(dtProdBrands, strRegion, intTotalRecs)

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
                        lblRecCountMsg2.Text = sSelRegion & " Regional Plant View"
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

    Private Function SetupDataTableForGridView(ByVal dtProdBrands As DataTable, ByVal strRegionID As String, ByRef intTotalRecs As Integer) As DataTable

        Dim strTestText As String = ""
        Dim strFeatCat As String = ""
        Dim strAvDetail_ProdBrandID As String = ""
        Dim dr2 As Integer = 0
        Dim i As Integer = 0
        Dim intRowsToSkipTo As Integer = 0
        Dim bolAddRec As Boolean = False
        Dim bolUnSelectedRec As Boolean = False

        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim dcAVDetailID As New DataColumn("AvDetailID") 'lblAVDetailID
        dt.Columns.Add(dcAVDetailID)
        Dim dcRCTOPlants_AVDetailProductBrand As New DataColumn("RCTOPlants_AVDetailProductBrand") 'lblRCTOPlants_AVDetailProductBrand
        dt.Columns.Add(dcRCTOPlants_AVDetailProductBrand)
        Dim dcFeatCat As New DataColumn("AvFeatureCategory") 'lblFeatCat
        dt.Columns.Add(dcFeatCat)
        Dim dcSel As New DataColumn("Select") 'chkSelect
        dcSel.DataType = GetType(Boolean)
        dt.Columns.Add(dcSel)
        Dim dcCatField As New DataColumn("CatField") 'chkCatField
        'dcCatField.DataType = GetType(Boolean)
        dt.Columns.Add(dcCatField)
        Dim dcPlantCPLBlind As New DataColumn("PlantStartDate") 'lblPlantCPLBlind
        dt.Columns.Add(dcPlantCPLBlind)
        Dim dcPlantRASDics As New DataColumn("PlantEndDate") 'lblPlantRASDics
        dt.Columns.Add(dcPlantRASDics)
        Dim dcGPGDesc As New DataColumn("GPGDescription") 'lblGPGDesc
        dt.Columns.Add(dcGPGDesc)
        Dim dcPlantName As New DataColumn("PlantName") 'lblPlantName
        dt.Columns.Add(dcPlantName)
        Dim dcAVNo As New DataColumn("AVNo") 'lblAVNo
        dt.Columns.Add(dcAVNo)
        Dim dcRegCPLBlind As New DataColumn("RegionalCPLBlindDate") 'lblRegCPLBlind
        dt.Columns.Add(dcRegCPLBlind)
        Dim dcGlobalRASDisc As New DataColumn("RegionalRasDiscDate") 'lblRegRASDics
        dt.Columns.Add(dcGlobalRASDisc)
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
        Dim dcRegCheckedRec As New DataColumn("RegCheckedRecFlag") 'chkRegCheckedRec
        dt.Columns.Add(dcRegCheckedRec)
        Dim dcRegRecSelected As New DataColumn("RegRecSelectedFlag") 'chkRegRecSelected
        dt.Columns.Add(dcRegRecSelected)

        If (dtProdBrands.Rows.Count > 0) Then
            For dr2 = 0 To (dtProdBrands.Rows.Count - 1)
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
                        dr("AvDetailID") = ""
                        dr("RCTOPlants_AVDetailProductBrand") = ""
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                        dr("ConfigRules") = dtProdBrands.Rows(dr2).Item("CategoryRules").ToString()
                        dr("Select") = False
                        dr("CheckedRecFlag") = False
                        dr("RecSelectedFlag") = False
                        dr("RegCheckedRecFlag") = False
                        dr("RegRecSelectedFlag") = False
                        dr("CatField") = True

                        strFeatCat = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        dt.Rows.Add(dr)
                    End If

                    If (dtProdBrands.Rows(dr2).Item("AvDetailID").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()) Then
                        dr = dt.NewRow()
                        dr("AvDetailID") = dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()
                        dr("RCTOPlants_AVDetailProductBrand") = dtProdBrands.Rows(dr2).Item("RCTOPlants_AVDetailProductBrand").ToString()
                        strTestText = dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        If IsDBNull(dtProdBrands.Rows(dr2).Item("RecSelectedFlag")) Then
                            dr("Select") = False
                        Else
                            dr("Select") = dtProdBrands.Rows(dr2).Item("RecSelectedFlag")
                        End If

                        dr("CatField") = False

                        Dim strRegStartDate As String
                        If (dtProdBrands.Rows(dr2).Item("PlantStartDate").ToString() = "") Or IsDBNull(dtProdBrands.Rows(dr2).Item("PlantStartDate").ToString()) Then
                            dr("PlantStartDate") = dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString()
                        Else
                            If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                                strRegStartDate = CDate(dtProdBrands.Rows(dr2).Item("PlantStartDate").ToString())
                                dr("PlantStartDate") = strRegStartDate
                            Else
                                dr("PlantStartDate") = dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString()
                            End If
                        End If

                        Dim strRegEndDate As String
                        If (dtProdBrands.Rows(dr2).Item("PlantEndDate").ToString() = "") Or IsDBNull(dtProdBrands.Rows(dr2).Item("PlantEndDate").ToString()) Then
                            dr("PlantEndDate") = dtProdBrands.Rows(dr2).Item("RegionalRASDiscDate").ToString()
                        Else
                            If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                                strRegEndDate = CDate(dtProdBrands.Rows(dr2).Item("PlantEndDate").ToString())
                                dr("PlantEndDate") = strRegEndDate
                            Else
                                dr("PlantEndDate") = dtProdBrands.Rows(dr2).Item("RegionalRASDiscDate").ToString()
                            End If
                        End If

                        dr("GPGDescription") = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()
                        dr("PlantName") = dtProdBrands.Rows(dr2).Item("PlantName").ToString()
                        dr("AVNo") = dtProdBrands.Rows(dr2).Item("AVNo").ToString()
                        dr("RegionalCPLBlindDate") = dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString()
                        dr("RegionalRasDiscDate") = dtProdBrands.Rows(dr2).Item("RegionalRASDiscDate").ToString()
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
                        strTestText = dtProdBrands.Rows(dr2).Item("CheckedRecFlag")
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("CheckedRecFlag")) Then
                            dr("CheckedRecFlag") = "N"
                        Else
                            If (dtProdBrands.Rows(dr2).Item("CheckedRecFlag") = "1") Then
                                dr("CheckedRecFlag") = "Y"
                            Else
                                dr("CheckedRecFlag") = "N"
                            End If
                        End If

                        strTestText = dtProdBrands.Rows(dr2).Item("RecSelectedFlag")
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("RecSelectedFlag")) Then
                            dr("RecSelectedFlag") = "N"
                        Else
                            If (dtProdBrands.Rows(dr2).Item("RecSelectedFlag") = "1") Then
                                dr("RecSelectedFlag") = "Y"
                            Else
                                dr("RecSelectedFlag") = "N"
                            End If
                        End If

                        strTestText = dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag")
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag")) Then
                            dr("RegCheckedRecFlag") = "N"
                        Else
                            If (dtProdBrands.Rows(dr2).Item("RegCheckedRecFlag") = "1") Then
                                dr("RegCheckedRecFlag") = "Y"
                            Else
                                dr("RegCheckedRecFlag") = "N"
                            End If
                        End If

                        strTestText = dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag")
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag")) Then
                            dr("RegRecSelectedFlag") = "N"
                        Else
                            If (dtProdBrands.Rows(dr2).Item("RegRecSelectedFlag") = "1") Then
                                dr("RegRecSelectedFlag") = "Y"
                            Else
                                dr("RegRecSelectedFlag") = "N"
                            End If
                        End If

                        dt.Rows.Add(dr)
                    End If

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
        Dim j As Integer
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
            Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim RCTOPlants_AVDetailProductBrand As Label = Row.FindControl("lblRCTOPlants_AVDetailProductBrand")
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
            Dim PlantCPLBlindLabel As Label = Row.FindControl("lblPlantCPLBlind") '-2
            Dim PlantCPLBlindText As TextBox = Row.FindControl("txtPlantCPLBlind") '-2
            Dim PlantRASDiscLabel As Label = Row.FindControl("lblPlantRASDisc") '-3
            Dim PlantRASDiscText As TextBox = Row.FindControl("txtPlantRASDisc") '-3
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
            Dim p_ProductBrandID As Label = Row.FindControl("lblProductBrandID") '-14
            Dim FeatureCategoryID As Label = Row.FindControl("lblAvFeatureCatID") '-15
            Dim GeoID As Label = Row.FindControl("lblGeoID") '16
            Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
            Dim AVRegionalDetailID As Label = Row.FindControl("lblAVRegionalDetailID") '18
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
            Dim RegCheckedRec As Label = Row.FindControl("chkRegCheckedRec") '21
            Dim RegRecSelected As Label = Row.FindControl("chkRegRecSelected") '22
            'FeatureCat.Text = FeatureCat.Text + "-Test"

            If (FeatureCat.Text <> strCurrFeatCat) Then
                'Row.Controls.Add(
                strCurrFeatCat = FeatureCat.Text
                ChangeRowColor(iCurrRowColor, i, True)
                chkSel.Visible = True
                CheckedRec.Visible = False
                RecSelected.Visible = False
                RegCheckedRec.Visible = False
                RegRecSelected.Visible = False

                PlantCPLBlindLabel.Visible = False
                PlantCPLBlindText.Visible = False
                PlantRASDiscLabel.Visible = False
                PlantRASDiscText.Visible = False
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
                    PlantCPLBlindLabel.Visible = True
                    PlantCPLBlindText.Visible = False
                    PlantRASDiscLabel.Visible = True
                    PlantRASDiscText.Visible = False
                ElseIf (chkSel.Checked = True) Then
                    PlantCPLBlindLabel.Visible = False
                    PlantCPLBlindText.Visible = True
                    PlantRASDiscLabel.Visible = False
                    PlantRASDiscText.Visible = True
                End If

                If (RegRecSelected.Text = "N") Then
                    Row.CssClass = "O"
                    chkSel.Visible = False
                    PlantCPLBlindLabel.Visible = True
                    PlantCPLBlindText.Visible = False
                    PlantRASDiscLabel.Visible = True
                    PlantRASDiscText.Visible = False
                End If
            End If
        Next

    End Sub

    Protected Sub ChangeRowColor(ByVal intColor As Integer, ByVal i As Integer, _
                                 ByVal bolAllColors As Boolean)

        Dim j As Integer
        Dim k As Integer

        If (bolAllColors = False) Then
            j = 5
        Else
            j = 0
        End If

        For k = j To 19
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

        'Dim chk As CheckBox = sender
        'Dim Row As GridViewRow = chk.NamingContainer

        'Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
        'Dim PlantCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
        'Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
        'Dim PlantRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
        'Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3

        'If (chkSel.Checked = False) Then
        '    PlantCPLBlindLabel.Visible = True
        '    RegCPLBlindText.Visible = False
        '    PlantRASDiscLabel.Visible = True
        '    RegRASDiscText.Visible = False
        'Else
        '    PlantCPLBlindLabel.Visible = False
        '    RegCPLBlindText.Visible = True
        '    PlantRASDiscLabel.Visible = False
        '    RegRASDiscText.Visible = True
        'End If

        Dim i As Integer
        Dim j As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim strRegion As String = ""

        'For i = 0 To (intRowCount - 1)
        '    Dim Row As GridViewRow
        '    Row = gvRegAVSelToolGrid.Rows(i)

        'Dim k As Integer
        Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
        Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
        Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
        Dim RCTOPlants_AVDetailProductBrand As Label = Row.FindControl("lblRCTOPlants_AVDetailProductBrand")
        Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
        Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
        Dim PlantCPLBlindLabel As Label = Row.FindControl("lblPlantCPLBlind") '-2
        Dim PlantCPLBlindText As TextBox = Row.FindControl("txtPlantCPLBlind") '-2
        Dim PlantRASDiscLabel As Label = Row.FindControl("lblPlantRASDisc") '-3
        Dim PlantRASDiscText As TextBox = Row.FindControl("txtPlantRASDisc") '-3
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
        Dim p_ProductBrandID As Label = Row.FindControl("lblProductBrandID") '-14
        Dim FeatureCategoryID As Label = Row.FindControl("lblAvFeatureCatID") '-15
        Dim GeoID As Label = Row.FindControl("lblGeoID") '16
        Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
        Dim AVRegionalDetailID As Label = Row.FindControl("lblAVRegionalDetailID") '18
        Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
        Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
        Dim RegCheckedRec As Label = Row.FindControl("chkCheckedRec") '21
        Dim RegRecSelected As Label = Row.FindControl("chkRecSelected") '22

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
                If (strCurrFeatCat <> FeatureCat2.Text) Then Exit Sub
                Dim RCTOPlantsID2 As Label = Row2.FindControl("lblRCTOPlantsID")
                Dim AVDetailID2 As Label = Row2.FindControl("lblAVDetailID") '-0
                Dim RCTOPlants_AVDetailProductBrand2 As Label = Row2.FindControl("lblRCTOPlants_AVDetailProductBrand")
                Dim chkSel2 As CheckBox = Row2.FindControl("chkSelect") '-1
                Dim PlantCPLBlindLabel2 As Label = Row2.FindControl("lblPlantCPLBlind") '-2
                Dim PlantCPLBlindText2 As TextBox = Row2.FindControl("txtPlantCPLBlind") '-2
                Dim PlantRASDiscLabel2 As Label = Row2.FindControl("lblPlantRASDisc") '-3
                Dim PlantRASDiscText2 As TextBox = Row2.FindControl("txtPlantRASDisc") '-3
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
                Dim GlobalSeriesConfigEOL2 As Label = Row2.FindControl("lblGlobalSeriesConfigEOL") '18
                Dim AVRegionalDetailID2 As Label = Row2.FindControl("lblAVRegionalDetailID") '18.5
                Dim CheckedRec2 As Label = Row2.FindControl("chkCheckedRec") '19
                Dim RecSelected2 As Label = Row2.FindControl("chkRecSelected") '20


                strTest = chk.Parent.Parent.ToString()

                If (strCurrFeatCat = FeatureCat2.Text) Then
                    If (chkSel.Checked = False) Then
                        If (chkSel2.Visible = True) Then
                            chkSel2.Checked = False

                            If (CheckedRec2.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                                Row2.CssClass = "O"

                                PlantCPLBlindLabel2.Font.Strikeout = True
                                PlantRASDiscLabel2.Font.Strikeout = True

                                'UpdateAvRegionalDates
                                HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID2.Text, AVRegionalDetailID2.Text, strRegion, _
                                                                             PlantCPLBlindText2.Text, PlantRASDiscText2.Text, _
                                                                             RCTOPlants_AVDetailProductBrand2.Text, "1", "0")

                                RecSelected2.Text = "N"
                            Else
                                PlantCPLBlindLabel2.Font.Strikeout = False
                                PlantRASDiscLabel2.Font.Strikeout = False
                            End If

                            PlantCPLBlindLabel2.Visible = True
                            PlantCPLBlindText2.Visible = False
                            PlantRASDiscLabel2.Visible = True
                            PlantRASDiscText2.Visible = False

                            RegCPLBlindLabel2.BackColor = GPGDesc.BackColor
                            RegRASDiscLabel2.BackColor = GPGDesc.BackColor
                        End If
                    Else
                        If (chkSel2.Visible = True) Then
                            chkSel2.Checked = True

                            If (CheckedRec2.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                                HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID2.Text, AVRegionalDetailID2.Text, strRegion, _
                                                                             PlantCPLBlindText2.Text, PlantRASDiscText2.Text, _
                                                                             RCTOPlants_AVDetailProductBrand2.Text, "1", "1")

                                RecSelected2.Text = "Y"
                                Row2.CssClass = ""
                            Else
                                HPQ.Excalibur.SupplyChain.InsertAvPlantDates_OneRecAtATtime(RCTOPlantsID2.Text, AVRegionalDetailID2.Text, _
                                                                                            strRegion, PlantCPLBlindText2.Text, PlantRASDiscText2.Text)

                                CheckedRec2.Text = "Y"
                                RecSelected2.Text = "Y"
                            End If

                            PlantCPLBlindLabel2.Visible = False
                            PlantCPLBlindText2.Visible = True
                            PlantRASDiscLabel2.Visible = False
                            PlantRASDiscText2.Visible = True
                            PlantCPLBlindText2.BackColor = Drawing.Color.White
                            PlantRASDiscText2.BackColor = Drawing.Color.White
                        End If
                    End If
                End If

                'chkSel2.Visible = True
            Next l
        Else
            If (chkSel.Checked = False) Then
                If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                    Row.CssClass = "O"

                    PlantCPLBlindLabel.Font.Strikeout = True
                    PlantRASDiscLabel.Font.Strikeout = True

                    'UpdateAvRegionalDates
                    HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                 PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                 RCTOPlants_AVDetailProductBrand.Text, "1", "0")

                    RecSelected.Text = "N"
                Else
                    PlantCPLBlindLabel.Font.Strikeout = False
                    PlantRASDiscLabel.Font.Strikeout = False
                End If

                PlantCPLBlindLabel.Visible = True
                PlantCPLBlindText.Visible = False
                PlantRASDiscLabel.Visible = True
                PlantRASDiscText.Visible = False

                RegCPLBlind.BackColor = GPGDesc.BackColor
                RegRASDisc.BackColor = GPGDesc.BackColor
            Else
                If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                    HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                 PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                 RCTOPlants_AVDetailProductBrand.Text, "1", "1")

                    RecSelected.Text = "Y"
                    Row.CssClass = ""
                Else
                    HPQ.Excalibur.SupplyChain.InsertAvPlantDates_OneRecAtATtime(RCTOPlantsID.Text, AVRegionalDetailID.Text, _
                                                                                strRegion, RegCPLBlind.Text, RegRASDisc.Text)

                    CheckedRec.Text = "Y"
                    RecSelected.Text = "Y"
                End If

                PlantCPLBlindLabel.Visible = False
                PlantCPLBlindText.Visible = True
                PlantRASDiscLabel.Visible = False
                PlantRASDiscText.Visible = True
                PlantCPLBlindText.BackColor = Drawing.Color.White
                PlantRASDiscText.BackColor = Drawing.Color.White
            End If
        End If
        'Next i
    End Sub

    Protected Sub txtPlantCPLBlind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim tb As TextBox = sender
        Dim row As GridViewRow = tb.NamingContainer

        Dim PlantCPLBlindLabel As Label = row.FindControl("lblPlantCPLBlind") '-2
        Dim PlantCPLBlindText As TextBox = row.FindControl("txtPlantCPLBlind") '-2
        Dim PlantRASDiscLabel As Label = row.FindControl("lblPlantRASDisc") '-3
        Dim PlantRASDiscText As TextBox = row.FindControl("txtPlantRASDisc") '-3
        Dim GPGDesc As Label = row.FindControl("lblGPGDesc") '-4
        Dim RegCPLBlind As Label = row.FindControl("lblRegCPLBlind") '-7
        Dim RegRASDisc As Label = row.FindControl("lblRegRASDics") '-8
        Dim RCTOPlantsID As Label = row.FindControl("lblRCTOPlantsID")
        Dim AVRegionalDetailID As Label = row.FindControl("lblAVRegionalDetailID") '18
        Dim RCTOPlants_AVDetailProductBrand As Label = row.FindControl("lblRCTOPlants_AVDetailProductBrand")

        Dim datGlobalCPLBlind As Date
        Dim datGlobalRASDate As Date
        Dim strRegion As String = ""
        Select Case sSelRegion
            Case "Americas"
                strRegion = "1"
            Case "EMEA"
                strRegion = "2"
            Case "APJ"
                strRegion = "3"
        End Select

        If (RegCPLBlind.Text <> "") Or IsDBNull(RegCPLBlind.Text) Then
            datGlobalCPLBlind = CDate(RegCPLBlind.Text)
        End If
        If (RegRASDisc.Text <> "") Or IsDBNull(RegCPLBlind.Text) Then
            datGlobalRASDate = CDate(RegRASDisc.Text)
        End If

        Dim datRegCPLBlind As Date

        If (PlantCPLBlindText.Text <> "") Then
            Try
                datRegCPLBlind = CDate(PlantCPLBlindText.Text)
            Catch ex As Exception
                PlantCPLBlindText.BackColor = Drawing.Color.Yellow
                PlantCPLBlindText.Text = ""
                RegCPLBlind.BackColor = Drawing.Color.Yellow
                RegRASDisc.BackColor = Drawing.Color.Yellow

                lblErrorMessage.Text = "The date you entered is not a valid date!"

                'Response.Write("<div style='position: absolute; top: 10px; left: 100px'>The Regional CPL Blind Date you entered does not fit within the range of the Global dates!</div>")
                lblErrorMessage.Visible = True

                Exit Sub
            End Try
        End If

        If (datRegCPLBlind < datGlobalCPLBlind) Or (datRegCPLBlind > datGlobalRASDate) Then
            PlantCPLBlindText.BackColor = Drawing.Color.Yellow
            PlantCPLBlindText.Text = ""
            PlantCPLBlindText.BackColor = Drawing.Color.Yellow
            RegCPLBlind.BackColor = Drawing.Color.Yellow

            lblErrorMessage.Text = "The date you entered does not fit within the range of the Plant dates!"
            lblErrorMessage.Visible = True
        Else
            PlantCPLBlindText.BackColor = Drawing.Color.White
            PlantCPLBlindLabel.Text = PlantCPLBlindText.Text
            RegCPLBlind.BackColor = GPGDesc.BackColor
            RegRASDisc.BackColor = GPGDesc.BackColor

            lblErrorMessage.Visible = False
        End If

        HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                     PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                     RCTOPlants_AVDetailProductBrand.Text, "1", "1")

    End Sub

    Protected Sub txtPlantRASDics_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim tb As TextBox = sender
        Dim row As GridViewRow = tb.NamingContainer

        Dim PlantCPLBlindLabel As Label = row.FindControl("lblPlantCPLBlind") '-2
        Dim PlantCPLBlindText As TextBox = row.FindControl("txtPlantCPLBlind") '-2
        Dim PlantRASDiscLabel As Label = row.FindControl("lblPlantRASDisc") '-3
        Dim PlantRASDiscText As TextBox = row.FindControl("txtPlantRASDisc") '-3
        Dim GPGDesc As Label = row.FindControl("lblGPGDesc") '-4
        Dim RegCPLBlind As Label = row.FindControl("lblRegCPLBlind") '-7
        Dim RegRASDisc As Label = row.FindControl("lblRegRASDics") '-8
        Dim RCTOPlantsID As Label = row.FindControl("lblRCTOPlantsID")
        Dim AVRegionalDetailID As Label = row.FindControl("lblAVRegionalDetailID") '18
        Dim RCTOPlants_AVDetailProductBrand As Label = row.FindControl("lblRCTOPlants_AVDetailProductBrand")

        Dim datGlobalCPLBlind As Date
        Dim datGlobalRASDate As Date
        Dim strRegion As String = ""
        Select Case sSelRegion
            Case "Americas"
                strRegion = "1"
            Case "EMEA"
                strRegion = "2"
            Case "APJ"
                strRegion = "3"
        End Select

        If (RegCPLBlind.Text <> "") Or IsDBNull(RegCPLBlind.Text) Then
            datGlobalCPLBlind = CDate(RegCPLBlind.Text)
        End If
        If (RegRASDisc.Text <> "") Or IsDBNull(RegCPLBlind.Text) Then
            datGlobalRASDate = CDate(RegRASDisc.Text)
        End If

        Dim datRegRASDate As Date

        If (PlantRASDiscText.Text <> "") Then
            Try
                datRegRASDate = CDate(PlantRASDiscText.Text)
            Catch ex As Exception
                PlantRASDiscText.BackColor = Drawing.Color.Yellow
                PlantRASDiscText.Text = ""
                RegCPLBlind.BackColor = Drawing.Color.Yellow
                RegRASDisc.BackColor = Drawing.Color.Yellow

                lblErrorMessage.Text = "The date you entered is not a valid date!"
                lblErrorMessage.Visible = True

                Exit Sub
            End Try
        End If

        If (datRegRASDate < datGlobalCPLBlind) Or (datRegRASDate > datGlobalRASDate) Then
            PlantRASDiscText.BackColor = Drawing.Color.Yellow
            PlantRASDiscText.Text = ""
            RegRASDisc.BackColor = Drawing.Color.Yellow

            'Response.Write("<div style='position: absolute; top: 10px; left: 100px'>The Regional CPL Blind Date you entered does not fit within the range of the Global dates!</div>")
            lblErrorMessage.Text = "The date you entered does not fit within the range of the Plant dates!"
            lblErrorMessage.Visible = True
        Else
            PlantRASDiscText.BackColor = Drawing.Color.White
            PlantRASDiscLabel.Text = PlantRASDiscText.Text
            RegCPLBlind.BackColor = GPGDesc.BackColor
            RegRASDisc.BackColor = GPGDesc.BackColor

            lblErrorMessage.Visible = False
        End If

        HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                     PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                     RCTOPlants_AVDetailProductBrand.Text, "1", "1")

    End Sub

    Protected Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged

        Dim i As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim strRegion As String = ""

        Select Case sSelRegion
            Case "Americas"
                strRegion = "1"
            Case "EMEA"
                strRegion = "2"
            Case "APJ"
                strRegion = "3"
        End Select

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            'Dim strTest As String = "" 'sender.ToString()
            'strTest = chk.Parent.Parent.ToString()

            'Dim k As Integer
            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim RCTOPlants_AVDetailProductBrand As Label = Row.FindControl("lblRCTOPlants_AVDetailProductBrand")
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim PlantCPLBlindLabel As Label = Row.FindControl("lblPlantCPLBlind") '-2
            Dim PlantCPLBlindText As TextBox = Row.FindControl("txtPlantCPLBlind") '-2
            Dim PlantRASDiscLabel As Label = Row.FindControl("lblPlantRASDisc") '-3
            Dim PlantRASDiscText As TextBox = Row.FindControl("txtPlantRASDisc") '-3
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
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
            Dim RegCheckedRec As Label = Row.FindControl("chkCheckedRec") '21
            Dim RegRecSelected As Label = Row.FindControl("chkRecSelected") '22

            If (strCurrFeatCat <> FeatureCat.Text) Then
                strCurrFeatCat = FeatureCat.Text
            Else
                If (chkSel.Visible = True) Then
                    If (chkAll.Checked = False) Then
                        chkSel.Checked = False

                        'PlantCPLBlindLabel.Visible = True
                        'PlantCPLBlindText.Visible = False
                        'PlantRASDiscLabel.Visible = True
                        'PlantRASDiscText.Visible = False

                        If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                            Row.CssClass = "O"

                            PlantCPLBlindLabel.Font.Strikeout = True
                            PlantRASDiscLabel.Font.Strikeout = True

                            'UpdateAvRegionalDates
                            HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                         PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                         RCTOPlants_AVDetailProductBrand.Text, "1", "0")

                            RecSelected.Text = "N"
                        Else
                            PlantCPLBlindLabel.Font.Strikeout = False
                            PlantRASDiscLabel.Font.Strikeout = False
                        End If

                        PlantCPLBlindLabel.Visible = True
                        PlantCPLBlindText.Visible = False
                        PlantRASDiscLabel.Visible = True
                        PlantRASDiscText.Visible = False

                        RegCPLBlind.BackColor = GPGDesc.BackColor
                        RegRASDisc.BackColor = GPGDesc.BackColor
                    Else
                        chkSel.Checked = True

                        'PlantCPLBlindLabel.Visible = False
                        'PlantCPLBlindText.Visible = True
                        'PlantRASDiscLabel.Visible = False
                        'PlantRASDiscText.Visible = True

                        If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                            HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                         PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                         RCTOPlants_AVDetailProductBrand.Text, "1", "1")

                            RecSelected.Text = "Y"
                            Row.CssClass = ""
                        Else
                            HPQ.Excalibur.SupplyChain.InsertAvPlantDates_OneRecAtATtime(RCTOPlantsID.Text, AVRegionalDetailID.Text, _
                                                                                        strRegion, RegCPLBlind.Text, RegRASDisc.Text)

                            CheckedRec.Text = "Y"
                            RecSelected.Text = "Y"
                        End If

                        PlantCPLBlindLabel.Visible = False
                        PlantCPLBlindText.Visible = True
                        PlantRASDiscLabel.Visible = False
                        PlantRASDiscText.Visible = True
                        PlantCPLBlindText.BackColor = Drawing.Color.White
                        PlantRASDiscText.BackColor = Drawing.Color.White
                    End If
                End If
            End If
        Next i
    End Sub

    'Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveChanges.Click

    '    Dim i As Integer
    '    Dim j As Integer
    '    Dim bolFound As Boolean
    '    'Dim k As Integer
    '    Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
    '    Dim strCurrFeatCat As String = ""
    '    Dim strRCTOPlantsID As String = ""
    '    Dim strAvRegionalDatesID As String = ""
    '    Dim strSelRegion As String = ""
    '    Dim dt As New DataTable()
    '    Dim dr As DataRow
    '    Dim dcID As New DataColumn("RCTOPlantsID") '-0
    '    dt.Columns.Add(dcID)
    '    Dim dcAvRegionalDatesID As New DataColumn("AvRegionalDatesID") '-1
    '    dt.Columns.Add(dcAvRegionalDatesID)
    '    Dim dcPlantStartDate As New DataColumn("PlantStartDate") '-2
    '    dt.Columns.Add(dcPlantStartDate)
    '    Dim dcPlantEndDate As New DataColumn("PlantEndDate") '-3
    '    dt.Columns.Add(dcPlantEndDate)
    '    'Dim dcStatus As New DataColumn("Status") '-4
    '    'dt.Columns.Add(dcStatus)
    '    Dim dcGeoID As New DataColumn("GeoID") '-5
    '    dt.Columns.Add(dcGeoID)

    '    lblErrorMessage.Text = ""

    '    For i = 0 To (intRowCount - 1)
    '        Dim Row As GridViewRow
    '        Row = gvRegAVSelToolGrid.Rows(i)

    '        Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
    '        Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
    '        Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
    '        Dim PlantCPLBlindLabel As Label = Row.FindControl("lblPlantCPLBlind") '-2
    '        Dim PlantCPLBlindText As TextBox = Row.FindControl("txtPlantCPLBlind") '-2
    '        Dim PlantRASDiscLabel As Label = Row.FindControl("lblPlantRASDisc") '-3
    '        Dim PlantRASDiscText As TextBox = Row.FindControl("txtPlantRASDisc") '-3
    '        Dim p_ProductBrandID As Label = Row.FindControl("lblProductBrandID") '-14
    '        Dim FeatureCategoryID As Label = Row.FindControl("lblAvFeatureCatID") '-15
    '        Dim GeoID As Label = Row.FindControl("lblRCTOGEOID") '("lblGeoID") '16
    '        Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17
    '        Row.FindControl("lblAVRegionalDetailID").Visible = True
    '        Dim AvRegionalDatesID As Label = Row.FindControl("lblAVRegionalDetailID")
    '        Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")

    '        bolFound = False
    '        If (strRCTOPlantsID = "") Then
    '            strRCTOPlantsID = RCTOPlantsID.Text
    '        Else
    '            Dim strRCTOPlantsIDs() As String = strRCTOPlantsID.Split(",")
    '            For j = 0 To UBound(strRCTOPlantsIDs)
    '                If (strRCTOPlantsIDs(j) = RCTOPlantsID.Text) Then
    '                    bolFound = True
    '                End If
    '            Next j

    '            If (bolFound = False) Then
    '                strRCTOPlantsID = strRCTOPlantsID & "," & RCTOPlantsID.Text
    '            End If
    '        End If

    '        bolFound = False
    '        If (strAvRegionalDatesID = "") Then
    '            strAvRegionalDatesID = AvRegionalDatesID.Text
    '        Else
    '            Dim strAvRegionalDatesIDs() As String = strAvRegionalDatesID.Split(",")
    '            For j = 0 To UBound(strAvRegionalDatesIDs)
    '                If (strAvRegionalDatesIDs(j) = AvRegionalDatesID.Text) Then
    '                    bolFound = True
    '                End If
    '            Next j

    '            If (bolFound = False) Then
    '                strAvRegionalDatesID = strAvRegionalDatesID & "," & AvRegionalDatesID.Text
    '            End If
    '        End If

    '        If (strCurrFeatCat <> FeatureCat.Text) Then
    '            strCurrFeatCat = FeatureCat.Text
    '        ElseIf (PlantCPLBlindText.Text = "") Then
    '            lblErrorMessage.Text = "Date Missing! Enter a Plant Start Date and try again!"
    '            lblErrorMessage.Visible = True

    '            Exit Sub
    '        ElseIf (PlantRASDiscText.Text = "") Then
    '            lblErrorMessage.Text = "Date Missing! Enter a Plant End Date and try again!"
    '            lblErrorMessage.Visible = True

    '            Exit Sub
    '        Else
    '            If (chkSel.Checked = True) Then
    '                dr = dt.NewRow()

    '                dr("RCTOPlantsID") = RCTOPlantsID.Text
    '                dr("AvRegionalDatesID") = AvRegionalDatesID.Text

    '                If (chkSel.Checked = False) Then
    '                    dr("PlantStartDate") = PlantCPLBlindLabel.Text
    '                    dr("PlantEndDate") = PlantRASDiscLabel.Text
    '                Else
    '                    dr("PlantStartDate") = PlantCPLBlindText.Text
    '                    dr("PlantEndDate") = PlantRASDiscText.Text
    '                End If

    '                dr("GeoID") = GeoID.Text
    '                strSelRegion = GeoID.Text

    '                'dr("Status") = "2"

    '                dt.Rows.Add(dr)
    '            End If
    '        End If
    '    Next i

    '    HPQ.Excalibur.SupplyChain.InsertAvPlantDates(dt, strRCTOPlantsID, strAvRegionalDatesID, strSelRegion)
    '    lblErrorMessage.Text = "Data has been saved Successfully!"
    '    lblErrorMessage.Visible = True

    '    'ProcessFilter(sProdBrandIDs, sCatIDs, sSelRegion)

    'End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Dim i As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim strRegion As String = ""

        Select Case sSelRegion
            Case "Americas"
                strRegion = "1"
            Case "EMEA"
                strRegion = "2"
            Case "APJ"
                strRegion = "3"
        End Select

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            'Dim strTest As String = "" 'sender.ToString()
            'strTest = chk.Parent.Parent.ToString()

            'Dim k As Integer
            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim RCTOPlantsID As Label = Row.FindControl("lblRCTOPlantsID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim RCTOPlants_AVDetailProductBrand As Label = Row.FindControl("lblRCTOPlants_AVDetailProductBrand")
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim PlantCPLBlindLabel As Label = Row.FindControl("lblPlantCPLBlind") '-2
            Dim PlantCPLBlindText As TextBox = Row.FindControl("txtPlantCPLBlind") '-2
            Dim PlantRASDiscLabel As Label = Row.FindControl("lblPlantRASDisc") '-3
            Dim PlantRASDiscText As TextBox = Row.FindControl("txtPlantRASDisc") '-3
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
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
            Dim RegCheckedRec As Label = Row.FindControl("chkCheckedRec") '21
            Dim RegRecSelected As Label = Row.FindControl("chkRecSelected") '22

            If (strCurrFeatCat <> FeatureCat.Text) Then
                strCurrFeatCat = FeatureCat.Text
            Else
                If (chkSel.Visible = True) Then
                    If (chkSel.Checked = False) Then
                        If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                            Row.CssClass = "O"

                            PlantCPLBlindLabel.Font.Strikeout = True
                            PlantRASDiscLabel.Font.Strikeout = True

                            'UpdateAvRegionalDates
                            HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                         PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                         RCTOPlants_AVDetailProductBrand.Text, "1", "0")

                            RecSelected.Text = "N"
                        Else
                            PlantCPLBlindLabel.Font.Strikeout = False
                            PlantRASDiscLabel.Font.Strikeout = False
                        End If

                        PlantCPLBlindLabel.Visible = True
                        PlantCPLBlindText.Visible = False
                        PlantRASDiscLabel.Visible = True
                        PlantRASDiscText.Visible = False

                        RegCPLBlind.BackColor = GPGDesc.BackColor
                        RegRASDisc.BackColor = GPGDesc.BackColor
                    Else
                        If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                            HPQ.Excalibur.SupplyChain.UpdateAvPlantDates(RCTOPlantsID.Text, AVRegionalDetailID.Text, strRegion, _
                                                                         PlantCPLBlindText.Text, PlantRASDiscText.Text, _
                                                                         RCTOPlants_AVDetailProductBrand.Text, "1", "1")

                            RecSelected.Text = "Y"
                            Row.CssClass = ""
                        Else
                            HPQ.Excalibur.SupplyChain.InsertAvPlantDates_OneRecAtATtime(RCTOPlantsID.Text, AVRegionalDetailID.Text, _
                                                                                        strRegion, RegCPLBlind.Text, RegRASDisc.Text)

                            CheckedRec.Text = "Y"
                            RecSelected.Text = "Y"
                        End If

                        PlantCPLBlindLabel.Visible = False
                        PlantCPLBlindText.Visible = True
                        PlantRASDiscLabel.Visible = False
                        PlantRASDiscText.Visible = True
                        PlantCPLBlindText.BackColor = Drawing.Color.White
                        PlantRASDiscText.BackColor = Drawing.Color.White
                    End If
                End If
            End If
        Next i
    End Sub
End Class








