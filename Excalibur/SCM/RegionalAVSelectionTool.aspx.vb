Imports System.Data

Partial Class SCM_RegionalAVSelectorTool
    Inherits System.Web.UI.Page

    Public Shared Function GetSessionStateValue(ByVal id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Function AddSessionStateValue(ByVal id As String, ByVal obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Function

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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            MessageLabel.Visible = True
            lblErrorMessage.Visible = False
            lblRecCountMsg.Visible = True
            lblRecCountMsg1.Visible = True
            lblRecCountMsg2.Visible = True

            If Not Me.Page.IsPostBack Then
                Dim dtProdBrands As DataTable
                Dim sIDQSPara As String = "" 'This will be populated by the ID Query String Parameter.
                Dim intPos As Integer = "-1"
                Dim intTotalRecs As Integer = 0

                sIDQSPara = Request.QueryString("ID").ToString()
                'sIDQSPara = "763,761,769,787:1085,1051,1050,1048,1062::1"
                'sIDQSPara = "763:1050:-1,6,51:1"
                'sIDQSPara = "763:1050:51:1"

                'Get the Values passed to this page and deal with them appropriately.
                Dim strSplitUpIDQSPara() As String = sIDQSPara.Split(":")
                For i As Integer = 0 To UBound(strSplitUpIDQSPara)
                    Select Case i
                        Case 0
                            sProdVerIDs = strSplitUpIDQSPara(i)
                        Case 1
                            sProdBrandIDs = strSplitUpIDQSPara(i)
                        Case 2
                            sCatIDs = strSplitUpIDQSPara(i)
                        Case 3
                            sSelRegion = strSplitUpIDQSPara(i)
                    End Select
                Next i

                If (sCatIDs = "-1") Then
                    sCatIDs = ""
                End If

                Select Case sSelRegion
                    Case 1
                        lblRegion.Text = "Americas"
                    Case 2
                        lblRegion.Text = "EMEA"
                    Case 3
                        lblRegion.Text = "APJ"
                End Select

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

                If dtProdBrands.Rows.Count = 0 Then
                    lblRecCountMsg.Visible = False
                    lblRecCountMsg1.Visible = True
                    lblRecCountMsg1.Text = "No AV Items found for these brands."
                    lblRecCountMsg2.Visible = False
                Else
                    Dim dtProdBrandsCustom As DataTable
                    dtProdBrands = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView(sProdVerIDs, sProdBrandIDs, sCatIDs, sSelRegion)

                    Dim dtRegionsTableFromExcalibur As DataTable
                    dtRegionsTableFromExcalibur = HPQ.Excalibur.SupplyChain.Regions_Select_All()

                    dtProdBrandsCustom = SetupDataTableForGridView(dtProdBrands, sSelRegion, intTotalRecs, dtRegionsTableFromExcalibur)

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
                        lblRecCountMsg2.Text = lblRegion.Text & " Regional View"

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
            'sSelRegion,sCatIDs,sProdBrandIDs,sProdVerIDs
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sProdBrands & ":" & sCategory & ":" & strSelRegion & ":True"))
            'Response.Write("<script language='javascript'> { window.close();}</script>")

            'End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.InnerException.ToString
        End Try
    End Function

    Private Function SetupDataTableForGridView(ByVal dtProdBrands As DataTable, ByVal strRegionID As String, _
                                               ByRef intTotalRecs As Integer, ByVal dtRegionsTableFromExcalibur As DataTable) As DataTable

        Dim strTestText As String = ""
        Dim strFeatCat As String = ""
        Dim strAvDetail_ProdBrandID As String = ""
        Dim dr2 As Integer = 0
        Dim i As Integer = 0
        Dim intRowsToSkipTo As Integer = 0
        Dim bolAddRec As Boolean = False
        Dim bolUnSelectedRec As Boolean = False

        'Dim GlobalSeriesConfigEOL As Label = Row.FindControl("lblGlobalSeriesConfigEOL") '18

        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim dcAVDetailID As New DataColumn("AvDetailID") 'lblAVDetailID
        dt.Columns.Add(dcAVDetailID)
        Dim dcFeatCat As New DataColumn("AvFeatureCategory") 'lblFeatCat
        dt.Columns.Add(dcFeatCat)
        Dim dcAVRegionalDatesID As New DataColumn("AvRegionalDatesID") 'lblAVRegionalDatesID
        dt.Columns.Add(dcAVRegionalDatesID)
        Dim dcSel As New DataColumn("Select") 'chkSelect
        dcSel.DataType = GetType(Boolean)
        dt.Columns.Add(dcSel)
        Dim dcCatField As New DataColumn("CatField") 'chkCatField
        'dcCatField.DataType = GetType(Boolean)
        dt.Columns.Add(dcCatField)
        Dim dcRegCPLBlind As New DataColumn("RegCPLBlind") 'lblRegCPLBlind
        dt.Columns.Add(dcRegCPLBlind)
        Dim dcRegRASDics As New DataColumn("RegRASDics") 'lblRegRASDics
        dt.Columns.Add(dcRegRASDics)
        Dim dcGPGDesc As New DataColumn("GPGDescription") 'lblGPGDesc
        dt.Columns.Add(dcGPGDesc)
        Dim dcProdName As New DataColumn("ShortProdName") 'lblProdName
        dt.Columns.Add(dcProdName)
        Dim dcAVNo As New DataColumn("AVNo") 'lblAVNo
        dt.Columns.Add(dcAVNo)
        Dim dcGlobalCPLBlind As New DataColumn("CBLBlindDate") 'lblGlobalCPLBlind
        dt.Columns.Add(dcGlobalCPLBlind)
        Dim dcGlobalRASDisc As New DataColumn("RASDiscontinueDate") 'lblGlobalRASDisc
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
        Dim dcCheckedRec As New DataColumn("CheckedRecFlag") 'chkSelect
        dt.Columns.Add(dcCheckedRec)
        Dim dcRecSelected As New DataColumn("RecSelectedFlag") 'chkRecSelected
        dt.Columns.Add(dcRecSelected)

        If (dtProdBrands.Rows.Count > 0) Then
            For dr2 = 0 To (dtProdBrands.Rows.Count - 1)
                'Row = gvRegAVSelToolGrid.Rows(i)
                'For Each dr2 As DataRow In dtProdBrands.Rows
                'Check if the region in this current record is the same as the region the user is filtering by.
                If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
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
                            If (dtProdBrands.Rows(i + 1).Item("p_ProductBrandID").ToString() = dtProdBrands.Rows(i).Item("p_ProductBrandID").ToString()) Then
                                intRowsToSkipTo = intRowsToSkipTo + 1
                            Else 'Else, it is the not the same "p_ProductBrandID" and it needs to be processed next.
                                Exit Do
                            End If

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
                    'Now loop thru the regions datatable to find a matching OptionConfig and GEOID
                    Dim dtRTFE As Integer = 0
                    Dim bolLocalizedAVNo As Boolean = False
                    Dim intPos As Integer = 0
                    Dim strTestAVNoValue As String = dtProdBrands.Rows(dr2).Item("AVNo").ToString()
                    Dim strTestOptionConfigValue As String = dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("OptionConfig").ToString()
                    Dim strTestOptionConfigGeoIDValue As String = dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("GeoID").ToString()

                    If (dtProdBrands.Rows(dr2).Item("AVNo").ToString().Contains("#") = True) Then
                        intPos = dtProdBrands.Rows(dr2).Item("AVNo").ToString().IndexOf("#")
                        If (dtRegionsTableFromExcalibur.Rows.Count > 0) Then
                            For dtRTFE = 0 To (dtRegionsTableFromExcalibur.Rows.Count - 1)
                                strTestOptionConfigValue = dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("OptionConfig").ToString()
                                strTestOptionConfigGeoIDValue = dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("GeoID").ToString()
                                If (dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("OptionConfig").ToString() = dtProdBrands.Rows(dr2).Item("AVNo").ToString().Remove(0, (intPos + 1))) Then
                                    If (strRegionID = dtRegionsTableFromExcalibur.Rows(dtRTFE).Item("GeoID").ToString()) Then
                                        bolLocalizedAVNo = True

                                        Exit For
                                    End If
                                End If
                            Next dtRTFE

                            If (bolLocalizedAVNo = False) Then
                                bolAddRec = False
                            End If
                        End If
                    End If
                End If

                If (bolAddRec = True) Then
                    intTotalRecs = intTotalRecs + 1
                    dr = dt.NewRow()

                    If (strFeatCat <> dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()) Then
                        dr("AvDetailID") = ""
                        dr("AvRegionalDatesID") = ""
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                        dr("ConfigRules") = dtProdBrands.Rows(dr2).Item("CategoryRules").ToString()
                        dr("Select") = False
                        dr("CheckedRecFlag") = False
                        dr("RecSelectedFlag") = False
                        dr("CatField") = True

                        strFeatCat = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        dt.Rows.Add(dr)
                    End If

                    If (dtProdBrands.Rows(dr2).Item("AvDetailID").ToString() <> "") And Not IsDBNull(dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()) Then
                        dr = dt.NewRow()
                        dr("AvDetailID") = dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()
                        dr("AvRegionalDatesID") = dtProdBrands.Rows(dr2).Item("AvRegionalDatesID").ToString()
                        strTestText = dtProdBrands.Rows(dr2).Item("AvDetailID").ToString()
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                            If IsDBNull(dtProdBrands.Rows(dr2).Item("RecSelectedFlag")) Then
                                dr("Select") = False
                            Else
                                dr("Select") = dtProdBrands.Rows(dr2).Item("RecSelectedFlag")
                            End If
                        Else
                            dr("Select") = False
                        End If

                        dr("CatField") = False

                        Dim strRegStartDate As String
                        If (dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString() = "") Or IsDBNull(dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString()) Then
                            dr("RegCPLBlind") = dtProdBrands.Rows(dr2).Item("CBLBlindDate").ToString()
                        Else
                            If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                                strRegStartDate = CDate(dtProdBrands.Rows(dr2).Item("RegionalCPLBlindDate").ToString())
                                dr("RegCPLBlind") = strRegStartDate
                            Else
                                dr("RegCPLBlind") = dtProdBrands.Rows(dr2).Item("CBLBlindDate").ToString()
                            End If
                        End If

                        Dim strRegEndDate As String
                        If (dtProdBrands.Rows(dr2).Item("RegionalRasDiscDate").ToString() = "") Or IsDBNull(dtProdBrands.Rows(dr2).Item("RegionalRasDiscDate").ToString()) Then
                            dr("RegRASDics") = dtProdBrands.Rows(dr2).Item("RASDiscontinueDate").ToString()
                        Else
                            If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
                                strRegEndDate = CDate(dtProdBrands.Rows(dr2).Item("RegionalRasDiscDate").ToString())
                                dr("RegRASDics") = strRegEndDate
                            Else
                                dr("RegRASDics") = dtProdBrands.Rows(dr2).Item("RASDiscontinueDate").ToString()
                            End If
                        End If

                        dr("GPGDescription") = dtProdBrands.Rows(dr2).Item("GPGDescription").ToString()
                        dr("ShortProdName") = dtProdBrands.Rows(dr2).Item("ShortProdName").ToString()
                        dr("AVNo") = dtProdBrands.Rows(dr2).Item("AVNo").ToString()
                        dr("CBLBlindDate") = dtProdBrands.Rows(dr2).Item("CBLBlindDate").ToString()
                        dr("RASDiscontinueDate") = dtProdBrands.Rows(dr2).Item("RASDiscontinueDate").ToString()
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

                        If (dtProdBrands.Rows(dr2).Item("GeoID").ToString() = strRegionID) Then
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
                        Else
                            dr("CheckedRecFlag") = "N"
                            dr("RecSelectedFlag") = "N"
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
            Dim AVRegionalDatesID As Label = Row.FindControl("lblAVRegionalDatesID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
            Dim RegCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
            Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
            Dim RegRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
            Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3
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
            Dim GlobalSeriesConfigEOL As Label = Row.FindControl("lblGlobalSeriesConfigEOL") '18
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20
            'FeatureCat.Text = FeatureCat.Text + "-Test"

            If (FeatureCat.Text <> strCurrFeatCat) Then
                strCurrFeatCat = FeatureCat.Text
                ChangeRowColor(iCurrRowColor, i, True)
                chkSel.Visible = True
                CheckedRec.Visible = False
                RecSelected.Visible = False

                RegCPLBlindLabel.Visible = False
                RegCPLBlindText.Visible = False
                RegRASDiscLabel.Visible = False
                RegRASDiscText.Visible = False
                chkCatField.Text = "True"
            Else
                FeatureCat.Visible = False
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

                If (chkSel.Checked = False) Then
                    If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                        Row.CssClass = "O"
                    End If
                    RegCPLBlindLabel.Visible = True
                    RegCPLBlindText.Visible = False
                    RegRASDiscLabel.Visible = True
                    RegRASDiscText.Visible = False
                ElseIf (chkSel.Checked = True) Then
                    RegCPLBlindLabel.Visible = False
                    RegCPLBlindText.Visible = True
                    RegRASDiscLabel.Visible = False
                    RegRASDiscText.Visible = True
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

        Dim i As Integer
        Dim l As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""

        'Dim k As Integer
        Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
        Dim AVRegionalDatesID As Label = Row.FindControl("lblAVRegionalDatesID")
        Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
        Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
        Dim chkCatField As Label = Row.FindControl("chkCatField") '-1.5
        Dim RegCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
        Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
        Dim RegRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
        Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3
        Dim GPGDesc As Label = Row.FindControl("lblGPGDesc") '-4
        Dim ProdName As Label = Row.FindControl("lblProdName") '-5
        Dim AVNo As Label = Row.FindControl("lblAVNo") '-6
        strTest = AVNo.Text
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
        Dim GlobalSeriesConfigEOL As Label = Row.FindControl("lblGlobalSeriesConfigEOL") '18
        Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
        Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20

        strCurrFeatCat = FeatureCat.Text

        'chkSel.Visible = False
        'CheckedRec.Visible = False
        'RegCPLBlindLabel.Visible = False
        'RegCPLBlindText.Visible = False
        'RegRASDiscLabel.Visible = False
        'RegRASDiscText.Visible = False

        If (chkCatField.Text = "True") Then
            bolFeatCat = True 'This is a feature category row.


            'Dim Row As GridViewRow = chk.NamingContainer

            Dim m As Integer = (Row.RowIndex + 1)
            Dim bolCatFeatChecked = chkSel.Checked

            For l = m To (intRowCount - 1)
                Dim Row2 As GridViewRow
                Row2 = gvRegAVSelToolGrid.Rows(l)

                Dim FeatureCat2 As Label = Row2.FindControl("lblFeatCat")
                Dim AVRegionalDatesID2 As Label = Row2.FindControl("lblAVRegionalDatesID")
                Dim AVDetailID2 As Label = Row2.FindControl("lblAVDetailID") '-0
                Dim chkSel2 As CheckBox = Row2.FindControl("chkSelect") '-1
                Dim RegCPLBlindLabel2 As Label = Row2.FindControl("lblRegCPLBlind") '-2
                Dim RegCPLBlindText2 As TextBox = Row2.FindControl("txtRegCPLBlind") '-2
                Dim RegRASDiscLabel2 As Label = Row2.FindControl("lblRegRASDics") '-3
                Dim RegRASDiscText2 As TextBox = Row2.FindControl("txtRegRASDics") '-3
                Dim GPGDesc2 As Label = Row2.FindControl("lblGPGDesc") '-4
                Dim ProdName2 As Label = Row2.FindControl("lblProdName") '-5
                Dim AVNo2 As Label = Row2.FindControl("lblAVNo") '-6
                strTest = AVNo.Text
                Dim GlobalCPLBlind2 As Label = Row2.FindControl("lblGlobalCPLBlind") '-7
                Dim GlobalRASDisc2 As Label = Row2.FindControl("lblGlobalRASDisc") '-8
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
                Dim CheckedRec2 As Label = Row2.FindControl("chkCheckedRec") '19
                Dim RecSelected2 As Label = Row2.FindControl("chkRecSelected") '20


                strTest = chk.Parent.Parent.ToString()

                If (strCurrFeatCat = FeatureCat2.Text) Then
                    If (chkSel.Checked = False) Then
                        chkSel2.Checked = False

                        If (CheckedRec2.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                            Row2.CssClass = "O"

                            RegCPLBlindLabel2.Font.Strikeout = True
                            RegRASDiscLabel2.Font.Strikeout = True

                            'UpdateAvRegionalDates
                            HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID2.Text, sSelRegion, _
                                                                            AvDetail_ProdBrandID2.Text, RegCPLBlindText2.Text, _
                                                                            RegRASDiscText2.Text, "3", p_ProductBrandID2.Text, _
                                                                            FeatureCategoryID2.Text, "1", "0")

                            RecSelected2.Text = "N"
                        Else
                            RegCPLBlindLabel2.Font.Strikeout = False
                            RegRASDiscLabel2.Font.Strikeout = False
                        End If

                        RegCPLBlindLabel2.Visible = True
                        RegCPLBlindLabel2.CssClass = "0"
                        RegCPLBlindText2.Visible = False
                        RegRASDiscLabel2.Visible = True
                        RegRASDiscLabel2.CssClass = "0"
                        RegRASDiscText2.Visible = False

                        RegCPLBlindLabel2.Text = RegCPLBlindText2.Text
                        RegRASDiscLabel2.Text = RegRASDiscText2.Text

                        GlobalCPLBlind2.BackColor = GPGDesc2.BackColor
                        GlobalRASDisc2.BackColor = GPGDesc2.BackColor
                    Else
                        chkSel2.Checked = True

                        If (CheckedRec2.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                            HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID2.Text, sSelRegion, _
                                                                            AvDetail_ProdBrandID2.Text, RegCPLBlindText2.Text, _
                                                                            RegRASDiscText2.Text, "4", p_ProductBrandID2.Text, _
                                                                            FeatureCategoryID2.Text, "1", "1")

                            RecSelected2.Text = "Y"
                            Row2.CssClass = ""
                        Else
                            HPQ.Excalibur.SupplyChain.InsertAvRegionalDates(sSelRegion, AvDetail_ProdBrandID2.Text, GlobalCPLBlind2.Text, _
                                                                            GlobalRASDisc2.Text, "2", p_ProductBrandID2.Text, _
                                                                            FeatureCategoryID2.Text)

                            CheckedRec2.Text = "Y"
                            RecSelected2.Text = "Y"
                        End If

                        'InsertAvRegionalDates
                        RegCPLBlindLabel2.Visible = False
                        RegCPLBlindText2.Visible = True
                        RegRASDiscLabel2.Visible = False
                        RegRASDiscText2.Visible = True
                        RegCPLBlindText2.BackColor = Drawing.Color.White
                        RegRASDiscText2.BackColor = Drawing.Color.White
                    End If
                End If

                chkSel2.Visible = True
            Next l
        Else
            If (chkSel.Checked = False) Then
                If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                    Row.CssClass = "O"

                    RegCPLBlindLabel.Font.Strikeout = True
                    RegRASDiscLabel.Font.Strikeout = True

                    'UpdateAvRegionalDates
                    HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                                    AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                    RegRASDiscText.Text, "3", p_ProductBrandID.Text, _
                                                                    FeatureCategoryID.Text, "1", "0")

                    RecSelected.Text = "N"
                Else
                    RegCPLBlindLabel.Font.Strikeout = False
                    RegRASDiscLabel.Font.Strikeout = False
                End If

                RegCPLBlindLabel.Visible = True
                RegCPLBlindLabel.CssClass = "0"
                RegCPLBlindText.Visible = False
                RegRASDiscLabel.Visible = True
                RegRASDiscLabel.CssClass = "0"
                RegRASDiscText.Visible = False

                RegCPLBlindLabel.Text = RegCPLBlindText.Text
                RegRASDiscLabel.Text = RegRASDiscText.Text

                GlobalCPLBlind.BackColor = GPGDesc.BackColor
                GlobalRASDisc.BackColor = GPGDesc.BackColor
            Else
                If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                    HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                                    AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                    RegRASDiscText.Text, "4", p_ProductBrandID.Text, _
                                                                    FeatureCategoryID.Text, "1", "1")

                    RecSelected.Text = "Y"
                    Row.CssClass = ""
                Else
                    HPQ.Excalibur.SupplyChain.InsertAvRegionalDates(sSelRegion, AvDetail_ProdBrandID.Text, GlobalCPLBlind.Text, _
                                                                    GlobalRASDisc.Text, "2", p_ProductBrandID.Text, _
                                                                    FeatureCategoryID.Text)

                    CheckedRec.Text = "Y"
                    RecSelected.Text = "Y"
                End If

                'InsertAvRegionalDates
                RegCPLBlindLabel.Visible = False
                RegCPLBlindText.Visible = True
                RegRASDiscLabel.Visible = False
                RegRASDiscText.Visible = True
                RegCPLBlindText.BackColor = Drawing.Color.White
                RegRASDiscText.BackColor = Drawing.Color.White
            End If
        End If
        'Next i
    End Sub

    Protected Sub txtRegCPLBlind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim tb As TextBox = sender
        Dim row As GridViewRow = tb.NamingContainer

        Dim AVRegionalDatesID As Label = row.FindControl("lblAVRegionalDatesID")
        Dim RegCPLBlindLabel As Label = row.FindControl("lblRegCPLBlind") '-2
        Dim RegCPLBlindText As TextBox = row.FindControl("txtRegCPLBlind") '-2
        Dim RegRASDiscLabel As Label = row.FindControl("lblRegRASDics") '-3
        Dim RegRASDiscText As TextBox = row.FindControl("txtRegRASDics") '-3
        Dim GPGDesc As Label = row.FindControl("lblGPGDesc") '-4
        Dim GlobalCPLBlind As Label = row.FindControl("lblGlobalCPLBlind") '-7
        Dim GlobalRASDisc As Label = row.FindControl("lblGlobalRASDisc") '-8
        Dim p_ProductBrandID As Label = row.FindControl("lblProductBrandID") '-14
        Dim FeatureCategoryID As Label = row.FindControl("lblAvFeatureCatID") '-15
        Dim AvDetail_ProdBrandID As Label = row.FindControl("lblAvDetail_ProductBrandID") '17

        Dim datGlobalCPLBlind As Date
        Dim datGlobalRASDate As Date

        If (GlobalCPLBlind.Text <> "") Or IsDBNull(GlobalCPLBlind.Text) Then
            datGlobalCPLBlind = CDate(GlobalCPLBlind.Text)
        End If
        If (GlobalRASDisc.Text <> "") Or IsDBNull(GlobalCPLBlind.Text) Then
            datGlobalRASDate = CDate(GlobalRASDisc.Text)
        End If

        Dim datRegCPLBlind As Date

        If (RegCPLBlindText.Text <> "") Then
            Try
                datRegCPLBlind = CDate(RegCPLBlindText.Text)
            Catch ex As Exception
                RegCPLBlindText.BackColor = Drawing.Color.Yellow
                RegCPLBlindText.Text = ""
                GlobalCPLBlind.BackColor = Drawing.Color.Yellow
                GlobalRASDisc.BackColor = Drawing.Color.Yellow

                lblErrorMessage.Text = "The date you entered is not a valid date!"

                'Response.Write("<div style='position: absolute; top: 10px; left: 100px'>The Regional CPL Blind Date you entered does not fit within the range of the Global dates!</div>")
                lblErrorMessage.Visible = True

                Exit Sub
            End Try
        End If

        If (datRegCPLBlind < datGlobalCPLBlind) Or (datRegCPLBlind > datGlobalRASDate) Then
            RegCPLBlindText.BackColor = Drawing.Color.Yellow
            RegCPLBlindText.Text = ""
            GlobalCPLBlind.BackColor = Drawing.Color.Yellow
            GlobalRASDisc.BackColor = Drawing.Color.Yellow

            lblErrorMessage.Text = "The date you entered does not fit within the range of the Global dates!"
            lblErrorMessage.Visible = True
        Else
            RegCPLBlindText.BackColor = Drawing.Color.White
            RegCPLBlindLabel.Text = RegCPLBlindText.Text
            GlobalCPLBlind.BackColor = GPGDesc.BackColor
            GlobalRASDisc.BackColor = GPGDesc.BackColor

            lblErrorMessage.Visible = False
        End If

        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                        AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                        RegRASDiscText.Text, "4", p_ProductBrandID.Text, _
                                                        FeatureCategoryID.Text, "1", "1")
    End Sub

    Protected Sub txtRegRASDics_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim tb As TextBox = sender
        Dim row As GridViewRow = tb.NamingContainer

        Dim AVRegionalDatesID As Label = row.FindControl("lblAVRegionalDatesID")
        Dim RegCPLBlindLabel As Label = row.FindControl("lblRegCPLBlind") '-2
        Dim RegCPLBlindText As TextBox = row.FindControl("txtRegCPLBlind") '-2
        Dim RegRASDiscLabel As Label = row.FindControl("lblRegRASDics") '-3
        Dim RegRASDiscText As TextBox = row.FindControl("txtRegRASDics") '-3
        Dim GPGDesc As Label = row.FindControl("lblGPGDesc") '-4
        Dim GlobalCPLBlind As Label = row.FindControl("lblGlobalCPLBlind") '-7
        Dim GlobalRASDisc As Label = row.FindControl("lblGlobalRASDisc") '-8
        Dim p_ProductBrandID As Label = row.FindControl("lblProductBrandID") '-14
        Dim FeatureCategoryID As Label = row.FindControl("lblAvFeatureCatID") '-15
        Dim AvDetail_ProdBrandID As Label = row.FindControl("lblAvDetail_ProductBrandID") '17

        Dim datGlobalCPLBlind As Date
        Dim datGlobalRASDate As Date

        If (GlobalCPLBlind.Text <> "") Or IsDBNull(GlobalCPLBlind.Text) Then
            datGlobalCPLBlind = CDate(GlobalCPLBlind.Text)
        End If
        If (GlobalRASDisc.Text <> "") Or IsDBNull(GlobalCPLBlind.Text) Then
            datGlobalRASDate = CDate(GlobalRASDisc.Text)
        End If

        Dim datRegRASDate As Date

        If (RegRASDiscText.Text <> "") Then
            Try
                datRegRASDate = CDate(RegRASDiscText.Text)
            Catch ex As Exception
                RegRASDiscText.BackColor = Drawing.Color.Yellow
                RegRASDiscText.Text = ""
                GlobalCPLBlind.BackColor = Drawing.Color.Yellow
                GlobalRASDisc.BackColor = Drawing.Color.Yellow

                lblErrorMessage.Text = "The date you entered is not a valid date!"
                lblErrorMessage.Visible = True

                Exit Sub
            End Try
        End If

        If (datRegRASDate < datGlobalCPLBlind) Or (datRegRASDate > datGlobalRASDate) Then
            RegRASDiscText.BackColor = Drawing.Color.Yellow
            RegRASDiscText.Text = ""
            GlobalCPLBlind.BackColor = Drawing.Color.Yellow
            GlobalRASDisc.BackColor = Drawing.Color.Yellow

            'Response.Write("<div style='position: absolute; top: 10px; left: 100px'>The Regional CPL Blind Date you entered does not fit within the range of the Global dates!</div>")
            lblErrorMessage.Text = "The date you entered does not fit within the range of the Global dates!"
            lblErrorMessage.Visible = True
        Else
            RegRASDiscText.BackColor = Drawing.Color.White
            RegRASDiscLabel.Text = RegRASDiscText.Text
            GlobalCPLBlind.BackColor = GPGDesc.BackColor
            GlobalRASDisc.BackColor = GPGDesc.BackColor

            lblErrorMessage.Visible = False
        End If

        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                RegRASDiscText.Text, "4", p_ProductBrandID.Text, _
                                                FeatureCategoryID.Text, "1", "1")

    End Sub

    Protected Sub chkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged

        Dim i As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim chk As CheckBox = sender
        Dim j As Integer
        'Dim Row As GridViewRow = chk.NamingContainer

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            Dim strTest As String = "" 'sender.ToString()
            strTest = chk.Parent.Parent.ToString()

            'Dim k As Integer
            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim AVRegionalDatesID As Label = Row.FindControl("lblAVRegionalDatesID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim RegCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
            Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
            Dim RegRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
            Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3
            Dim GPGDesc As Label = Row.FindControl("lblGPGDesc") '-4
            Dim ProdName As Label = Row.FindControl("lblProdName") '-5
            Dim AVNo As Label = Row.FindControl("lblAVNo") '-6
            strTest = AVNo.Text
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
            Dim GlobalSeriesConfigEOL As Label = Row.FindControl("lblGlobalSeriesConfigEOL") '18
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20

            If (strCurrFeatCat <> FeatureCat.Text) Then
                strCurrFeatCat = FeatureCat.Text
            Else
                If (chkAll.Checked = False) Then
                    chkSel.Checked = False

                    'RegCPLBlindLabel.Visible = True
                    'RegCPLBlindText.Visible = False
                    'RegRASDiscLabel.Visible = True
                    'RegRASDiscText.Visible = False

                    If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                        Row.CssClass = "O"

                        RegCPLBlindLabel.Font.Strikeout = True
                        RegRASDiscLabel.Font.Strikeout = True

                        'UpdateAvRegionalDates
                        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                                        AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                        RegRASDiscText.Text, "3", p_ProductBrandID.Text, _
                                                                        FeatureCategoryID.Text, "1", "0")

                        RecSelected.Text = "N"
                    Else
                        RegCPLBlindLabel.Font.Strikeout = False
                        RegRASDiscLabel.Font.Strikeout = False
                    End If

                    RegCPLBlindLabel.Visible = True
                    RegCPLBlindLabel.CssClass = "0"
                    RegCPLBlindText.Visible = False
                    RegRASDiscLabel.Visible = True
                    RegRASDiscLabel.CssClass = "0"
                    RegRASDiscText.Visible = False

                    RegCPLBlindLabel.Text = RegCPLBlindText.Text
                    RegRASDiscLabel.Text = RegRASDiscText.Text

                    GlobalCPLBlind.BackColor = GPGDesc.BackColor
                    GlobalRASDisc.BackColor = GPGDesc.BackColor
                Else
                    chkSel.Checked = True

                    'RegCPLBlindLabel.Visible = False
                    'RegCPLBlindText.Visible = True
                    'RegRASDiscLabel.Visible = False
                    'RegRASDiscText.Visible = True

                    If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                                        AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                        RegRASDiscText.Text, "4", p_ProductBrandID.Text, _
                                                                        FeatureCategoryID.Text, "1", "1")

                        RecSelected.Text = "Y"
                        Row.CssClass = ""
                    Else
                        HPQ.Excalibur.SupplyChain.InsertAvRegionalDates(sSelRegion, AvDetail_ProdBrandID.Text, GlobalCPLBlind.Text, _
                                                                        GlobalRASDisc.Text, "2", p_ProductBrandID.Text, _
                                                                        FeatureCategoryID.Text)

                        CheckedRec.Text = "Y"
                        RecSelected.Text = "Y"
                    End If

                    'InsertAvRegionalDates
                    RegCPLBlindLabel.Visible = False
                    RegCPLBlindText.Visible = True
                    RegRASDiscLabel.Visible = False
                    RegRASDiscText.Visible = True
                    RegCPLBlindText.BackColor = Drawing.Color.White
                    RegRASDiscText.BackColor = Drawing.Color.White
                End If
            End If
        Next i

    End Sub

    'Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveChanges.Click

    'Dim i As Integer
    ''Dim k As Integer
    'Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
    'Dim strCurrFeatCat As String = ""
    'Dim dt As New DataTable()
    'Dim dr As DataRow
    'Dim dcID As New DataColumn("AvRegionalDatesID") '-0
    'dt.Columns.Add(dcID)
    'Dim dcGeoID As New DataColumn("GeoID") '-1
    'dt.Columns.Add(dcGeoID)
    'Dim dcAVDetailProdBrandID As New DataColumn("MainAvDetProdBrandID") '-2
    'dt.Columns.Add(dcAVDetailProdBrandID)
    'Dim dcRegCPLBlind As New DataColumn("RegionalCPLBlindDate") '-3
    'dt.Columns.Add(dcRegCPLBlind)
    'Dim dcRegRASDics As New DataColumn("RegionalRASDiscDate") '-4
    'dt.Columns.Add(dcRegRASDics)
    'Dim dcStatus As New DataColumn("Status") '-5
    'dt.Columns.Add(dcStatus)
    'Dim dcp_ProdBrandID As New DataColumn("p_ProductBrandID") '-6
    'dt.Columns.Add(dcp_ProdBrandID)
    'Dim dcAvFeatureCatID As New DataColumn("AvFeatureCategoryID") '-7
    'dt.Columns.Add(dcAvFeatureCatID)

    'lblErrorMessage.Text = ""

    'For i = 0 To (intRowCount - 1)
    '    Dim Row As GridViewRow
    '    Row = gvRegAVSelToolGrid.Rows(i)

    '    Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
    '    Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
    '    Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
    '    Dim RegCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
    '    Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
    '    Dim RegRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
    '    Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3
    '    Dim p_ProductBrandID As Label = Row.FindControl("lblProductBrandID") '-14
    '    Dim FeatureCategoryID As Label = Row.FindControl("lblAvFeatureCatID") '-15
    '    Dim GeoID As Label = Row.FindControl("lblGeoID") '16
    '    Dim AvDetail_ProdBrandID As Label = Row.FindControl("lblAvDetail_ProductBrandID") '17

    '    If (strCurrFeatCat <> FeatureCat.Text) Then
    '        strCurrFeatCat = FeatureCat.Text
    '    ElseIf (RegCPLBlindText.Text = "") Then
    '        lblErrorMessage.Text = "Date Missing! Enter a Regional SA Date and try again!"
    '        lblErrorMessage.Visible = True

    '        Exit Sub
    '    ElseIf (RegRASDiscText.Text = "") Then
    '        lblErrorMessage.Text = "Date Missing! Enter a Regional EM Date and try again!"
    '        lblErrorMessage.Visible = True

    '        Exit Sub
    '    Else
    '        If (chkSel.Checked = True) Then
    '            dr = dt.NewRow()

    '            dr("AvRegionalDatesID") = "0"
    '            dr("GeoID") = sSelRegion
    '            dr("MainAvDetProdBrandID") = AvDetail_ProdBrandID.Text

    '            If (chkSel.Checked = False) Then
    '                dr("RegionalCPLBlindDate") = RegCPLBlindLabel.Text
    '                dr("RegionalRASDiscDate") = RegRASDiscLabel.Text
    '            Else
    '                dr("RegionalCPLBlindDate") = RegCPLBlindText.Text
    '                dr("RegionalRASDiscDate") = RegRASDiscText.Text
    '            End If

    '            dr("Status") = "2"
    '            dr("p_ProductBrandID") = p_ProductBrandID.Text
    '            dr("AvFeatureCategoryID") = FeatureCategoryID.Text


    '            dt.Rows.Add(dr)
    '        End If
    '    End If
    'Next i

    'HPQ.Excalibur.SupplyChain.InsertAvRegionalDates(dt, sProdBrandIDs, sCatIDs, sSelRegion)
    'HPQ.Excalibur.SupplyChain.InsertAvPlantDates_OneRecAtATtime(dt, 
    'lblErrorMessage.Text = "Data has been saved Successfully!"
    'lblErrorMessage.Visible = True

    'ProcessFilter(sProdBrandIDs, sCatIDs, sSelRegion)

    'End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Dim i As Integer
        Dim intRowCount As Integer = gvRegAVSelToolGrid.Rows.Count
        Dim strCurrFeatCat As String = ""
        Dim j As Integer

        For i = 0 To (intRowCount - 1)
            Dim Row As GridViewRow
            Row = gvRegAVSelToolGrid.Rows(i)

            'Dim k As Integer
            Dim FeatureCat As Label = Row.FindControl("lblFeatCat")
            Dim AVRegionalDatesID As Label = Row.FindControl("lblAVRegionalDatesID")
            Dim AVDetailID As Label = Row.FindControl("lblAVDetailID") '-0
            Dim chkSel As CheckBox = Row.FindControl("chkSelect") '-1
            Dim RegCPLBlindLabel As Label = Row.FindControl("lblRegCPLBlind") '-2
            Dim RegCPLBlindText As TextBox = Row.FindControl("txtRegCPLBlind") '-2
            Dim RegRASDiscLabel As Label = Row.FindControl("lblRegRASDics") '-3
            Dim RegRASDiscText As TextBox = Row.FindControl("txtRegRASDics") '-3
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
            Dim GlobalSeriesConfigEOL As Label = Row.FindControl("lblGlobalSeriesConfigEOL") '18
            Dim CheckedRec As Label = Row.FindControl("chkCheckedRec") '19
            Dim RecSelected As Label = Row.FindControl("chkRecSelected") '20

            If (strCurrFeatCat <> FeatureCat.Text) Then
                strCurrFeatCat = FeatureCat.Text
            Else
                If (chkSel.Checked = False) Then

                    'RegCPLBlindLabel.Visible = True
                    'RegCPLBlindText.Visible = False
                    'RegRASDiscLabel.Visible = True
                    'RegRASDiscText.Visible = False

                    If (CheckedRec.Text = "Y") Then 'This means its been previously checked before and now it needs to have a strikethrough it.
                        Row.CssClass = "O"

                        RegCPLBlindLabel.Font.Strikeout = True
                        RegRASDiscLabel.Font.Strikeout = True

                        'UpdateAvRegionalDates
                        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                        RegRASDiscText.Text, "3", p_ProductBrandID.Text, FeatureCategoryID.Text, "1", "0")

                        RecSelected.Text = "N"
                    Else
                        RegCPLBlindLabel.Font.Strikeout = False
                        RegRASDiscLabel.Font.Strikeout = False
                    End If

                    RegCPLBlindLabel.Visible = True
                    RegCPLBlindLabel.CssClass = "0"
                    RegCPLBlindText.Visible = False
                    RegRASDiscLabel.Visible = True
                    RegRASDiscLabel.CssClass = "0"
                    RegRASDiscText.Visible = False

                    RegCPLBlindLabel.Text = RegCPLBlindText.Text
                    RegRASDiscLabel.Text = RegRASDiscText.Text

                    GlobalCPLBlind.BackColor = GPGDesc.BackColor
                    GlobalRASDisc.BackColor = GPGDesc.BackColor
                Else
                    If (CheckedRec.Text = "Y") Then 'This means that the user just checked it but its been saved in the database before so an update to the record needs to happen, not an insert.
                        HPQ.Excalibur.SupplyChain.UpdateAvRegionalDates(AVRegionalDatesID.Text, sSelRegion, _
                                                                        AvDetail_ProdBrandID.Text, RegCPLBlindText.Text, _
                                                                        RegRASDiscText.Text, "4", p_ProductBrandID.Text, _
                                                                        FeatureCategoryID.Text, "1", "1")

                        RecSelected.Text = "Y"
                        Row.CssClass = ""
                    Else
                        HPQ.Excalibur.SupplyChain.InsertAvRegionalDates(sSelRegion, AvDetail_ProdBrandID.Text, GlobalCPLBlind.Text, _
                                                                        GlobalRASDisc.Text, "2", p_ProductBrandID.Text, _
                                                                        FeatureCategoryID.Text)

                        CheckedRec.Text = "Y"
                        RecSelected.Text = "Y"
                    End If

                    'InsertAvRegionalDates
                    RegCPLBlindLabel.Visible = False
                    RegCPLBlindText.Visible = True
                    RegRASDiscLabel.Visible = False
                    RegRASDiscText.Visible = True
                    RegCPLBlindText.BackColor = Drawing.Color.White
                    RegRASDiscText.BackColor = Drawing.Color.White
                End If
            End If
        Next i

    End Sub

    Protected Sub lkbChangeSelection_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lkbChangeSelection.Click
        ProcessFilter(sProdBrandIDs, sCatIDs, sSelRegion)
    End Sub
End Class





