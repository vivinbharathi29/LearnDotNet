Imports System.Data
Partial Class FilterByMktCampaignsAVSelector
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    Public _GeoID As String
    Public _MktCampRowCount As Integer

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
            'Return "1"
        End Get
    End Property

    Public ReadOnly Property GeoID2() As String
        Get
            '#If DEBUG Then
            '                        Return "Americas"
            '#End If
            Return Request("GeoID2")
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            '#If DEBUG Then
            '            Return "1050"
            '#End If
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            '#If DEBUG Then
            '            Return "763"
            '#End If
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property ProdBrands() As String
        Get
            Return Request("ProdBrands")
        End Get
    End Property

    Public ReadOnly Property Categories() As String
        Get
            Return Request("Categories")
        End Get
    End Property

    Public Property GeoID() As String
        Get
            '#If DEBUG Then
            '            Return "Americas"
            '#End If
            Return _GeoID
        End Get
        Set(ByVal value As String)
            _GeoID = value
        End Set
    End Property

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public Property MktCampRowCount() As Integer
        Get
            Return (GetSessionStateValue("MktCampRowCount"))
        End Get
        Set(ByVal value As Integer)
            AddSessionStateValue("MktCampRowCount", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim strGeoID As String = ""

                'strBusinessID = Request.QueryString("BusinessID").ToString()

                Select Case GeoID2
                    Case "Americas"
                        strGeoID = "1"
                    Case "EMEA"
                        strGeoID = "2"
                    Case "APJ"
                        strGeoID = "3"
                End Select
                _GeoID = strGeoID

                FilterByProductType(strGeoID) 'Default is Commercial

                'lblHeader.Text = "BID: " & BID
                lblHeader.Text = "Please Select a plant and category To Filter By!  (Cateory is optional)"

                'Dim bolEdit As Boolean = False
                'Dim bolActive As Boolean = False

                'HyperLink1.Target = "/SCM/MktCampaigns Edit Screen.aspx?EditType=" & bolEdit & "&GeoID=" & GeoID2 _
                '& "&ActiveCamp=" & bolActive & "&bolInsertRec=" & "'True'"
                HyperLink1.Visible = False

                'TestLabel.Text = "PVID: " & PVID
                TestLabel.Visible = False

                'txtGeoID.Text = "ID: " & strGeoID & " and String Value: " & GeoID2
                txtGeoID.Text = GeoID2
                'txtGeoID.Visible = True
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Private Sub FilterByProductType(ByVal strGeoID As String)
        Try
            TestLabel.Visible = False

            'Dim strGeoID As String

            'strGeoID = "1"
            'strGeoID = Request.QueryString("GeoID").ToString()

            'Load the GoTo Marketing Campaigns dropdown box.
            Dim dtMktCampList As DataTable
            dtMktCampList = HPQ.Excalibur.SupplyChain.ListAllMktCamps_ByRegion("1", strGeoID)
            MktCampRowCount = dtMktCampList.Rows.Count
            cboMktCamp.Items.Clear()
            cboMktCamp.DataSource = dtMktCampList
            cboMktCamp.DataBind()

            'Load the Av Feature Categories listbox - Only show the categories that actually exist in the selected Region.
            Dim dtCategories As DataTable
            dtCategories = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView_WithOutCats(PVID, BID, strGeoID)
            lbCategories.Items.Clear()
            lbCategories.DataSource = ProductParentCategoriesTable(dtCategories, BID)
            lbCategories.DataBind()

            Call PopulateHiddenFields()

        Catch ex As Exception
            TestLabel.Text = ex.ToString
            TestLabel.Visible = True
        End Try
    End Sub

    Private Function ProductParentCategoriesTable(ByVal dtProdBrands As DataTable, ByVal strProdBrands As String) As DataTable

        Dim strFeatCatID As String = ""
        Dim strTest As String = ""
        Dim dr2 As Integer = 0
        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim dcAVFeatureCategoryID As New DataColumn("AVFeatureCategoryID") 'lblAvFeatureCatID
        dt.Columns.Add(dcAVFeatureCategoryID)
        Dim dcp_AvFeatureCategory As New DataColumn("AvFeatureCategory") 'lblProductBrandID
        dt.Columns.Add(dcp_AvFeatureCategory)

        If (dtProdBrands.Rows.Count > 0) Then
            For dr2 = 0 To (dtProdBrands.Rows.Count - 1)
                'First check if the last Feature Category ID that was added to the datatable is different from the current
                'Feature Category ID.  If it is different then a new record needs to be added.
                If Not IsDBNull(dtProdBrands.Rows(dr2).Item("ProductBrandID").ToString()) And (dtProdBrands.Rows(dr2).Item("ProductBrandID").ToString() <> "") Then
                    If (strFeatCatID <> dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()) Then
                        dr = dt.NewRow()

                        strTest = dtProdBrands.Rows(dr2).Item("ParentCategoryID").ToString()
                        If IsDBNull(dtProdBrands.Rows(dr2).Item("ParentCategoryID").ToString()) Or (dtProdBrands.Rows(dr2).Item("ParentCategoryID").ToString() = "") Then
                            strTest = dtProdBrands.Rows(dr2).Item("FeatureCategoryID").ToString()
                            dr("AVFeatureCategoryID") = dtProdBrands.Rows(dr2).Item("FeatureCategoryID").ToString()
                        Else
                            strTest = dtProdBrands.Rows(dr2).Item("ParentCategoryID").ToString()
                            dr("AVFeatureCategoryID") = dtProdBrands.Rows(dr2).Item("ParentCategoryID").ToString()
                        End If
                        strTest = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                        dr("AvFeatureCategory") = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()
                        strFeatCatID = dtProdBrands.Rows(dr2).Item("AvFeatureCategory").ToString()

                        dt.Rows.Add(dr)
                    End If
                End If
            Next dr2
        End If

        ProductParentCategoriesTable = dt

    End Function

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        'Try
        Dim applicationRoot As String = Session("ApplicationRoot")

        Dim bolFound As Boolean = False
        Dim sProdVerID As String = ""
        Dim sProdBrandID As String = ""
        Dim sCategories As String = ""
        Dim strSelRegion As String = "-1"
        Dim sMktCampID As String = ""
        Dim sMktCampName As String = ""
        Dim sMktCampStartDate As String = ""
        Dim sMktCampEndDate As String = ""
        Dim item As ListItem
        Dim intPos As Integer = -1
        Dim intPos2 As Integer = -1

        For Each item In lbCategories.Items
            If item.Selected Then
                If sCategories = "-1" Then
                    sCategories = item.Value
                Else
                    sCategories = sCategories & "," & item.Value
                End If
            End If
        Next
        If (sCategories = "") Then sCategories = "-1"

        sMktCampID = cboMktCamp.SelectedValue
        sMktCampName = cboMktCamp.SelectedItem.ToString()

        Dim dt As DataTable
        'dt = HPQ.Excalibur.SupplyChain.MktCampaigns_Dates(sMktCampID)
        dt = HPQ.Excalibur.SupplyChain.MktCampaigns_GetAllDataForSingleRec(sMktCampID)

        If (dt.Rows.Count > 0) Then
            sMktCampStartDate = dt.Rows(0).Item("MktStartDate").ToString()
            sMktCampEndDate = dt.Rows(0).Item("MktEndDate").ToString()
            txtPlantID.Text = dt.Rows(0).Item("PlantID").ToString()
            txtPlantName.Text = dt.Rows(0).Item("PlantName").ToString()
        End If

        ProcessFilter(txtPlantID.Text, txtPlantName.Text, sCategories, sMktCampID, sMktCampName, sMktCampStartDate, sMktCampEndDate)

        'Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    End Sub

    Private Function ProcessFilter(ByVal sPlantID As String, ByVal sPlantName As String, ByVal sCategory As String, _
                                   ByVal sMktCampID As String, ByVal sMktCampName As String, ByVal sMktStartDate As String, _
                                   ByVal sMktEndDate As String) As Boolean
        Try
            Dim URL As String = Nothing
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sPlantID _
                                                               & ":" & sPlantName & ":" & sCategory & ":" & sMktCampID _
                                                               & ":" & sMktCampName & ":" & sMktStartDate & ":" & sMktEndDate))
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub cmdCancel_onclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Response.Write("<script language='javascript'> {window.close();}</script>")

        Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    End Sub

    Public Sub PopulateHiddenFields()

        Dim strSelRegion As String = "-1"
        Dim sMktCampID As String = ""
        Dim sMktCampName As String = ""
        Dim sMktCampStartDate As String = ""
        Dim sMktCampEndDate As String = ""
        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim bolActive As Boolean = False
        Dim bolEdit As Boolean = True

        sMktCampID = cboMktCamp.SelectedValue
        txtMktCampID.Text = sMktCampID
        sMktCampName = cboMktCamp.SelectedItem.ToString()
        txtMktCampName.Text = sMktCampName

        Dim dt As DataTable
        dt = HPQ.Excalibur.SupplyChain.MktCampaigns_GetAllDataForSingleRec(sMktCampID)

        If (dt.Rows.Count > 0) Then
            strSelRegion = dt.Rows(0).Item("GeoID").ToString()
            txtGeoID.Text = dt.Rows(0).Item("GeoID").ToString()
            sMktCampStartDate = dt.Rows(0).Item("MktStartDate").ToString()
            txtMktStartDate.Text = dt.Rows(0).Item("MktStartDate").ToString()
            sMktCampEndDate = dt.Rows(0).Item("MktEndDate").ToString()
            txtMktEndDate.Text = dt.Rows(0).Item("MktEndDate").ToString()
            bolActive = dt.Rows(0).Item("Active")
            chkActive.Checked = bolActive
            sPlantID = dt.Rows(0).Item("PlantID").ToString()
            txtPlantID.Text = dt.Rows(0).Item("PlantID").ToString()
            sPlantName = dt.Rows(0).Item("PlantName").ToString()
            txtPlantName.Text = dt.Rows(0).Item("PlantName").ToString()
        End If

    End Sub

    Protected Sub btnEditCamps_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditCamps.Click

        If (MktCampRowCount = 0) Then
            TestLabel.Text = "There are no Campaigns to edit!"
            TestLabel.Visible = True
        Else
            TestLabel.Visible = False

            Dim URL As String = Nothing
            Dim applicationRoot As String = Session("ApplicationRoot")
            Dim strSelRegion As String = "-1"
            Dim sMktCampID As String = ""
            Dim sMktCampName As String = ""
            Dim sMktCampStartDate As String = ""
            Dim sMktCampEndDate As String = ""
            Dim sPlantID As String = ""
            Dim sPlantName As String = ""
            Dim bolActive As Boolean = False
            Dim bolEdit As Boolean = True

            sMktCampID = cboMktCamp.SelectedValue
            sMktCampName = cboMktCamp.SelectedItem.ToString()
            Dim dt As DataTable
            dt = HPQ.Excalibur.SupplyChain.MktCampaigns_GetAllDataForSingleRec(sMktCampID)

            If (dt.Rows.Count > 0) Then
                strSelRegion = dt.Rows(0).Item("GeoID").ToString()
                txtGeoID.Text = dt.Rows(0).Item("GeoID").ToString()
                sMktCampStartDate = dt.Rows(0).Item("MktStartDate").ToString()
                txtMktStartDate.Text = dt.Rows(0).Item("MktStartDate").ToString()
                sMktCampEndDate = dt.Rows(0).Item("MktEndDate").ToString()
                txtMktEndDate.Text = dt.Rows(0).Item("MktEndDate").ToString()
                bolActive = dt.Rows(0).Item("Active")
                chkActive.Checked = bolActive
                sPlantID = dt.Rows(0).Item("PlantID").ToString()
                txtPlantID.Text = dt.Rows(0).Item("PlantID").ToString()
                sPlantName = dt.Rows(0).Item("PlantName").ToString()
                txtPlantName.Text = dt.Rows(0).Item("PlantName").ToString()
            End If

            URL = applicationRoot & "/SCM/MktCampaigns Edit Screen Frame.asp?EditType=" & bolEdit & "&CampaignID=" & sMktCampID & "&CampaignName=" & sMktCampName _
            & "&GeoID=" & txtGeoID.Text & "&StartDate=" & sMktCampStartDate & "&EndDate=" & sMktCampEndDate _
            & "&ActiveCamp=" & bolActive & "&bolInsertRec=False" & "&PlantID=" & sPlantID & "&PlantName=" & sPlantName

            'Response.Write("<script language='javascript'> { location.href = '" & URL & "';}</script>")
            Response.Write("<script language='javascript'> { window.open('" & URL & "', '_blank', 'toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=700, height=625');}</script>")
        End If

    End Sub

    Protected Sub btnAddCamps_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddCamps.Click

        Dim URL As String = Nothing
        Dim applicationRoot As String = Session("ApplicationRoot")
        Dim bolActive As Boolean = False

        'URL = applicationRoot & "/SCM/MktCampaigns Edit Screen Frame.asp?GeoID=" & GeoID2 _
        '& "&ActiveCamp=" & bolActive & "&bolInsertRec=True"
        URL = applicationRoot & "/SCM/MktCampaigns Edit Screen Frame.asp?&GeoID3=" & txtGeoID.Text _
        & "&ActiveCamp=" & bolActive & "&bolInsertRec=True" & "&GeoID=" & txtGeoID.Text
        TestLabel.Text = URL
        'TestLabel.Visible = True

        'Response.Write("<script language='javascript'> { location.href = '" & URL & "';}</script>")
        Response.Write("<script language='javascript'> { window.open('" & URL & "', '_blank', 'toolbar=no, location=yes, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=700, height=625');}</script>")
        'Response.Write("<script language='javascript'> { var retValue; retValue = window.parent.showModalDialog('" & URL & "', '_blank', 'toolbar=no, location=yes, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=650, height=500');}</script>")

    End Sub

    Protected Sub cboMktCamp_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMktCamp.SelectedIndexChanged

        Call PopulateHiddenFields()

    End Sub
End Class
