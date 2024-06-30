Imports System.Data

Partial Public Class SCM_FilterByPlantsAVSelector
    Inherits System.Web.UI.Page

    Private dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    Public _GeoID As String

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
            'Return "1"
        End Get
    End Property

    Public ReadOnly Property GeoID2() As String
        Get
            Return Request("GeoID2")
            'Return "1"
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
            'Return "1050"
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
            'Return "763"
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
            Return _GeoID
            'Return "Americas"
        End Get
        Set(ByVal value As String)
            _GeoID = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim strGeoID As String

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

                'TestLabel.Text = "PVID: " & PVID
                TestLabel.Visible = False
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Private Sub FilterByProductType(ByVal strGeoID As String)
        Try
            lbPlants.BackColor = Drawing.Color.White
            TestLabel.Visible = False

            'Dim strGeoID As String

            'strGeoID = "1"
            'strGeoID = Request.QueryString("GeoID").ToString()

            'Load the Plants listbox.
            Dim dtPlants As DataTable
            dtPlants = HPQ.Excalibur.SupplyChain.SelectRCTOPlants_ByGeoID(strGeoID)
            lbPlants.Items.Clear()
            lbPlants.DataSource = dtPlants
            lbPlants.DataBind()

            'Load the Av Feature Categories listbox - Only show the categories that actually exist in the selected Region.
            Dim dtCategories As DataTable
            'dtCategories = dw.SelectAvFeatureCategoriesFilter(strBusinessID)
            dtCategories = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView_WithOutCats(PVID, BID, strGeoID)
            'dtCategories = HPQ.Excalibur.SupplyChain.SelectScmDetail_RegionAndPlatformsView_WithOutCats(PVID, BID, strGeoID)
            lbCategories.Items.Clear()
            lbCategories.DataSource = ProductParentCategoriesTable(dtCategories, BID)
            'lbCategories.DataSource = ProductParentCategoriesTable(dtCategories, BID)
            lbCategories.DataBind()

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
        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim sProdVerID As String = ""
        Dim sProdBrandID As String = ""
        Dim sCategories As String = ""
        Dim strSelRegion As String = "-1"
        Dim item As ListItem
        Dim intPos As Integer = -1
        Dim intPos2 As Integer = -1

        For Each item In lbPlants.Items
            If item.Selected Then
                If sPlantID = "" Then
                    sPlantID = item.Value
                Else
                    sPlantID = sPlantID & "," & item.Value
                End If
            End If
        Next

        For Each item In lbPlants.Items
            If item.Selected Then
                If sPlantName = "" Then
                    sPlantName = item.Text
                Else
                    sPlantName = sPlantName & "," & item.Text
                End If
            End If
        Next

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

        If (sPlantName = "") Then
            lbPlants.BackColor = Drawing.Color.Yellow

            TestLabel.Visible = True
            TestLabel.Text = "You must select at least one Plant!"

            Exit Sub
        Else
            lbPlants.BackColor = Drawing.Color.White
            TestLabel.Visible = False
        End If

        ProcessFilter(sPlantID, sPlantName, sCategories)

        'Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    End Sub

    Private Function ProcessFilter(ByVal sPlantID As String, ByVal sPlantName As String, ByVal sCategory As String) As Boolean
        Try
            Dim URL As String = Nothing
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sPlantID & ":" & sPlantName & ":" & sCategory))
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub cmdCancel_onclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Response.Write("<script language='javascript'> {window.close();}</script>")

        Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))
    End Sub

    Protected Sub lbProductBands_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbPlants.SelectedIndexChanged

    End Sub
End Class
