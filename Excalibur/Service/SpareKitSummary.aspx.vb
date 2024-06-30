Imports System.Data


Partial Class Search_SpareKitSummary
    Inherits System.Web.UI.Page

    Private Const EXPORT_EXCEL As String = "1"
    Private Const EXPORT_WORD As String = "2"

    Private Const ALL As Integer = 0
    Private Const NA As Integer = 1
    Private Const LA As Integer = 4
    Private Const APJ As Integer = 3
    Private Const EMEA As Integer = 2

    Private Property ProductName() As String
        Get
            Return ViewState("ProductName")
        End Get
        Set(ByVal value As String)
            ViewState("ProductName") = value
        End Set
    End Property

    Private Property ProductType() As String
        Get
            Return ViewState("ProductType")
        End Get
        Set(ByVal value As String)
            ViewState("ProductType") = value
        End Set
    End Property

    Private Property GeoSKU() As String
        Get
            Return ViewState("GeoSKU")
        End Get
        Set(ByVal value As String)
            ViewState("GeoSKU") = value
        End Set
    End Property

    Private Property GeoSPSNA() As String
        Get
            Return ViewState("GeoNA")
        End Get
        Set(ByVal value As String)
            ViewState("GeoNA") = value
        End Set
    End Property

    Private Property GeoSPSLA() As String
        Get
            Return ViewState("GeoLA")
        End Get
        Set(ByVal value As String)
            ViewState("GeoLA") = value
        End Set
    End Property

    Private Property GeoSPSAPJ() As String
        Get
            Return ViewState("GeoAPJ")
        End Get
        Set(ByVal value As String)
            ViewState("GeoAPJ") = value
        End Set
    End Property

    Private Property GeoSPSEMEA() As String
        Get
            Return ViewState("GeoEMEA")
        End Get
        Set(ByVal value As String)
            ViewState("GeoEMEA") = value
        End Set
    End Property

    'Private Property SkuStartDate() As String
    '    Get
    '        Return ViewState("SkuStartDate")
    '    End Get
    '    Set(ByVal value As String)
    '        ViewState("SkuStartDate") = value
    '    End Set
    'End Property

    'Private Property SkuEndDate() As String
    '    Get
    '        Return ViewState("SkuEndDate")
    '    End Get
    '    Set(ByVal value As String)
    '        ViewState("SkuEndDate") = value
    '    End Set
    'End Property

    Private Property SpsStartDate() As String
        Get
            Return ViewState("SpsStartDate")
        End Get
        Set(ByVal value As String)
            ViewState("SpsStartDate") = value
        End Set
    End Property

    Private Property SpsEndDate() As String
        Get
            Return ViewState("SpsEndDate")
        End Get
        Set(ByVal value As String)
            ViewState("SpsEndDate") = value
        End Set
    End Property

    Private Property SpareKitCategories() As String
        Get
            Return ViewState("SpareKitCategories")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitCategories") = value
        End Set
    End Property

    Private Property OSSP() As String
        Get
            Return ViewState("OSSP")
        End Get
        Set(ByVal value As String)
            ViewState("OSSP") = value
        End Set
    End Property

    Private Property ServiceFamilyPartNumber() As String
        Get
            Return ViewState("ServiceFamilyPartNumber")
        End Get
        Set(ByVal value As String)
            ViewState("ServiceFamilyPartNumber") = value
        End Set
    End Property

    Private Property SpareKitPartNumbers() As String
        Get
            Return ViewState("SpareKitPartNyumbers")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitPartNyumbers") = value
        End Set
    End Property


    Private Property SKUNumber() As String
        Get
            Return ViewState("SKUNumber")
        End Get
        Set(ByVal value As String)
            ViewState("SKUNumber") = value
        End Set
    End Property

    Private Property KMAT() As String
        Get
            Return ViewState("KMAT")
        End Get
        Set(ByVal value As String)
            ViewState("KMAT") = value
        End Set
    End Property

    Private Property MaxRows() As String
        Get
            Return ViewState("MaxRows")
        End Get
        Set(ByVal value As String)
            ViewState("MaxRows") = value
        End Set
    End Property

    Const FIELD_NAME_DATAGRID_ORDER As String = "CategoryName"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            Response.CacheControl = "No-cache"
            Response.Clear()
            Response.ClearContent()
            Response.ClearHeaders()


            If Not Page.IsPostBack Then
                pnlNoData.Visible = False
                pnlData.Visible = False

                ' Initial Order and field
                session("OrderField") = FIELD_NAME_DATAGRID_ORDER

                If Not PreviousPage Is Nothing Then
                    Dim sbProductName As StringBuilder = New StringBuilder()
                    Dim sbSpareKitCategory As StringBuilder = New StringBuilder()
                    Dim sbOSSP As StringBuilder = New StringBuilder()
                    Dim sbGeoSKU As StringBuilder = New StringBuilder()
             
                    Dim lstSpareCategory As ListBox = PreviousPage.FindControl("lstSpareCategory")
                    Dim lstProducts As ListBox = PreviousPage.FindControl("lstProducts")
                    Dim lstOSSP As ListBox = PreviousPage.FindControl("lstOSSP")
                    Dim rdProductType As RadioButtonList = PreviousPage.FindControl("rdProductType")
                    Dim chkSPSGeo As CheckBoxList = PreviousPage.FindControl("chkSpsGeo")
                    Dim chkSKUGeo As CheckBoxList = PreviousPage.FindControl("chkSKUGeo")

                    Dim txtSKUNumber As TextBox = PreviousPage.FindControl("txtSKUNumber")
                    Dim txtKMAT As TextBox = PreviousPage.FindControl("txtKmat")
                    Dim txtServiceFamPartNum As TextBox = PreviousPage.FindControl("txtServiceFamPartNum")
                    Dim txtSpareKitNumbers As TextBox = PreviousPage.FindControl("txtSpsNumbers")
                    Dim txtSpsStartDate As TextBox = PreviousPage.FindControl("DatepickerSPSStart")
                    Dim txtSpsEndDate As TextBox = PreviousPage.FindControl("DatepickerSPSEnd")
                    'Dim txtSkuStartDate As TextBox = PreviousPage.FindControl("DatepickerSKUStart")
                    'Dim txtSkuEndDate As TextBox = PreviousPage.FindControl("DatepickerSKUEnd")

                    Dim txtMaxRows As TextBox = PreviousPage.FindControl("txtMaxRows")

                    Dim ddlReportFormat As DropDownList = PreviousPage.FindControl("ddlReportFormat")

                    Select Case ddlReportFormat.SelectedValue
                        Case EXPORT_EXCEL
                            Response.ContentType = "application/vnd.ms-excel"
                            'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        Case EXPORT_WORD
                            Response.ContentType = "application/msword"
                            'Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    End Select

                    ' get the product names selected in a stringbuilder
                    For Each item As ListItem In lstProducts.Items
                        If item.Selected And item.Value <> String.Empty Then
                            sbProductName.Append(item.Value & ",")
                        End If
                    Next

                    ' get the spare kit categories selected in a stringbuilder
                    For Each item As ListItem In lstSpareCategory.Items
                        If item.Selected And item.Value <> String.Empty Then
                            sbSpareKitCategory.Append(item.Value & ",")
                        End If
                    Next

                    ' get the OSSP selected in a stringbuilder
                    For Each item As ListItem In lstOSSP.Items
                        If item.Selected And item.Value <> String.Empty Then
                            sbOSSP.Append(item.Value & ",")
                        End If
                    Next

                    ' get the GeoSKU selected in a stringbuilder
                    For Each item As ListItem In chkSKUGeo.Items
                        If item.Selected And item.Value <> ALL Then
                            sbGeoSKU.Append(item.Value & ",")
                        End If
                    Next

                    Dim sSQlSKUNumber As New StringBuilder
                    If txtSKUNumber.Text <> String.Empty Then
                        Dim sSku As Array = txtSKUNumber.Text.Split(",")

                        For Each sCad As String In sSku
                            If sSku.Length = 1 Then
                                sSQlSKUNumber.Append(sCad.ToString.Trim)
                            Else
                                If sSQlSKUNumber.ToString = String.Empty Then
                                    sSQlSKUNumber.Append(sCad.ToString.Trim)
                                Else
                                    If InStr(sCad.ToString.Trim, sSQlSKUNumber.ToString, CompareMethod.Text) = 0 Then
                                        sSQlSKUNumber.Append(",")
                                        sSQlSKUNumber.Append(sCad.ToString.Trim)
                                    End If
                                End If
                            End If
                        Next
                    End If

                    Dim sSQlKMAT As New StringBuilder
                    If txtKMAT.Text <> String.Empty Then
                        Dim sKMAT As Array = txtKMAT.Text.Split(",")

                        For Each sCad As String In sKMAT
                            If sKMAT.Length = 1 Then
                                sSQlKMAT.Append(sCad.ToString.Trim)
                            Else
                                If sSQlKMAT.ToString = String.Empty Then
                                    sSQlKMAT.Append(sCad.ToString.Trim)
                                Else
                                    If InStr(sCad.ToString.Trim, sSQlKMAT.ToString, CompareMethod.Text) = 0 Then
                                        sSQlKMAT.Append(",")
                                        sSQlKMAT.Append(sCad.ToString.Trim)
                                    End If
                                End If
                            End If
                        Next
                    End If

                    Dim sSQlFamPartNum As New StringBuilder
                    If txtServiceFamPartNum.Text <> String.Empty Then
                        Dim sServiceFamPartNum As Array = txtServiceFamPartNum.Text.Split(",")

                        For Each sCad As String In sServiceFamPartNum
                            If sServiceFamPartNum.Length = 1 Then
                                sSQlFamPartNum.Append(sCad.ToString.Trim)
                            Else
                                If sSQlFamPartNum.ToString = String.Empty Then
                                    sSQlFamPartNum.Append(sCad.ToString.Trim)
                                Else
                                    'look if this element already exists in the string
                                    If InStr(sCad.ToString.Trim, sSQlFamPartNum.ToString, CompareMethod.Text) = 0 Then
                                        sSQlFamPartNum.Append(",")
                                        sSQlFamPartNum.Append(sCad.ToString.Trim)
                                    End If
                                End If
                            End If
                        Next
                    End If

                    Dim sbSpsPartNumbers As StringBuilder = New StringBuilder()
                    If txtSpareKitNumbers.Text <> String.Empty Then
                        Dim sSpsPartNumbers As Array = txtSpareKitNumbers.Text.Split(",")
                        ' get the OSSP selected in a stringbuilder
                        For Each sCad As String In sSpsPartNumbers
                            sbSpsPartNumbers.Append(sCad.ToString.Trim & ",")
                        Next
                    End If


                    GeoSPSNA = String.Empty
                    GeoSPSLA = String.Empty
                    GeoSPSAPJ = String.Empty
                    GeoSPSEMEA = String.Empty

                    For Each item As ListItem In chkSPSGeo.Items
                        If item.Selected = True And item.Value <> ALL Then
                            Select Case item.Value
                                Case NA
                                    GeoSPSNA = item.Value
                                Case LA
                                    GeoSPSLA = item.Value
                                Case APJ
                                    GeoSPSAPJ = item.Value
                                Case EMEA
                                    GeoSPSEMEA = item.Value
                            End Select
                        End If
                    Next

                    ' delete the last ','
                    If sbProductName.Length > 0 Then sbProductName.Remove(sbProductName.Length - 1, 1)
                    If sbSpareKitCategory.Length > 0 Then sbSpareKitCategory.Remove(sbSpareKitCategory.Length - 1, 1)
                    If sbGeoSKU.Length > 0 Then sbGeoSKU.Remove(sbGeoSKU.Length - 1, 1)
                    If sbOSSP.Length > 0 Then sbOSSP.Remove(sbOSSP.Length - 1, 1)
                    If sbSpsPartNumbers.Length > 0 Then sbSpsPartNumbers.Remove(sbSpsPartNumbers.Length - 1, 1)

                    ' fill the property with the value
                    ProductName = sbProductName.ToString()
                    SpareKitCategories = sbSpareKitCategory.ToString()
                    OSSP = sbOSSP.ToString
                    SpareKitPartNumbers = sbSpsPartNumbers.ToString
                    GeoSKU = sbGeoSKU.ToString()

                    SKUNumber = sSQlSKUNumber.ToString.Trim
                    KMAT = sSQlKMAT.ToString.Trim
                    ServiceFamilyPartNumber = sSQlFamPartNum.ToString.Trim
                    ProductType = rdProductType.SelectedValue

                    If txtMaxRows.Text = "" Then MaxRows = String.Empty Else MaxRows = txtMaxRows.Text
                    If txtSpsStartDate.Text = "" Then SpsStartDate = String.Empty Else SpsStartDate = txtSpsStartDate.Text
                    If txtSpsEndDate.Text = "" Then SpsEndDate = String.Empty Else SpsEndDate = txtSpsEndDate.Text
                    'If txtSkuStartDate.Text = "" Then SkuStartDate = String.Empty Else SkuStartDate = txtSkuStartDate.Text
                    'If txtSkuEndDate.Text = "" Then SkuEndDate = String.Empty Else SkuEndDate = txtSkuEndDate.Text

                    'Load Data
                    getSpareKits()

                    lblLastRunDate.Text = Date.Now.ToLongDateString()
                Else
                    Response.Write("<h1>You must enter this page through the Service Advanced Search & Report screen.</h1>")
                    Response.End()
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSpareKits()
        Try
            Dim dtData As New DataTable

            ', SkuStartDate, SkuEndDate, , ,

            dtData = HPQ.Excalibur.Service.getServiceSpareKits(SKUNumber, KMAT, ProductName, SpareKitCategories, ServiceFamilyPartNumber, SpareKitPartNumbers, OSSP, ProductType, GeoSKU, GeoSPSNA, GeoSPSLA, GeoSPSAPJ, GeoSPSEMEA, SpsStartDate, SpsEndDate, MaxRows)

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                gvData.DataSource = dtData
                gvData.DataBind()
                Session("dtData") = dtData
            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "There are not Spare Kits for the filters selected."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvData_Sorting(sender As Object, e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvData.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvData.DataSource = dtData
                gvData.DataBind()
            End If
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetSortDirection(ByVal column As String) As String
        Try
            ' By default, set the sort direction to ascending.
            Dim sortDirection = "ASC"

            ' Retrieve the last column that was sorted.
            Dim sortExpression = TryCast(ViewState("SortExpression"), String)

            If sortExpression IsNot Nothing Then
                ' Check if the same column is being sorted.
                ' Otherwise, the default value can be returned.
                If sortExpression = column Then
                    Dim lastDirection = TryCast(ViewState("SortDirection"), String)
                    If lastDirection IsNot Nothing _
                      AndAlso lastDirection = "ASC" Then
                        sortDirection = "DESC"
                    End If
                End If
            End If

            ' Save new values in ViewState.
            ViewState("SortDirection") = sortDirection
            ViewState("SortExpression") = column

            Return sortDirection
        Catch ex As Exception
            Throw ex
        End Try
    End Function


End Class
