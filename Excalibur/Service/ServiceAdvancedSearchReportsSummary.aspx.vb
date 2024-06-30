Imports System.Data
Imports System.IO

Partial Class Service_ServiceAdvancedSearchReportsSummary
    Inherits System.Web.UI.Page

    'Const REPORT_TYPE_SERVICE_BOM_AVNUMBERS As Integer = 1
    Const REPORT_TYPE_AV_TO_SPS As Integer = 2
    'Const REPORT_TYPE_SPS_TO_PRODUCTS As Integer = 3
    Const REPORT_TYPE_SERVICE_SPS_BOM As Integer = 4
    Const REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS As Integer = 5
    Const REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS As Integer = 6
    Const REPORT_TYPE_SPS_BY_CATEGORY As Integer = 7
    Const REPORT_TYPE_USED_BY As Integer = 8
    Const REPORT_TYPE_RSL_CHANGELOG As Integer = 9

    Private Const EXPORT_EXCEL As String = "1"
    Private Const EXPORT_WORD As String = "2"
    Private Const EXPORT_CSV As String = "3"

    Const FIELD_NAME_DATAGRID_ORDER_REPORT_SPS_TO_PRODUCTS As String = "DOTSName"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_AV_TO_SPS As String = "AvNo"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_SERVICE_SPS_BOM As String = "SpareKitNumber"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_SKUs_TO_SPAREKITS As String = "SKU"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_BOM_FROM_AVNUMBERs As String = "AvNumber"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_FAMILY_TO_SKU_TO_SPAREKITS As String = "dotsname"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_SERVICE_SPS_BY_CATEGORY As String = "SpareKitNumber"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_USEDBY_SKU As String = "Sku"
    Const FIELD_NAME_DATAGRID_ORDER_REPORT_RSL_CHANGELOG As String = "ChangeDt"

    Const USEDBY_FAMILYNAME As String = "1"
    Const USEDBY_SPAREKITS As String = "2"
    Const USEDBY_SUBASSEMBLY As String = "3"
    Const USEDBY_COMPONENT As String = "4"


#Region "Properties"

    Private Property ReportExportTO() As Integer
        Get
            Return ViewState("ReportExportTO")
        End Get
        Set(ByVal value As Integer)
            ViewState("ReportExportTO") = value
        End Set
    End Property

    Private Property ReportType() As Integer
        Get
            Return ViewState("ReportType")
        End Get
        Set(ByVal value As Integer)
            ViewState("ReportType") = value
        End Set
    End Property

    Private Property SpareKitNumbers() As String
        Get
            Return ViewState("SpareKitNumbers")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitNumbers") = value
        End Set
    End Property

    Private Property SpareCategoryIDs() As String
        Get
            Return ViewState("SpareCategoryIDs")
        End Get
        Set(ByVal value As String)
            ViewState("SpareCategoryIDs") = value
        End Set
    End Property

    Private Property SubAssemblies() As String
        Get
            Return ViewState("SubAssemblies")
        End Get
        Set(ByVal value As String)
            ViewState("SubAssemblies") = value
        End Set
    End Property

    Private Property Components() As String
        Get
            Return ViewState("Components")
        End Get
        Set(ByVal value As String)
            ViewState("Components") = value
        End Set
    End Property

    Private Property AVNumbers() As String
        Get
            Return ViewState("AVNumbers")
        End Get
        Set(ByVal value As String)
            ViewState("AVNumbers") = value
        End Set
    End Property

    Private Property SKUs() As String
        Get
            Return ViewState("SKUs")
        End Get
        Set(ByVal value As String)
            ViewState("SKUs") = value
        End Set
    End Property

    Private Property ProductVersionIds() As String
        Get
            Return ViewState("ProductVersionIds")
        End Get
        Set(ByVal value As String)
            ViewState("ProductVersionIds") = value
        End Set
    End Property

   
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                pnlNoData.Visible = False
                pnlData.Visible = False

                If Not Request.QueryString("ReportType") Is Nothing Then ReportType = Request.QueryString("ReportType")
                If Not Request.QueryString("ExportType") Is Nothing Then ReportExportTO = Request.QueryString("ExportType")

                Select Case ReportType
                    Case REPORT_TYPE_SPS_BY_CATEGORY
                        lblTitle.Text += "SpareKits By Category"
                        If Not Request.QueryString("txtSpsCategoryIDs") Is Nothing Then SpareCategoryIDs = Request.QueryString("txtSpsCategoryIDs")

                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_SERVICE_SPS_BY_CATEGORY

                    Case REPORT_TYPE_SERVICE_SPS_BOM
                        lblTitle.Text += "SpareKits Bom"

                        If Not Request.QueryString("txtSPSBom") Is Nothing Then SpareKitNumbers = Request.QueryString("txtSPSBom")

                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_SERVICE_SPS_BOM
                    Case REPORT_TYPE_AV_TO_SPS
                        lblTitle.Text += "Av to SpareKits"

                        If Not Request.QueryString("txtAVNumber") Is Nothing Then AVNumbers = Request.QueryString("txtAVNumber")

                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_AV_TO_SPS
                    Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                        lblTitle.Text += "Sku to Sparekits"

                        If Not Request.QueryString("txtSKUNumber") Is Nothing Then SKUs = Request.QueryString("txtSKUNumber")

                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_SKUs_TO_SPAREKITS
                    Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                        lblTitle.Text += "Product Name Bom Report"
                        If Not Request.QueryString("ProductVersionID") Is Nothing Then ProductVersionIds = Request.QueryString("ProductVersionID")
                        If Not Request.QueryString("txtSkuNumberProduct") Is Nothing Then SKUs = Request.QueryString("txtSkuNumberProduct")
                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_FAMILY_TO_SKU_TO_SPAREKITS
                    Case REPORT_TYPE_USED_BY
                        Dim UsedBy As String = Request.QueryString("UsedByType").ToString
                        Select Case UsedBy
                            Case USEDBY_FAMILYNAME
                                lblTitle.Text += "Used By Family Name Report"
                                If Not Request.QueryString("ProductVersionID") Is Nothing Then ProductVersionIds = Request.QueryString("ProductVersionID")
                                If Not Request.QueryString("txtUsedBySkuNumber") Is Nothing Then SKUs = Request.QueryString("txtUsedBySkuNumber")
                            Case USEDBY_SPAREKITS
                                lblTitle.Text += "Used By Sparekit Report"
                                If Not Request.QueryString("txtUsedByNumber") Is Nothing Then SpareKitNumbers = Request.QueryString("txtUsedByNumber")
                            Case USEDBY_SUBASSEMBLY
                                lblTitle.Text += "Used By SubAsemblies Report"
                                If Not Request.QueryString("txtUsedByNumber") Is Nothing Then SubAssemblies = Request.QueryString("txtUsedByNumber")
                            Case USEDBY_COMPONENT
                                lblTitle.Text += "Used By Components Report"
                                If Not Request.QueryString("txtUsedByNumber") Is Nothing Then Components = Request.QueryString("txtUsedByNumber")
                        End Select
                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_USEDBY_SKU
                    Case REPORT_TYPE_RSL_CHANGELOG
                        lblTitle.Text += "RSL ChangeLog"

                        If Not Request.QueryString("ProductVersionID") Is Nothing Then ProductVersionIds = Request.QueryString("ProductVersionID")

                        ' Initial Order and field
                        Session("OrderField") = FIELD_NAME_DATAGRID_ORDER_REPORT_RSL_CHANGELOG
                End Select

                GetReport()
                lblLastRunDate.Text = Date.Now.ToLongDateString()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetReport()
        Try
            Dim dtData As New DataTable

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False

            Select Case ReportType
                Case REPORT_TYPE_SERVICE_SPS_BOM
                    dtData = HPQ.Excalibur.Service.GetServiceReport_SPS_BOM(SpareKitNumbers)
                Case REPORT_TYPE_SPS_BY_CATEGORY
                    dtData = HPQ.Excalibur.Service.GetServiceReport_SPS_By_Category(SpareCategoryIDs)
                Case REPORT_TYPE_AV_TO_SPS
                    dtData = HPQ.Excalibur.Service.GetServiceReport_SPS_From_AvNumbers(AVNumbers)
                Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                    dtData = HPQ.Excalibur.Service.GetServiceReport_SKU_To_Sparekits(SKUs)
                Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                    dtData = HPQ.Excalibur.Service.GetServiceReport_ProductVersion_To_Sku_Sparekits(ProductVersionIds, SKUs)
                Case REPORT_TYPE_USED_BY
                    dtData = HPQ.Excalibur.Service.GetServiceReport_UsedBy(ProductVersionIds, SKUs, SpareKitNumbers, SubAssemblies, Components)
                Case REPORT_TYPE_RSL_CHANGELOG
                    Dim dwExcalibur As New HPQ.Excalibur.Data
                    dtData = dwExcalibur.SelectRslChangeLog(ProductVersionIds)

            End Select

            If dtData.Rows.Count > 0 Then
                Session("dtData") = dtData

                If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                    Select Case ReportType
                        Case REPORT_TYPE_SERVICE_SPS_BOM
                            gvSpareKitsBom.AllowPaging = False
                            gvSpareKitsBom.AllowSorting = False
                        Case REPORT_TYPE_SPS_BY_CATEGORY
                            gvSpareKitsByCategory.AllowPaging = False
                            gvSpareKitsByCategory.AllowSorting = False
                        Case REPORT_TYPE_AV_TO_SPS
                            gvReportAvNumberToSps.AllowPaging = False
                            gvReportAvNumberToSps.AllowSorting = False
                        Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                            gvReportSkuToSpareKits.AllowPaging = False
                            gvReportSkuToSpareKits.AllowSorting = False
                        Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                            gvReportProductToSkuToSparekits.AllowPaging = False
                            gvReportProductToSkuToSparekits.AllowSorting = False
                        Case REPORT_TYPE_USED_BY
                            gvUsedBy.AllowPaging = False
                            gvUsedBy.AllowSorting = False
                        Case REPORT_TYPE_RSL_CHANGELOG
                            gvRSLChangeLog.AllowPaging = False
                            gvRSLChangeLog.AllowSorting = False
                    End Select
                End If

                Select Case ReportType
                    Case REPORT_TYPE_SERVICE_SPS_BOM
                        gvSpareKitsBom.DataSource = dtData
                        gvSpareKitsBom.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvSpareKitsBom)
                        End If
                    Case REPORT_TYPE_SPS_BY_CATEGORY
                        gvSpareKitsByCategory.DataSource = dtData
                        gvSpareKitsByCategory.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvSpareKitsByCategory)
                        End If
                    Case REPORT_TYPE_AV_TO_SPS
                        gvReportAvNumberToSps.DataSource = dtData
                        gvReportAvNumberToSps.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvReportAvNumberToSps)
                        End If
                    Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                        gvReportSkuToSpareKits.DataSource = dtData
                        gvReportSkuToSpareKits.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvReportSkuToSpareKits)
                        End If
                    Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                        gvReportProductToSkuToSparekits.DataSource = dtData
                        gvReportProductToSkuToSparekits.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvReportProductToSkuToSparekits)
                        End If
                    Case REPORT_TYPE_USED_BY
                        gvUsedBy.DataSource = dtData
                        gvUsedBy.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvUsedBy)
                        End If
                    Case REPORT_TYPE_RSL_CHANGELOG
                        gvRSLChangeLog.DataSource = dtData
                        gvRSLChangeLog.DataBind()
                        If ReportExportTO = EXPORT_EXCEL Or ReportExportTO = EXPORT_WORD Or ReportExportTO = EXPORT_CSV Then
                            Export(gvRSLChangeLog)
                        End If
                End Select

            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "No data for the filters selected."
            End If
        Catch ex As Exception
            If InStr(ex.Message.ToUpper, "TIMEOUT", CompareMethod.Text) = 1 Then
                Response.Write("<b>Time Out: The amount of rows to read is too big. Try with another PartNumbers.</b>")
                Response.End()
            Else
                Throw ex

            End If
        End Try
    End Sub

    Private Sub Export(ByVal DataGridView As GridView)
        Try

            'EnableViewState = False
            If ReportExportTO = EXPORT_WORD Then
                Dim FileName As String = String.Empty
                Dim FileNameNumber As Integer = 0
                FileNameNumber = Session("FILENAMENUMBER_WORD")
                If FileNameNumber > 0 Then
                    FileName = "AdvancedSearchReports" + FileNameNumber.ToString + ".doc"
                Else
                    FileName = "AdvancedSearchReports.doc"
                End If
                FileNameNumber = FileNameNumber + 1
                Session("FILENAMENUMBER_WORD") = FileNameNumber

                ExportTo(DataGridView, FileName, "application/msword")
            End If

            If ReportExportTO = EXPORT_CSV Then
                'Response.ClearContent()
                HttpContext.Current.Response.AddHeader("Pragma", "public")
                Response.ContentType = "application/text"

                ExportToCSV(DataGridView)
            End If

            If ReportExportTO = EXPORT_EXCEL Then
                Dim FileName As String = String.Empty
                Dim FileNameNumber As Integer = 0
                FileNameNumber = Session("FILENAMENUMBER_EXCEL")
                If FileNameNumber > 0 Then
                    FileName = "AdvancedSearchReports" + FileNameNumber.ToString + ".xls"
                Else
                    FileName = "AdvancedSearchReports.xls"
                End If
                FileNameNumber = FileNameNumber + 1
                Session("FILENAMENUMBER_EXCEL") = FileNameNumber
                ExportTo(DataGridView, FileName, "application/excel")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ExportTo(ByVal gv As GridView, ByVal FileName As String, ByVal ContentType As String)
        Try
            Response.ClearContent()
            Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName)
            Response.ContentType = ContentType
            Dim sWriter As New StringWriter()
            Dim hWriter As New HtmlTextWriter(sWriter)

            gv.RenderControl(hWriter)

            Response.Write(sWriter.ToString())
            Response.End()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub ExportToCSV(ByVal DataGridView As GridView)
        Try
            Dim stringbuilder As New StringBuilder

            For Each column As DataControlField In DataGridView.Columns
                stringbuilder.Append(column.HeaderText + ",")
            Next
            ' delete the last ','
            If stringbuilder.Length > 0 Then stringbuilder.Remove(stringbuilder.Length - 1, 1)

            stringbuilder.Append(Environment.NewLine)

            For Each elem As GridViewRow In DataGridView.Rows
                For Each col As TableCell In elem.Cells
                    stringbuilder.Append(col.Text.Trim + ",")
                Next
                ' delete the last ','
                If stringbuilder.Length > 0 Then stringbuilder.Remove(stringbuilder.Length - 1, 1)
                stringbuilder.Append(Environment.NewLine)
            Next

            Response.Output.Write(stringbuilder.ToString())
            Response.Flush()
            Response.End()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)

    End Sub

#Region "Sorting"


    Protected Sub gvRSLChangeLog_Sorting(sender As Object, e As GridViewSortEventArgs) Handles gvRSLChangeLog.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvRSLChangeLog.DataSource = dtData
                gvRSLChangeLog.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvSpareKitsByCategory_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvSpareKitsByCategory.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvSpareKitsByCategory.DataSource = dtData
                gvSpareKitsByCategory.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvSpareKitsBom_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvSpareKitsBom.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvSpareKitsBom.DataSource = dtData
                gvSpareKitsBom.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvReportAvNumberToSps_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvReportAvNumberToSps.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvReportAvNumberToSps.DataSource = dtData
                gvReportAvNumberToSps.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvReportSkuToSpareKits_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvReportSkuToSpareKits.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvReportSkuToSpareKits.DataSource = dtData
                gvReportSkuToSpareKits.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Protected Sub gvReportBomFromAvNumbers_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvReportBomFromAvNumbers.Sorting
    '    Try
    '        'Retrieve the table from the session object.
    '        Dim dtData As DataTable = CType(Session("dtData"), DataTable)

    '        If Not dtData Is Nothing Then
    '            dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
    '            gvReportBomFromAvNumbers.DataSource = dtData
    '            gvReportBomFromAvNumbers.DataBind()
    '        End If

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Protected Sub gvReportProductToSkuToSparekits_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvReportProductToSkuToSparekits.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvReportProductToSkuToSparekits.DataSource = dtData
                gvReportProductToSkuToSparekits.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvUsedBy_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvUsedBy.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvUsedBy.DataSource = dtData
                gvUsedBy.DataBind()
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

#End Region

#Region "Pagination"

    Protected Sub gvRSLChangeLog_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gvRSLChangeLog.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvRSLChangeLog.PageIndex = e.NewPageIndex
            gvRSLChangeLog.SelectedIndex = -1
            gvRSLChangeLog.DataSource = dtData
            gvRSLChangeLog.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
   
    Protected Sub gvSpareKitsByCategory_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvSpareKitsByCategory.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvSpareKitsByCategory.PageIndex = e.NewPageIndex
            gvSpareKitsByCategory.SelectedIndex = -1
            gvSpareKitsByCategory.DataSource = dtData
            gvSpareKitsByCategory.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvSpareKitsBom_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvSpareKitsBom.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvSpareKitsBom.PageIndex = e.NewPageIndex
            gvSpareKitsBom.SelectedIndex = -1
            gvSpareKitsBom.DataSource = dtData
            gvSpareKitsBom.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvReportAvNumberToSps_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvReportAvNumberToSps.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvReportAvNumberToSps.PageIndex = e.NewPageIndex
            gvReportAvNumberToSps.SelectedIndex = -1
            gvReportAvNumberToSps.DataSource = dtData
            gvReportAvNumberToSps.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvReportSkuToSpareKits_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvReportSkuToSpareKits.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvReportSkuToSpareKits.PageIndex = e.NewPageIndex
            gvReportSkuToSpareKits.SelectedIndex = -1
            gvReportSkuToSpareKits.DataSource = dtData
            gvReportSkuToSpareKits.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvReportProductToSkuToSparekits_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvReportProductToSkuToSparekits.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvReportProductToSkuToSparekits.PageIndex = e.NewPageIndex
            gvReportProductToSkuToSparekits.SelectedIndex = -1
            gvReportProductToSkuToSparekits.DataSource = dtData
            gvReportProductToSkuToSparekits.DataBind()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvUsedBy_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvUsedBy.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvUsedBy.PageIndex = e.NewPageIndex
            gvUsedBy.SelectedIndex = -1
            gvUsedBy.DataSource = dtData
            gvUsedBy.DataBind()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region



End Class
