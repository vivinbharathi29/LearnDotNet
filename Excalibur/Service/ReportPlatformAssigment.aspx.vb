﻿Imports System.Data
Imports System.IO

Partial Class Service_ReportPlatformAssigment
    Inherits System.Web.UI.Page

    Private Const EXPORT_EXCEL As String = "1"
    Private Const OPT_NOTEBOOK As String = "0"
    Private Const OPT_DESKTOP As String = "1"
    Private Const DIVISION_NOTEBOOK As String = "1"
    Private Const DIVISION_DESKTOP As String = "2"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                pnlData.Visible = True
                pnlNotebook.Visible = True
                pnlDesktop.Visible = False
                tdPsm.Visible = True
                tdSpdm.Visible = True
                trBusiness.Visible = True
                trDates.Visible = True
                tdProductLine.Visible = False
                btnAddDesktop.Visible = False

                btnAddDesktop.Attributes.Add("onClick", "AddPlatformDesktop();")
                btnReport.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
                btnReport.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
                btnReset.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
                btnReset.Attributes.Add("onmouseout", "ActionCell_onmouseout();")


                'Load Filter Data for Notebook
                getProducts(DIVISION_NOTEBOOK)
                getGPLM()
                getODM()
                getSPDM()
                getPSM()
            End If

            Report()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub rdNotebookDesktop_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdNotebookDesktop.SelectedIndexChanged
        Try
            pnlData.Visible = True
            pnlDesktop.Visible = False
            pnlNotebook.Visible = False
            btnAddDesktop.Visible = False

            lblPlatform.Visible = True
            lstPlatform.Visible = True
            lblDesktopAddCommunication.Visible = False

            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_DESKTOP
                    '  btnAddDesktop.Visible = True
                    pnlDesktop.Visible = True
                    tdPsm.Visible = False
                    tdSpdm.Visible = False
                    trBusiness.Visible = False
                    trDates.Visible = False
                    tdProductLine.Visible = True
                    lstPlatform.Width = 250
                    getProducts(DIVISION_DESKTOP)
                    getProductLines()
                    Report()
                    lblDesktopAddCommunication.Visible = True
                Case OPT_NOTEBOOK
                    btnAddDesktop.Visible = False
                    pnlNotebook.Visible = True
                    tdProductLine.Visible = False
                    tdPsm.Visible = True
                    tdSpdm.Visible = True
                    trBusiness.Visible = True
                    trDates.Visible = True
                    lstPlatform.Width = 140
                    getProducts(DIVISION_NOTEBOOK)
                    Report()
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try

            lstPlatform.SelectedIndex = -1
            lstODM.SelectedIndex = -1
            lstGPLM.SelectedIndex = -1
            txtProjextNumber.Text = String.Empty
            txtServiceFamilyPn.Text = String.Empty

            'Dim Startdate As Date = DateAdd(DateInterval.Day, -30, DateTime.Now)
            'txtStartDate.Value = Startdate.ToString("d")
            'txtEndDate.Value = DateTime.Now.ToString("d")

            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_NOTEBOOK
                    lstSpdm.SelectedIndex = -1
                    lstPsm.SelectedIndex = -1
                    ckBusiness.SelectedIndex = -1
                Case OPT_DESKTOP
                    lstProductLine.SelectedIndex = -1
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReport.Click
        Try
            Report()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvPlatformAssignmentMetrics_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvPlatformAssignmentMetrics.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvPlatformAssignmentMetrics.DataSource = dtData
                gvPlatformAssignmentMetrics.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvPADesktop_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvPADesktop.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvPADesktop.DataSource = dtData
                gvPADesktop.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvPlatformAssignmentMetrics_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvPlatformAssignmentMetrics.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvPlatformAssignmentMetrics.PageIndex = e.NewPageIndex
            gvPlatformAssignmentMetrics.SelectedIndex = -1
            gvPlatformAssignmentMetrics.DataSource = dtData
            gvPlatformAssignmentMetrics.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvPADesktop_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvPADesktop.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvPADesktop.PageIndex = e.NewPageIndex
            gvPADesktop.SelectedIndex = -1
            gvPADesktop.DataSource = dtData
            gvPADesktop.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Report()
        Try
            Dim dtData As New DataTable
            pnlData.Visible = True
            pnlNoData.Visible = False

            Dim sPlatformSelected As StringBuilder = New StringBuilder()
            For Each elem As ListItem In lstPlatform.Items
                If elem.Selected Then
                    sPlatformSelected.Append(elem.Value & ",")
                End If
            Next
            If sPlatformSelected.Length > 0 Then sPlatformSelected.Remove(sPlatformSelected.Length - 1, 1)

            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_NOTEBOOK
                    pnlNotebook.Visible = True
                    dtData = HPQ.Excalibur.Service.GetPlatformAssignmentMetrics(sPlatformSelected.ToString, lstODM.SelectedValue, lstGPLM.SelectedValue, lstSpdm.SelectedValue, lstPsm.SelectedValue, txtServiceFamilyPn.Text, txtProjextNumber.Text, ckBusiness.SelectedValue, txtStartDate.Value, txtEndDate.Value)
                Case OPT_DESKTOP
                    pnlDesktop.Visible = True
                    dtData = HPQ.Excalibur.Service.GetPlatformAssignmentDesktop(sPlatformSelected.ToString, lstODM.SelectedValue, lstGPLM.SelectedValue, lstProductLine.SelectedValue, txtServiceFamilyPn.Text, txtProjextNumber.Text)
            End Select

            If dtData.Rows.Count > 0 Then
                Session("dtData") = dtData

                Select Case rdNotebookDesktop.SelectedValue
                    Case OPT_NOTEBOOK
                        If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                            gvPlatformAssignmentMetrics.AllowPaging = False
                            gvPlatformAssignmentMetrics.AllowSorting = False
                        End If

                        gvPlatformAssignmentMetrics.Visible = True
                        gvPlatformAssignmentMetrics.DataSource = dtData
                        gvPlatformAssignmentMetrics.DataBind()

                        If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                            Export(gvPlatformAssignmentMetrics)
                        End If
                    Case OPT_DESKTOP
                        If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                            gvPADesktop.AllowPaging = False
                            gvPADesktop.AllowSorting = False
                        End If

                        gvPADesktop.Visible = True
                        gvPADesktop.DataSource = dtData
                        gvPADesktop.DataBind()

                        If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                            Export(gvPADesktop)
                        End If
                End Select



            Else
                pnlData.Visible = False
                pnlNotebook.Visible = False
                pnlDesktop.Visible = False
                pnlNoData.Visible = True

                msgSearchNoData.Text = "No data for the filters selected."
            End If

            lblLastRunDate.Text = Date.Now.ToLongDateString()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getProductLines()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.GetProductLines()

            If dtData.Rows.Count > 0 Then
                lstProductLine.DataSource = dtData
                lstProductLine.DataTextField = "Description"
                lstProductLine.DataValueField = "ID"
                lstProductLine.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getProducts(ByVal Division As Integer)
        Try
            Dim dtData As New DataTable

            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_NOTEBOOK
                    dtData = HPQ.Excalibur.Service.GetProductsOnCommodityMatrix(Division)
                    lstPlatform.DataTextField = "dotsname"
                Case OPT_DESKTOP
                    dtData = HPQ.Excalibur.Service.GetProductsOnCommodityMatrix(Division)
                    lstPlatform.DataTextField = "ProductName"
            End Select

            If dtData.Rows.Count > 0 Then
                lstPlatform.DataSource = dtData
                lstPlatform.DataValueField = "ID"
                lstPlatform.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getGPLM()
        Try
            Dim dtData As New DataTable
            Dim objData As New HPQ.Excalibur.Data

            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_NOTEBOOK
                    dtData = objData.ListGplms()
                Case OPT_DESKTOP
                    'dtData = objData.ListGplms()
            End Select

            If dtData.Rows.Count > 0 Then
                lstGPLM.DataSource = dtData
                lstGPLM.DataTextField = "Name"
                lstGPLM.DataValueField = "ID"
                lstGPLM.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getODM()
        Try
            Dim dtData As New DataTable


            Select Case rdNotebookDesktop.SelectedValue
                Case OPT_NOTEBOOK
                    dtData = HPQ.Excalibur.Service.ListOsspPartners()
                Case OPT_DESKTOP
                    'dtData = HPQ.Excalibur.Service.ListOsspPartners()
            End Select

            If dtData.Rows.Count > 0 Then
                lstODM.DataSource = dtData
                lstODM.DataTextField = "Name"
                lstODM.DataValueField = "ID"
                lstODM.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSPDM()
        Try
            Dim dtData As New DataTable
            Dim objData As New HPQ.Excalibur.Data

            dtData = objData.ListSpdms()

            If dtData.Rows.Count > 0 Then
                lstSpdm.DataSource = dtData
                lstSpdm.DataTextField = "Name"
                lstSpdm.DataValueField = "ID"
                lstSpdm.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getPSM()
        Try
            Dim dtData As New DataTable
            Dim objData As New HPQ.Excalibur.Data

            dtData = objData.ListSvcManagers()

            If dtData.Rows.Count > 0 Then
                lstPsm.DataSource = dtData
                lstPsm.DataTextField = "Name"
                lstPsm.DataValueField = "ID"
                lstPsm.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Export(ByVal DataGridView As GridView)
        Try
            Dim FileName As String = String.Empty
            Dim FileNameNumber As Integer = 0
            FileNameNumber = Session("FILENAMENUMBER_EXCEL")
            If FileNameNumber > 0 Then
                FileName = "PlatformAssignment" + FileNameNumber.ToString + ".xls"
            Else
                FileName = "PlatformAssignment.xls"
            End If
            FileNameNumber = FileNameNumber + 1
            Session("FILENAMENUMBER_EXCEL") = FileNameNumber
            ExportTo(DataGridView, FileName, "application/excel")
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

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)

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