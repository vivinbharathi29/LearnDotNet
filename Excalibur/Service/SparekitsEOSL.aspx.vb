Imports System.Data
Imports System.IO


Partial Class Service_SparekitsEOSL
    Inherits System.Web.UI.Page

    Private Const EXPORT_EXCEL As String = "1"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                pnlNoData.Visible = False
                pnlData.Visible = False

                GetSparekitsMaxEOSDate()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReport.Click
        Try
            GetSparekitsMaxEOSDate()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetSparekitsMaxEOSDate()
        Try
            Dim dtData As New DataTable
            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False

            dtData = HPQ.Excalibur.Service.GetSparekitsMaxEOSDate()

            If dtData.Rows.Count > 0 Then

                Session("dtData") = dtData
                If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                    gvSparekitsEOSL.AllowPaging = False
                    gvSparekitsEOSL.AllowSorting = False
                End If

                gvSparekitsEOSL.DataSource = dtData
                gvSparekitsEOSL.DataBind()

                If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                    Export(gvSparekitsEOSL)
                End If

            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "No data."
            End If

            lblLastRunDate.Text = Date.Now.ToLongDateString()

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
                FileName = "SparekitsEOSL" + FileNameNumber.ToString + ".xls"
            Else
                FileName = "SparekitsEOSL.xls"
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



    Protected Sub gvSparekitsEOSL_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvSparekitsEOSL.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvSparekitsEOSL.PageIndex = e.NewPageIndex
            gvSparekitsEOSL.SelectedIndex = -1
            gvSparekitsEOSL.DataSource = dtData
            gvSparekitsEOSL.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Protected Sub gvSparekitsEOSL_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvSparekitsEOSL.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvSparekitsEOSL.DataSource = dtData
                gvSparekitsEOSL.DataBind()
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
