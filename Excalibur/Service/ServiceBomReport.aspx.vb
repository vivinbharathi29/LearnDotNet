Imports System.Data

Partial Class Service_ServiceBomReport
    Inherits System.Web.UI.Page

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

    Private Const EXPORT_EXCEL As String = "1"
    Private Const EXPORT_WORD As String = "2"

    Const FIELD_NAME_DATAGRID_ORDER As String = ""


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
                Session("OrderField") = FIELD_NAME_DATAGRID_ORDER

                If Not PreviousPage Is Nothing Then
                    Dim txtSKUNumber As TextBox = PreviousPage.FindControl("txtSKUNumber")
                    Dim txtKMAT As TextBox = PreviousPage.FindControl("txtKmat")
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

                    Dim sSQlSKUNumber As New StringBuilder
                    If txtSKUNumber.Text <> String.Empty Then
                        Dim sSku As Array = txtSKUNumber.Text.Split(",")

                        For Each sCad As String In sSku
                            If sSku.Length = 1 Then
                                sSQlSKUNumber.Append(sCad.ToString)
                            Else
                                If sSQlSKUNumber.ToString = String.Empty Then
                                    sSQlSKUNumber.Append(sCad.ToString)
                                Else
                                    sSQlSKUNumber.Append(",")
                                    sSQlSKUNumber.Append(sCad.ToString)
                                End If
                            End If
                        Next
                    End If

                    Dim sSQlKMAT As New StringBuilder
                    If txtKMAT.Text <> String.Empty Then
                        Dim sKMAT As Array = txtKMAT.Text.Split(",")

                        For Each sCad As String In sKMAT
                            If sKMAT.Length = 1 Then
                                sSQlKMAT.Append(sCad.ToString)
                            Else
                                If sSQlKMAT.ToString = String.Empty Then
                                    sSQlKMAT.Append(sCad.ToString)
                                Else
                                    sSQlKMAT.Append(",")
                                    sSQlKMAT.Append(sCad.ToString)
                                End If
                            End If
                        Next
                    End If


                    SKUNumber = sSQlSKUNumber.ToString.Trim
                    KMAT = sSQlKMAT.ToString.Trim
                    If txtMaxRows.Text = "" Then MaxRows = String.Empty Else MaxRows = txtMaxRows.Text

                    'Load Data
                    getServiceReportBom()

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


    Private Sub getServiceReportBom()
        Try
            Dim dtData As New DataTable


            dtData = HPQ.Excalibur.Service.getAdvancedServiceBomReport(SKUNumber, KMAT, MaxRows)

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                gvData.DataSource = dtData
                gvData.DataBind()
                Session("dtData") = dtData
            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "There are not bom Details for the filters selected."
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
