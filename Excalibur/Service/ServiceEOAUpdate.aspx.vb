Imports System.Data
Imports System.IO
Imports System.Data.OleDb

Partial Class Service_ServiceEOAUpdate
    Inherits System.Web.UI.Page

    Private Const EXPORT_EXCEL As String = "1"
    Private Const SERVICEEAO_UNAVAILABLE As String = "0"
    Private Const SERVICE_EAO_BLANK As String = "1"
    Private Const SERVICE_EAO_MIDDLE As String = "2"
    Private Const ROLE_UPLOAD_EOA_FILE As String = "50"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sUser As String = Session("LoggedInUser")


        If Not Page.IsPostBack Then

            UserIsInRoleToUpload(sUser)

            pnlNoData.Visible = False
            pnlData.Visible = False

            GetEOAUpdate()
            GetCategories()
            GetSupplier()

            pnlUploadEOAFile.Visible = IIf(hidEditPermission.Value = "1", True, False)

        End If


        lblErrorMsg.Text = String.Empty

    End Sub

    Protected Sub lnkAddSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAddSupplier.Click
        Try
            trNewSupplier.Visible = True
            lnkAddSupplier.Enabled = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReport.Click
        Try
            GetEOAUpdate()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try
            lstCategories.SelectedIndex = -1
            lstSuppplier.SelectedIndex = -1
            txtHPPartNumber.Text = ""
            For Each elen As ListItem In rdServiceEAO.Items
                elen.Selected = False
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntUploadFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUploadFile.Click
        Try
            'check to make sure a file is selected
            If flServiceEOA.HasFile Then

                Dim Extension As String = Path.GetExtension(flServiceEOA.PostedFile.FileName)
                Dim tempPath As String = Path.GetTempPath
                Dim tempFileName As String = String.Format("{0}{1}", Path.GetTempFileName, Path.GetExtension(flServiceEOA.PostedFile.FileName))
                Dim tempFileSpec As String = Path.Combine(tempPath, tempFileName)
                Dim strUploadFileName As String = Path.GetFileName(flServiceEOA.PostedFile.FileName)
                'Dim FileName As String = Path.Combine(Server.MapPath("~/Files"), FileUpload1.FileName)
                'save the file to our local path
                flServiceEOA.SaveAs(tempFileSpec)

                Dim connectionString As String = String.Empty

                Select Case Path.GetExtension(strUploadFileName).ToLower()
                    Case ".xls"
                        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1""", tempFileSpec)
                    Case ".xlsx"
                        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1""", tempFileSpec)
                    Case Else
                        'bodyPre.InnerText += vbCrLf & String.Format("Invalid File Extension ({0})", Path.GetExtension(uploadedFileName))
                End Select

                If connectionString <> String.Empty Then
                    Dim ds As DataSet = New System.Data.DataSet()
                    Dim MyConnection As OleDbConnection
                    Dim MyCommand As OleDbDataAdapter

                    Try
                        MyConnection = New OleDbConnection(connectionString)
                        MyCommand = New OleDbDataAdapter("select * from [ServiceEOAUpdate$]", MyConnection)
                        Try
                            MyCommand.Fill(ds)
                        Catch ex As Exception
                            lblErrorMsg.Text = "The sheet name has to be ServiceEOAUpdate. The format file is Microsoft Excel 97-2003." & ex.Message
                            'External table is not in the expected format
                            Exit Sub
                        End Try

                    Catch ex As Exception
                        lblErrorMsg.Text = "The file must be saved as a Excel 2007 file."
                    End Try



                    Dim sServiceEOA As String = String.Empty
                    Dim sExcaliburID As String = String.Empty

                    For Each row As DataRow In ds.Tables(0).Rows
                        sExcaliburID = row("ExcaliburID").ToString.Trim
                        sServiceEOA = row("Service EOA Date").ToString.Trim

                        If sExcaliburID <> String.Empty Then
                            If sServiceEOA <> "Unavailable" Then
                                If sServiceEOA = String.Empty Then
                                    UpdateServiceEOA(sExcaliburID, sServiceEOA)
                                ElseIf sServiceEOA = "TBD" Then
                                    Dim sSixMonthsLater As String = String.Empty
                                    'change it to the 15th day 6 months from now
                                    sSixMonthsLater = Now.Month + 6
                                    Select Case sSixMonthsLater
                                        Case 13
                                            sSixMonthsLater = 1
                                        Case 14
                                            sSixMonthsLater = 2
                                        Case 15
                                            sSixMonthsLater = 3
                                        Case 16
                                            sSixMonthsLater = 4
                                        Case 17
                                            sSixMonthsLater = 5
                                        Case 18
                                            sSixMonthsLater = 6
                                    End Select
                                    Dim sTBD As String = sSixMonthsLater + "/15/" + IIf(sSixMonthsLater = 1 Or sSixMonthsLater = 2 Or sSixMonthsLater = 3 Or sSixMonthsLater = 4 Or sSixMonthsLater = 5 Or sSixMonthsLater = 6, Now.Year + 1, Now.Year).ToString
                                    UpdateServiceEOA(sExcaliburID, sTBD)
                                ElseIf IsDate(sServiceEOA) Then
                                    UpdateServiceEOA(sExcaliburID, CDate(sServiceEOA))
                                Else
                                    lblErrorMsg.Text = "The Service EOA Date: " + sServiceEOA + " is not a valid date . You have to review it."
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next

                End If

                'delete the temporal file
                If File.Exists(tempFileSpec) Then
                    File.Delete(tempFileSpec)
                End If


                lblErrorMsg.Text = "Upload completed..."

            Else
                lblErrorMsg.Text = "You have to select a file to upload..."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lstCategories_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstCategories.SelectedIndexChanged
        Try
            GetSupplier()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetCategories()
        Try
            Dim dtData As New DataTable
            Dim objData As New HPQ.Excalibur.Data
            dtData = objData.ListDeliverablesCategories()


            If dtData.Rows.Count > 0 Then
                lstCategories.DataSource = dtData
                lstCategories.DataTextField = "Name"
                lstCategories.DataValueField = "ID"
                lstCategories.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetSupplier()
        Try
            Dim dtData As New DataTable

            Dim sCategories As StringBuilder = New StringBuilder()
            For Each elem As ListItem In lstCategories.Items
                If elem.Selected Then
                    sCategories.Append(elem.Value & ",")
                End If
            Next
            If sCategories.Length > 0 Then sCategories.Remove(sCategories.Length - 1, 1)


            dtData = HPQ.Excalibur.Service.GetSuppliers(sCategories.ToString)

            lstSuppplier.DataTextField = "Name"
            lstSuppplier.DataValueField = "ID"
           
            lstSuppplier.DataSource = dtData
            lstSuppplier.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetEOAUpdate()
        Try
            Dim dtData As New DataTable
            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False

            Dim sCategories As StringBuilder = New StringBuilder()
            For Each elem As ListItem In lstCategories.Items
                If elem.Selected Then
                    sCategories.Append(elem.Value & ",")
                End If
            Next
            If sCategories.Length > 0 Then sCategories.Remove(sCategories.Length - 1, 1)


            Dim sSupplier As StringBuilder = New StringBuilder()
            For Each elem As ListItem In lstSuppplier.Items
                If elem.Selected Then
                    sSupplier.Append(elem.Value & ",")
                End If
            Next
            If sSupplier.Length > 0 Then sSupplier.Remove(sSupplier.Length - 1, 1)

            dtData = HPQ.Excalibur.Service.GetServiceCommoditiesEOA(sSupplier.ToString, sCategories.ToString, txtHPPartNumber.Text, rdServiceEAO.SelectedValue)

            If dtData.Rows.Count > 0 Then
                Session("dtData") = dtData
                If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                    gvEOAUpdate.AllowPaging = False
                    gvEOAUpdate.AllowSorting = False
                End If

                gvEOAUpdate.DataSource = dtData
                gvEOAUpdate.DataBind()

                gvEOAUpdate.Columns(gvEOAUpdate.Columns.Count - 1).Visible = False
                gvEOAUpdate.Columns(gvEOAUpdate.Columns.Count - 2).Visible = False
                If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                    Export(gvEOAUpdate)
                End If
                gvEOAUpdate.Columns(gvEOAUpdate.Columns.Count - 1).Visible = IIf(hidEditPermission.Value = "1", True, False)
                gvEOAUpdate.Columns(gvEOAUpdate.Columns.Count - 2).Visible = True

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

    Private Sub UserIsInRoleToUpload(ByVal User As String)

        Dim NTUser As String
        Dim dtUser As DataTable

        NTUser = Right(User, Len(User) - InStr(User, "\"))
        dtUser = HPQ.Excalibur.Service.GetUserIDInNTUserAndRole(NTUser, ROLE_UPLOAD_EOA_FILE)
        If dtUser.Rows.Count > 0 Then
            hidEditPermission.Value = "1"
        Else
            hidEditPermission.Value = "0"
        End If


    End Sub

    Private Sub UpdateServiceEOA(ByVal sExcaliburID As String, ByVal sServiceEOA As String)
        Try
            Dim iRes As Integer = HPQ.Excalibur.Service.UpdateServiceEOADate(sExcaliburID, sServiceEOA)
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
                FileName = "ServiceEOAUpdate" + FileNameNumber.ToString + ".xls"
            Else
                FileName = "ServiceEOAUpdate.xls"
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

    Protected Sub gvEOAUpdate_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvEOAUpdate.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvEOAUpdate.PageIndex = e.NewPageIndex
            gvEOAUpdate.SelectedIndex = -1
            gvEOAUpdate.DataSource = dtData
            gvEOAUpdate.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvEOAUpdate_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvEOAUpdate.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvEOAUpdate.DataSource = dtData
                gvEOAUpdate.DataBind()
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

    Protected Sub btnAddSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddSupplier.Click
        Try
            'Insert Supplier Name in database
            If txtSupplierName.Text <> "" Then
                Dim iRes As Integer = HPQ.Excalibur.Service.InsertSupplier(txtSupplierName.Text)
            Else
                lblErrorMessage.Text = "You have to write a Supplier Name"
                Exit Sub
            End If

            'Refresh Supplier
            trNewSupplier.Visible = False
            GetSupplier()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnCancelAddSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancelAddSupplier.Click
        Try
            trNewSupplier.Visible = False
            txtSupplierName.Text = ""
            lblErrorMessage.Text = ""
            lnkAddSupplier.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
