Imports System.Data
Imports System.IO
Imports System.Data.OleDb

Partial Class Service_DesktopPartnumbers
    Inherits System.Web.UI.Page

    Public ServiceFamilyPartNumner, FamilyName, PVID, ProductBrandID As String
    Private Const ROLE_UPLOAD_EOA_FILE As String = "50"
    Private Const EXPORT_EXCEL As String = "1"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            ServiceFamilyPartNumner = Request.QueryString("servicefamilypn")
            FamilyName = Request.QueryString("FamilyName")
            PVID = Request.QueryString("ProductVersionID")
            ProductBrandID = Request.QueryString("ProductBrandID")
            ProductVersionId.Value = PVID

            Dim sUser As String = Session("LoggedInUser")

            If Not Page.IsPostBack Then
                If ServiceFamilyPartNumner <> String.Empty Then
                    pnlNoData.Visible = False
                    pnlData.Visible = False
                    lblTitle.Text = lblTitle.Text + ": " + FamilyName + " - " + ServiceFamilyPartNumner
                    'txtComparisonDate.Value = GetRslPublishDates(PVID)
                End If

                If UserIsInRoleToUpload(sUser) Then
                    pnlUploadLinkAvSparekits.Visible = True

                Else
                    pnlUploadLinkAvSparekits.Visible = False
                End If

            getDesktopServicefamilyPartnumbers()

            End If

          

            lblErrorMsg.Text = String.Empty

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReport.Click
        Try
            getDesktopServicefamilyPartnumbers()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntUploadFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntUploadFile.Click
        Try
            'check to make sure a file is selected
            If Me.flServiceLinkAvSparekits.HasFile Then

                Dim Extension As String = Path.GetExtension(flServiceLinkAvSparekits.PostedFile.FileName)
                Dim tempPath As String = Path.GetTempPath
                Dim tempFileName As String = String.Format("{0}{1}", Path.GetTempFileName, Path.GetExtension(flServiceLinkAvSparekits.PostedFile.FileName))
                Dim tempFileSpec As String = Path.Combine(tempPath, tempFileName)
                Dim strUploadFileName As String = Path.GetFileName(flServiceLinkAvSparekits.PostedFile.FileName)
                'Dim FileName As String = Path.Combine(Server.MapPath("~/Files"), FileUpload1.FileName)
                'save the file to our local path
                flServiceLinkAvSparekits.SaveAs(tempFileSpec)

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
                        MyCommand = New OleDbDataAdapter("select * from [LinkAvSparekits$]", MyConnection)
                        Try
                            MyCommand.Fill(ds)
                        Catch ex As Exception
                            lblErrorMsg.Text = "The sheet name has to be LinkAvSparekits. The format file is Microsoft Excel 97-2003." & ex.Message
                            'External table is not in the expected format
                            Exit Sub
                        End Try

                    Catch ex As Exception
                        lblErrorMsg.Text = "The file must be saved as a Excel 2007 file."
                    End Try

                    'Dim sServiceLinkAvSparekits As String = String.Empty
                    Dim sSparekitNumber As String = String.Empty
                    Dim sAvNumber As String = String.Empty
                    Dim sAvCategory As String = String.Empty
                    Dim sSparekitsShareOnNotebooks As String = String.Empty

                    'update Sparekits
                    Dim sGeoNa As String
                    Dim sGeoLa As String
                    Dim sGeoEMEA As String
                    Dim sGeoAPJ As String
                    Dim sLocalStockAdviceID As String
                    Dim sWarrantyTierID As String
                    Dim sDispositionID As String
                    Dim sCsrLEvelID As String
                    Dim sSpsDescription As String
                    Dim sSupplier As String



                    For Each row As DataRow In ds.Tables(0).Rows
                        sSparekitNumber = row("SparekitNumber").ToString.Trim
                        sAvNumber = row("AVNumber").ToString.Trim
                        'sAvCategory = row("AvCategoryID").ToString.Trim

                        'update Sparekits
                        sSpsDescription = row("spskitdescription").ToString.Trim
                        sGeoNa = IIf(Convert.ToBoolean(row("GeoNa")), 1, 0).ToString
                        sGeoLa = IIf(Convert.ToBoolean(row("GeoLA")), 1, 0).ToString
                        sGeoEMEA = IIf(Convert.ToBoolean(row("GeoEmea")), 1, 0).ToString
                        sGeoAPJ = IIf(Convert.ToBoolean(row("GeoAPJ")), 1, 0).ToString
                        sLocalStockAdviceID = row("LocalStockAdviceID").ToString.Trim
                        sWarrantyTierID = row("WarrantyTierID").ToString.Trim
                        sDispositionID = row("DispositionID").ToString.Trim
                        sCsrLEvelID = row("CSRLevelID").ToString.Trim
                        sSupplier = row("Supplier").ToString.Trim

                        'update Sparekits
                        If sSparekitNumber <> String.Empty Then
                            Dim iRes As Integer = HPQ.Excalibur.Service.UpdateDesktopSparekit(ServiceFamilyPartNumner, sSparekitNumber, sSpsDescription, String.Empty, sCsrLEvelID, sDispositionID, sWarrantyTierID, sLocalStockAdviceID, sGeoNa, sGeoLa, sGeoAPJ, sGeoEMEA, sSupplier)

                            If iRes <> -1 Then
                                lblErrorMsg.Text = "Error when inserting Sparekit Number. " & sSparekitNumber
                            End If

                        End If

                        'Delete sparekits for a family.
                        DeleteServiceFamilySparekits(ServiceFamilyPartNumner)

                        'mapping 
                        If sSparekitNumber <> String.Empty Then
                            If sAvCategory <> String.Empty And sAvNumber <> String.Empty Then
                                'if an Sparekit is in an Notebook Family I can not Reset the information, we have to do it manually...
                                If bSparekitOnNotebookFamily(sSparekitNumber) = True Then
                                    'lblErrorMsg.Text = "The SparekitNumber is share in a notebook family. You have to do this manually."
                                    If sSparekitsShareOnNotebooks = String.Empty Then
                                        sSparekitsShareOnNotebooks = sSparekitNumber
                                    Else
                                        sSparekitsShareOnNotebooks = sSparekitsShareOnNotebooks + ", " + sSparekitNumber
                                    End If

                                Else
                                    UpdateLinkAvSparekits(ServiceFamilyPartNumner, sSparekitNumber, sAvCategory, sAvNumber)

                                End If
                            Else
                                lblErrorMsg.Text = "The AVNumber or de Category can not be a empty value. You have to review it."
                                Exit Sub
                            End If
                        Else
                            lblErrorMsg.Text = "The Sparekit number field can not be a empty value. You have to review it."
                            Exit Sub
                        End If
                    Next

                    If sSparekitsShareOnNotebooks <> String.Empty Then
                        lblErrorMsg.Text = "Next sparekits are shared with notebooks families. You have to manually mapped the next ones: " & sSparekitsShareOnNotebooks
                    End If
                End If

                'delete the temporal file
                If File.Exists(tempFileSpec) Then
                    File.Delete(tempFileSpec)
                End If

                getDesktopServicefamilyPartnumbers()

                lblErrorMsg.Text = "Upload completed..."


            Else
                lblErrorMsg.Text = "You have to select a file to upload..."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvDesktopPartnumbers_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvDesktopPartnumbers.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If DataBinder.Eval(e.Row.DataItem, "SparekitCreated") = "0" Then
                    e.Row.Cells(2).Text = ""
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getDesktopServicefamilyPartnumbers()
        Try
            Dim dtData As New DataTable
            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False

            dtData = HPQ.Excalibur.Service.GetServiceDesktopFamilyPartNumbers(ServiceFamilyPartNumner)

            If dtData.Rows.Count > 0 Then

                Session("dtData") = dtData
                gvDesktopPartnumbers.DataSource = dtData
                gvDesktopPartnumbers.DataBind()

                If ddlReportFormat.SelectedValue = EXPORT_EXCEL Then
                    gvDesktopPartnumbersToExport.AllowPaging = False
                    gvDesktopPartnumbersToExport.AllowSorting = False

                    gvDesktopPartnumbersToExport.DataSource = dtData
                    gvDesktopPartnumbersToExport.DataBind()

                    Export(gvDesktopPartnumbersToExport)
                End If

            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "No Partnumbers for the ServicefamilyPn Selected."
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
                FileName = "LinkAvSparekits" + FileNameNumber.ToString + ".xls"
            Else
                FileName = "LinkAvSparekits.xls"
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

    Protected Sub gvDesktopPartnumbers_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvDesktopPartnumbers.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvDesktopPartnumbers.DataSource = dtData
                gvDesktopPartnumbers.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub gvDesktopPartnumbers_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvDesktopPartnumbers.PageIndexChanging
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            gvDesktopPartnumbers.PageIndex = e.NewPageIndex
            gvDesktopPartnumbers.SelectedIndex = -1
            gvDesktopPartnumbers.DataSource = dtData
            gvDesktopPartnumbers.DataBind()
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

    Private Function UserIsInRoleToUpload(ByVal User As String) As Boolean
        Try
            Dim NTUser As String = Right(User, Len(User) - InStr(User, "\"))

            Dim dtUser As DataTable = HPQ.Excalibur.Service.GetUserIDInNTUserAndRole(NTUser, ROLE_UPLOAD_EOA_FILE)

            If dtUser.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub UpdateLinkAvSparekits(ByVal ServiceFamilyPn As String, ByVal sSparekitNumber As String, ByVal sAvCategory As String, ByVal sAvNumber As String)
        Try
            Dim iRes As Integer = HPQ.Excalibur.Service.UpdateLinkAvSparekits(ServiceFamilyPn, sSparekitNumber, sAvCategory, sAvNumber)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub DeleteServiceFamilySparekits(ByVal ServiceFamilyPartNumner As String)
        Try
            Dim iRes As Integer = HPQ.Excalibur.Service.DeleteServiceFamilySparekits(ServiceFamilyPartNumner)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function bSparekitOnNotebookFamily(ByVal sSparekitNumber As String) As Boolean
        Try
            Dim dtSparekit As DataTable = HPQ.Excalibur.Service.GetSparekit(sSparekitNumber)

            If dtSparekit.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex

        End Try
    End Function

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)

    End Sub

    'Public Function GetRslPublishDates(ByVal ProductVersionId As String) As String
    '    Try
    '        Dim dtData As New DataTable
    '        Me.pnlData.Visible = True
    '        Me.pnlNoData.Visible = False

    '        GetRslPublishDates = String.Empty

    '        dtData = HPQ.Excalibur.Service.GetRslPublishDates(ProductVersionId)

    '        If dtData.Rows.Count > 0 Then
    '            GetRslPublishDates = dtData.Rows(0)("ChangeDt").ToString

    '        End If

    '        '    returnValue = "<select class=""form"" name=""selCompareDt"" id=""selCompareDt"">"
    '        '    returnValue = returnValue & "<option selected value=""" & rs("ChangeDt") & """>" & DayOfWeek(rs("ChangeDt")) & " " & rs("ChangeDt") & "</option>"
    '        '    rs.MoveNext()
    '        '    Do Until rs.EOF
    '        '        returnValue = returnValue & "<option value=""" & rs("ChangeDt") & """>" & DayOfWeek(rs("ChangeDt")) & " " & rs("ChangeDt") & "</option>"
    '        '        rs.MoveNext()
    '        '    Loop


    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

   
   
End Class
