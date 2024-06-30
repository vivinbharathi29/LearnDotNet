Imports System.Data

Partial Class DCRWorkflowStatus
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    'Dim de As HPQ.Excalibur.EmailMessage = New HPQ.Excalibur.EmailMessage

    Public ReadOnly Property DCRID() As String
        Get
            Return Request("DCRID")
        End Get
    End Property

    Public ReadOnly Property HistoryID() As String
        Get
            Return Request("HistoryID")
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property CompleteMilestone() As String
        Get
            Return Request("CompleteMilestone")
        End Get
    End Property

    Public ReadOnly Property ApplicationRoot() As String
        Get
            Return Request.Path
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                If Trim(CompleteMilestone) = "" Then
                    btnOK.Visible = True
                    btnSubmit.Visible = False
                    btnCancel.Visible = False
                    txtComments.Visible = False
                    lblAddComments.Visible = False
                    lblHeader.Text = "Change Request Workflow Status"
                    DisplayTerminateButton()
                Else
                    btnOK.Visible = False
                    btnSubmit.Visible = True
                    btnCancel.Visible = True
                    txtComments.Visible = True
                    lblAddComments.Visible = True
                    lblHeader.Text = "Change Request Workflow - Complete Milestone"
                End If
                Dim dt As New DataTable
                dt = dw.SelectDCRWorkflowStatus(DCRID)

                Dim row1 As DataRow
                For Each row1 In dt.Rows
                    lblCreatedByText.Text = row1("CreatedBy")
                    lblCreateDateText.Text = row1("CreateDate")
                    lblWorkflowTypeText.Text = row1("Name")
                    Exit For
                Next

                vDCRID.Value = DCRID
                vUserID.Value = UserID
                vApplicationRoot.Value = ApplicationRoot

                gvWorkflowStatus.DataSource = dt
                gvWorkflowStatus.DataBind()
            End If
        Catch ex As Exception
            lblHeader.Text = ex.ToString
        End Try
    End Sub

    Private Sub DisplayTerminateButton()
        Try
            If UserID = "5016" Or UserID = "8" Then
                btnTerminate.Visible = True
                Exit Sub
            End If

            Dim dt As DataTable = dw.SelectSuperUsersByProduct(PVID)
            For Each row As DataRow In dt.Rows
                If (row("SMID") = UserID) Or (row("PMID") = UserID) Then
                    btnTerminate.Visible = True
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Sub

    'Protected Sub gvWorkflowStatus_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvWorkflowStatus.DataBinding
    '    Try
    '        Dim i As Integer = 0
    '        Dim row As GridViewRow
    '        For Each row In gvWorkflowStatus.Rows
    '            Dim sCompleteDate As String = IIf(Not gvWorkflowStatus.DataKeys(i).Item("MilestoneCompleteDate") Is DBNull.Value, gvWorkflowStatus.DataKeys(i).Item("MilestoneCompleteDate"), "Pending")
    '            Dim sComments As String = IIf(Not gvWorkflowStatus.DataKeys(i).Item("Comments") Is DBNull.Value, gvWorkflowStatus.DataKeys(i).Item("Comments"), "")
    '        Next
    '    Catch ex As Exception
    '        Response.Write(ex.InnerException)
    '    End Try
    'End Sub

    Protected Sub ClosePopup(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click, btnCancel.Click
        Try
            'Response.Write("<script language='javascript'> { window.close();}</script>")
            Dim scriptKey As String = "UniqueKeyForThisScript"
            Dim javaScript As String = "<script language='javascript'> if (IsFromPulsarPlus()) {ClosePulsarPlusPopup();} else { window.close();}</script>"
            ClientScript.RegisterStartupScript(Me.GetType(), scriptKey, javaScript)
        Catch ex As Exception
            Response.Write(ex.InnerException)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            dw.UpdateDCRWorkflowComments(HistoryID, txtComments.Text)
            dw.UpdateDCRWorkflowComplete(HistoryID, DCRID, PVID)

            Dim dtEmail As New DataTable
            dtEmail = dw.SelectDCRWorkflowEmailList(DCRID, PVID)

            'Dim ew As New EmailWrapper.SendEmail
            Dim ew As New EmailQueue
            Dim row As DataRow
            For Each row In dtEmail.Rows
                If row("EmailList") <> "NA" Then
                    'de.AddTo = row("EmailList")
                    'de.AddFrom = "pulsar.support@hp.com"
                    'de.Subject = "Change Request (DCR) Workflow Complete - " & row("ProductName")
                    'de.Body = "<html><head><title></title></head><body><h4>The implementation workflow for DCR " & DCRID.ToString() & " is complete.</h4><h4>Summary: " & row("Summary") & "</h4><a href=""http://" & Session("ServerName") & "/excalibur/MobileSE/Today/Action.asp?ID=" & DCRID.ToString() & "&Type=3"">Click here to view DCR properties.</a></body></html>"
                    'de.IsBodyHtml = True
                    'de.Send()

                    ew.AddFrom = "pulsar.support@hp.com"
                    ew.AddTo = row("EmailList")
                    ew.Subject = "Change Request (DCR) Workflow Complete - " & row("ProductName")
                    ew.HtmlBody = "<html><head><title></title></head><body><h4>The implementation workflow for DCR " & DCRID.ToString() & " is complete.</h4><h4>Summary: " & row("Summary") & "</h4><a href=""http://" & Session("ServerName") & "/excalibur/MobileSE/Today/Action.asp?ID=" & DCRID.ToString() & "&Type=3"">Click here to view DCR properties.</a></body></html>"
                    ew.DSNOptions = "Base64"
                    ew.Send()
                End If
                Exit For
            Next
            'Response.Write("<script language='javascript'> { window.close();}</script>")
            Dim scriptKey As String = "UniqueKeyForThisScript"
            Dim javaScript As String = "<script language='javascript'> if (IsFromPulsarPlus()) {window.parent.parent.parent.popupCallBack(1); ClosePulsarPlusPopup();} else { window.close();}</script>"
            ClientScript.RegisterStartupScript(Me.GetType(), scriptKey, javaScript)
        Catch ex As Exception
            Response.Write(ex.InnerException)
        End Try
    End Sub

End Class
