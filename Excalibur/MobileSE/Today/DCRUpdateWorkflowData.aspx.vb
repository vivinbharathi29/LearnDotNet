Imports System.Data

Partial Class DCRUpdateWorkflowData
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                Dim functionCalled As String
                Dim HistoryID As String
                Dim DCRID As String
                Dim PVID As String

                functionCalled = Request.Form("Function")
                HistoryID = Request.Form("HistoryID")
                DCRID = Request.Form("DCRID")
                PVID = Request.Form("PVID")

                If functionCalled = "" Then
                    Response.Write("No Function Called")
                    Response.End()
                Else
                    Select Case functionCalled
                        Case "AddDCRWorkflowComments"
                        Case "UpdateDCRWorkflowComplete"
                            dw.UpdateDCRWorkflowComplete(HistoryID, DCRID, PVID)

                            Dim dtEmail As New DataTable
                            Dim row As DataRow

                            dtEmail = dw.SelectDCRWorkflowEmailList(DCRID, PVID)
                            'Dim ew As New EmailWrapper.SendEmail
                            Dim ew As New EmailQueue
                            For Each row In dtEmail.Rows
                                If row("EmailList") <> "NA" Then
                                    ew.AddFrom = "pulsar.support@hp.com"
                                    ew.AddTo = row("EmailList")
                                    ew.Subject = "Change Request (DCR) Workflow Complete - " & row("ProductName")
                                    ew.HtmlBody = "<html><head><title></title></head><body><h4>The implementation workflow for DCR " & DCRID.ToString() & " is complete.</h4><h4>Summary: " & row("Summary") & "</h4><a href=""http://" & Session("ServerName") & "/excalibur/MobileSE/Today/Action.asp?ID=" & DCRID.ToString() & "&Type=3"">Click here to view DCR properties.</a></body></html>"
                                    ew.DSNOptions = "Base64"
                                    ew.Send()
                                End If
                                Exit For
                            Next
                            Response.Write("<script language='javascript'> if (IsFromPulsarPlus()) {ClosePulsarPlusPopup();} else { window.close();}</script>")
                        Case Else
                            Response.Write("No Function Called")
                            Response.End()
                    End Select
                End If
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
            Response.End()
        End Try
    End Sub
End Class
