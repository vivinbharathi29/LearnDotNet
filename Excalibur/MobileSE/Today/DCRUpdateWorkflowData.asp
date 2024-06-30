<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/emailwrapper.asp" -->
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim rs, dw, cn, cmd, oMessage
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "AddDCRWorkflowComments"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateDCRWorkflowComments")
            dw.CreateParameter cmd, "@p_HistoryID", adInteger, adParamInput, 8, Request.Form("HistoryID")
            dw.CreateParameter cmd, "@p_Comments", adVarChar, adParamInput, 50, Request.Form("Comments")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            Response.Write("Added Record")
        Case "UpdateDCRWorkflowComplete"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateDCRWorkflowComplete")
            dw.CreateParameter cmd, "@p_HistoryID", adInteger, adParamInput, 8, Request.Form("HistoryID")
            dw.CreateParameter cmd, "@p_DCRID", adInteger, adParamInput, 8, Request.Form("DCRID")
            dw.CreateParameter cmd, "@p_PVID", adInteger, adParamInput, 8, Request.Form("PVID")
            dw.ExecuteNonQuery(cmd)
            Response.Write("Added Record")
            
            Set cmd = dw.CreateCommAndSP(cn, "usp_SelectDCRWorkflowEmailList")
            dw.CreateParameter cmd, "@DCRID", adInteger, adParamInput, 8, Request.Form("DCRID")
            dw.CreateParameter cmd, "@PVID", adInteger, adParamInput, 8, Request.Form("PVID")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            
            If rs("EmailList") <> "NA" Then              
              Set oMessage = New EmailWrapper	
		      oMessage.From = "pulsar.support@hp.com"
		      oMessage.To= rs("EmailList")
		      oMessage.Subject = "Change Request (DCR) Workflow Complete - " & rs("ProductName")
		      oMessage.HTMLBody = "<html><head><title></title></head><body><h4>The implementation workflow for DCR " & Request.Form("DCRID") & " is complete.</h4><h4>Summary: " & rs("Summary") & "</h4><a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/MobileSE/Today/Action.asp?ID=" & Request.Form("DCRID") & "&Type=3"">Click here to view DCR properties.</a></body></html>"
		      oMessage.Send 
		      Set oMessage = Nothing 	                                                                                                      
            End If            
            Response.Write(rs("EmailList"))
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>