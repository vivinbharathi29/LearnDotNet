<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim rs, dw, cn, cmd
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
        Case "UpdateMarketingMissingAvData"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateMarketingAvDetailViaTodayPage")
            dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request.Form("AVDetailID")
            dw.CreateParameter cmd, "@p_Last_Upd_User", adVarChar, adParamInput, 50, Request.Form("CurrentUser")
            dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 10, Request.Form("DiscontinueDate")
            dw.CreateParameter cmd, "@p_AvailabilityDt", adDate, adParamInput, 10, Request.Form("RTPDate")
            dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request.Form("BID")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            Response.Write("Added Record")
        Case Else
            Response.Write("No Function Called")
    End Select
End If
%>