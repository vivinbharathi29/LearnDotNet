<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="includes/DataWrapper.asp" -->
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
        Case "UpdateJupiterXLR8ReportProducts"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateJupiterXLR8ReportProducts")
            dw.CreateParameter cmd, "@p_SelectedProducts", adVarChar, adParamInput, 2147483647, Request.Form("Products")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            Response.Write("Added Record")
        Case Else
            Response.Write("No Function Called")
    End Select
End If
%>