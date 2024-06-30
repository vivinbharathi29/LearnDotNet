<%@  language="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")

'
' Setup the data connections
'
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
        Case "IgnoreAv"
            Set cmd = dw.CreateCommandSP(cn, "usp_ServiceAvIgnoreInsert")
            dw.CreateParameter cmd, "@AvNo", adVarchar, adParamInput, 15, Request.Form("AV")
            dw.CreateParameter cmd, "@KMAT", adVarchar, adParamInput, 10, Request.Form("KMAT")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            Response.Write("Added Record")
        Case Else
            Response.Write("No Function Called")
    End Select
End If
%>