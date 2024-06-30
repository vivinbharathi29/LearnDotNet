<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/emailwrapper.asp" -->
<%

Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd, rs, message
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
Set rs = Server.CreateObject("ADODB.RecordSet")
Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "ValidateConfigRule"
        Set cmd = dw.CreateCommAndSP(cn, "usp_SelectValidateSharedAVConfigRule")
        dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request("AVID")
        dw.CreateParameter cmd, "@p_PBID", adInteger, adParamInput, 8, Request("BID")
        dw.CreateParameter cmd, "@p_ConfigRules", adVarChar, adParamInput, 2000, Request("ConfigRules")
        Set rs = dw.ExecuteCommandReturnRS(cmd)
        message = rs("PassFail")
        rs.Close
        set cmd = nothing
        set cn = nothing
        set dw = nothing
        set rs = nothing
        Response.Write(message)
     Case Else
        Response.Write("No Function Called")
     End Select
End If
%>