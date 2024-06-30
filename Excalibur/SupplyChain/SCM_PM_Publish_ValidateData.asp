<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/emailwrapper.asp" -->
<%
'added this comment to restore this file to pulsartest; please remove this comment next time some body change this file
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
        Case "ValidateData"
        Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_Publish_ValidateDates")
        dw.CreateParameter cmd, "@p_PBID", adInteger, adParamInput, 8, Request("PBID")
        Set rs = dw.ExecuteCommandReturnRS(cmd)
        message = rs("ErrorMessage")
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