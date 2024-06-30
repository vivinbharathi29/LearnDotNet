<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<%
'added this comment to restore this file to pulsartest; please remove this comment next time some body change this file
Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd, rs, message
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
    Case "RollbackSCM"
        Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_SCMRollBack")
        dw.CreateParameter cmd, "@p_intProductBrandID", adInteger, adParamInput, 8, Request.Form("PBID")
        dw.CreateParameter cmd, "@p_chrUser", adVarChar, adParamInput, 250, Request.Form("UserName")
        dw.ExecuteNonQuery(cmd)       
        set cmd = nothing
        set cn = nothing
        set dw = nothing         
     Case Else
        Response.Write("No Function Called")
     End Select
End If

%>