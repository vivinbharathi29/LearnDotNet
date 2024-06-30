<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="includes/DataWrapper.asp" -->
<!-- #include file="includes/emailwrapper.asp" -->
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
        Case "UpdateMarketingName"
        Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateAddEditMarketingName")
        dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request("BID")
        dw.CreateParameter cmd, "@p_Name", adVarChar, adParamInput, 100, Request("Name")
	    dw.CreateParameter cmd, "@p_NameType", adInteger, adParamInput, 8, Request("NameType")
        dw.CreateParameter cmd, "@p_PBID", adInteger, adParamInput, 8, Request("PBID")
        dw.CreateParameter cmd, "@p_Series", adVarChar, adParamInput, 50, Request("Series")
        dw.CreateParameter cmd, "@p_PlatFormID", adInteger, adParamInput, 8, Request("PlatFormID")
        Set rs = dw.ExecuteCommandReturnRS(cmd)
        set cmd = nothing
        set cn = nothing
        set dw = nothing
        set rs = nothing
        Response.Write("Success")
     Case Else
        Response.Write("No Function Called")
     End Select
End If

%>
