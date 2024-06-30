<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/emailwrapper.asp" -->
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim functionCalled
functionCalled = Request.Form("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "UpdateActionItem"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateInitialOfferingActionItem")
            dw.CreateParameter cmd, "@p_ActionItemID", adInteger, adParamInput, 8, Request.Form("ID")
            dw.CreateParameter cmd, "@p_PDMUser", adInteger, adParamInput, 8, Request.Form("PDMUser")
            dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Added Record")
        Case "UpdateActionItemPC"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateInitialOfferingActionItemPC")
            dw.CreateParameter cmd, "@p_ActionItemID", adInteger, adParamInput, 8, Request.Form("ID")
            dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Added Record") 
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>