<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<%
    Dim dw, cn, cmd
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Dim FunctionCalled : FunctionCalled = Request.Form("Function")
    
    Select Case FunctionCalled
        Case "UpdateFeatureAction"
            Set cmd = dw.CreateCommAndSP(cn, "usp_Today_UpdateFeatureActionStatus")
            dw.CreateParameter cmd, "@p_intFeatureActionItemID", adInteger, adParamInput, 8, Request.Form("FeatureActionItemID")            
            dw.CreateParameter cmd, "@p_intActionType", adInteger, adParamInput, 8, Request.Form("ActionType")
            dw.CreateParameter cmd, "@p_chrUser", adVarChar, adParamInput, 250, Request.Form("CurrentUserName")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing                             
        Case else
            Response.Write("No Function Called")
     End Select

 %>