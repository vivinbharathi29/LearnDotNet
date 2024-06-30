<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<%
    Dim dw, cn, cmd
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Dim FunctionCalled : FunctionCalled = Request.Form("Function")
    
    Select Case FunctionCalled
        Case "UpdateAV"
            Set cmd = dw.CreateCommAndSP(cn, "usp_Image_AddObsoleteLocalizedAV")
            dw.CreateParameter cmd, "@p_intImageActionItemID", adInteger, adParamInput, 8, Request.Form("ImageActionItemID")
            dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, Request.Form("ProductVersionID")
            dw.CreateParameter cmd, "@p_intProductBrandID", adInteger, adParamInput, 8, Request.Form("ProductBrandID")
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