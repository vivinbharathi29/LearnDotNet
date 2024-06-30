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
        Case "UpdateDCRNo"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateAvDetail_ProductBrand_DCRNo")
            dw.CreateParameter cmd, "@p_DCRNo", adInteger, adParamInput, 8, Request.Form("DCRNo")
            dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request.Form("AvDetailID")
            dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request.Form("ProductBrandID")
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