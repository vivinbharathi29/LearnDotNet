<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="includes/DataWrapper.asp" -->
<!-- #include file="includes/emailwrapper.asp" -->
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
        Case "AddRemoveDeliverable"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateInitialOfferingStatus")
            dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request.Form("DRID")
            dw.CreateParameter cmd, "@p_Selected", adInteger, adParamInput, 8, Request.Form("Selected")
            dw.CreateParameter cmd, "@p_BusinessID", adInteger, adParamInput, 8, Request.Form("BusinessID")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Sucess") 
        Case "AddRemoveProduct"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateInitialOfferingDelRoot")
            dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request.Form("DRID")
            dw.CreateParameter cmd, "@p_PVID", adInteger, adParamInput, 8, Request.Form("PVID")
            dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request.Form("BID")
            dw.CreateParameter cmd, "@p_Selected", adInteger, adParamInput, 8, Request.Form("Selected")         
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Success")
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>