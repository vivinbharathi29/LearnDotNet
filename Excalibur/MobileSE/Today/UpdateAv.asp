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
        Case "UpdateAv"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateDeliverableNameChangeAvLinked")
            dw.CreateParameter cmd, "@p_UpdateID", adInteger, adParamInput, 8, Request.Form("UpdateID")
            dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
            dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request.Form("UserID")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Added Record")
        Case "UpdateDeletedAvMappedToSPS"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateServiceSpareKitMapAV")
            dw.CreateParameter cmd, "@p_AvDetailIds",adVarChar, adParamInput, 150, Request.Form("AvDetailID")
            dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request.Form("UserID")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Deleted Av Mapped To SPS- Status Changed To Deleted")
        Case "UpdateAvDescription"
            Set cmd = dw.CreateCommAndSP(cn, "usp_UpdateAvDescriptionDiscrepancies")
            dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request.Form("AvDetailID")
            dw.CreateParameter cmd, "@p_NotActionable", adInteger, adParamInput, 8, Request.Form("NotActionable")
            dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request.Form("UserID")
            dw.CreateParameter cmd, "@p_DescriptionType", adInteger, adParamInput, 8, Request.Form("DescriptionType")
            dw.ExecuteNonQuery(cmd)
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write("Description Updated")
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>