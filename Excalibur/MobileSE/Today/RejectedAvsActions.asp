<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<%
    Dim dw, cn, cmd
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    
    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_RejectedAVs_ActionItemComplete")
    dw.CreateParameter cmd, "@p_AVPHwebRejectionItemID", adInteger, adParamInput, 8, Request.Form("AVPHwebRejectionItemID")
    dw.CreateParameter cmd, "@p_chrUser", adVarChar, adParamInput, 250, Request.Form("CurrentUserName")
    dw.ExecuteNonQuery(cmd)
    set cmd = nothing
    set cn = nothing
    set dw = nothing

 %>