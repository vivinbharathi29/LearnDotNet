<%@ Language=VBScript %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file="../../includes/no-cache.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = "1";
	window.close();
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
Dim dw, cn, cmd

Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommAndSP(cn, "usp_UpdatePDMFeedback2")
dw.CreateParameter cmd, "@p_AvActionItemID", adInteger, adParamInput, 8, Request.Form("txtAvActionItemID")
dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, Request.Form("txtCurrentUser")
dw.CreateParameter cmd, "@p_PDMFeedback", adVarchar, adParamInput, 500, Request.Form("txtPDMFeedback")
dw.ExecuteCommandNonQuery(cmd)

'	set cn = server.CreateObject("ADODB.Connection")
'	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
'	cn.Open
'
'	cn.execute "usp_UpdatePDMFeedback2 " & clng(Request.Form("txtAvActionItemID")) & ",'" & Request.Form("txtCurrentUser") & "'," & HTMLEncode(Request.Form("txtPDMFeedback")&"")
'
'	cn.Close
'	set cn = nothing

%>

</BODY>
</HTML>
