<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../../includes/no-cache.asp" -->
<!-- #include file = "../../includes/noaccess.inc" -->
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file = "../../includes/Security.asp" --> 
<%
Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "usp_TerminateDCRWorkflow")
cmd.NamedParameters = True
dw.CreateParameter cmd, "@p_DCRID", adInteger, adParamInput, 8, Request("DCRID")
dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, Request("UserID")
cmd.Execute

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY onload="window.close();">
<P>&nbsp;</P>
</BODY>
</HTML>
