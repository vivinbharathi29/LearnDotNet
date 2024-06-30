<%@ Language=VBScript %>
<% Response.ContentType = "application/vnd.ms-excel" %>
<!-- #include file="../includes/DataWrapper.asp" -->	
<!-- #include file="../includes/common.asp" -->	
<%
Sub Main()
	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	Dim sTableRow
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spListApprovedDCRs")
	dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 8, Request("ID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		sTableRow = "<TR><TD class=cell>"
		sTableRow = sTableRow & rs("ID")
		sTableRow = sTableRow & "</TD><TD class=cell>"
		sTableRow = sTableRow & rs("Status")
		sTableRow = sTableRow & "</TD><TD class=cell>"
		sTableRow = sTableRow & rs("ApprovalDate")
		sTableRow = sTableRow & "</TD><TD class=cell>"
		sTableRow = sTableRow & rs("Summary")
		sTableRow = sTableRow & "</TD><TD class=cell>"
		sTableRow = sTableRow & rs("Submitter")
		sTableRow = sTableRow & "</TD></TR>"
		
		Response.Write sTableRow
		
		rs.movenext
	Loop
End Sub
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>DCR PDD Export</TITLE>
<STYLE>
<!--
<!-- #include file="../Style/pddExport.css" -->
//-->
</STYLE>
</HEAD>
<BODY>
<H1>Approved DCRs</H1>
<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TH class=header>Number</TH>
		<TH class=header>Status</TH>
		<TH class=header>Approval Date</TH>
		<TH class=header>Description</TH>
		<TH class=header>Submitter</TH>
	</TR>
<%
	Call Main()
%>
</TABLE>
</BODY>
</HTML>
