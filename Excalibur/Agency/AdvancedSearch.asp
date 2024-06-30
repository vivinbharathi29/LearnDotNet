<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "AgencyMatrix.asp" --> 
<%
Dim m_SelectedProduct
Dim m_SelectedDeliverable
Dim m_SelectedRegion
Dim m_SelectedCountry

m_SelectedProduct = Request.Form("selProduct")
m_SelectedDeliverable = Request.Form("selDeliverable")
m_SelectedRegion = Request.Form("selRegion")
m_SelectedCountry = Request.Form("selCountry")

Sub FillProductList()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spGetProducts")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(m_SelectedProduct) = trim(rs("ID")) then
			Response.Write "<option selected value=""" & rs("ID") & """>" & server.HTMLEncode(rs("Name")) & " " & server.HTMLEncode(rs("Version")) & "</option>"							else
			Response.Write "<option value=""" & rs("ID") & """>" & server.HTMLEncode(rs("Name")) & " " & server.HTMLEncode(rs("Version")) & "</option>"							end if		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

Sub FillDeliverableList()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListAgencyDeliverables")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(m_SelectedDeliverable) = trim(rs("ID")) then
			Response.Write "<option selected value=""" & rs("ID") & """>" & server.HTMLEncode(rs("name")) & "</option>"							else
			Response.Write "<option value=""" & rs("ID") & """>" & server.HTMLEncode(rs("name")) & "</option>"							end if		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

Sub FillCountryList
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListCountries")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(m_SelectedCountry) = trim(rs("Country_ID")) then
			Response.Write "<option selected value=""" & rs("Country_ID") & """>" & server.HTMLEncode(rs("Country_Name")) & "</option>"							else
			Response.Write "<option value=""" & rs("Country_ID") & """>" & server.HTMLEncode(rs("Country_Name")) & "</option>"							end if		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

Sub FillRegionList()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListRegions")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(m_SelectedRegion) = trim(rs("region")) then
			Response.Write "<option selected value=""" & server.HTMLEncode(rs("region")) & """>" & server.HTMLEncode(rs("display_text")) & "</option>"							else
			Response.Write "<option value=""" & server.HTMLEncode(rs("region")) & """>" & server.HTMLEncode(rs("display_text")) & "</option>"							end if		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY onload="document.all['lblLoading'].style.display='none';">
<H2>Agency Advanced Search</H2>
<form id=frmMain name=frmMain action=AdvancedSearch.asp method=post>
<table method=post>
	<tr>
		<td><STRONG>Product:</STRONG></td>
		<td rowspan=4>&nbsp;</td>
		<td><STRONG>Country:</STRONG></td></tr>
	<tr>
		<td><SELECT id=selProduct name=selProduct style="WIDTH: 350px">
			<OPTION value="">-- Please Make A Selection --</OPTION>
			<% Call FillProductList %>
			</SELECT></td>
		<td><SELECT id=selCountry name=selCountry style="WIDTH: 350px">
			<OPTION value="">-- Please Make A Selection --</OPTION>
			<% Call FillCountryList %>
			</SELECT></td></tr>
	<tr>
		<td><STRONG>Deliverable:</STRONG></td>
		<td><STRONG>Region:</STRONG></td></tr>
	<tr>
		<td><SELECT id=selDeliverable name=selDeliverable style="WIDTH: 350px">
			<OPTION value="">-- Please Make A Selection --</OPTION>
			<% Call FillDeliverableList %>
			</SELECT></td>
		<td><SELECT id=selRegion name=selRegion style="WIDTH: 350px">
			<OPTION value="">-- Please Make A Selection --</OPTION>
			<% Call FillRegionList %>
			</SELECT></td></tr>
	<tr>
		<td colspan=3 align=right><INPUT type="Submit" value="Search" id=btnSearch name=btnSearch></td></tr>
</table>
</form>
<HR>
<%= m_SelectedRegion%>
<span id=lblLoading name=lblLoading><font size=3 color=red face=verdana>Please Wait, Loading your Report ...</font></span>
<%
If Len(Trim(m_SelectedProduct)) <> 0 OR Len(Trim(m_SelectedDeliverable)) <> 0 Then
	Response.Write "<Table ID=TableAgency width=100% border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>"
	If Len(Trim(m_SelectedProduct)) > 0 Then
		Call DrawPMViewMatrix(m_SelectedProduct, m_SelectedDeliverable, m_SelectedRegion, m_SelectedCountry)
	Else
		Call DrawDMViewMatrix(m_SelectedDeliverable, m_SelectedProduct, m_SelectedRegion, m_SelectedCountry)
	End If
	Response.Write "</table><p>* Country added after POR by DCR</p>"
End If 
%>
</BODY>
</HTML>
