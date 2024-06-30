<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file = "../includes/DataWrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
<!--
H2
{
	font-family: tahoma, verdana;
	font-size: small;
	color: red;
}
H3
{
	font-family: tahoma, verdana;
}
THEAD TH
{
	font-family: verdana;
	font-size: xx-small;
	background-color: wheat;
	border-top: 1px solid dimgray;
	border-bottom: 1px solid dimgray;
	text-align: center;
}
TBODY TD
{
	border-bottom: 1px solid dimgray;
	text-align: center;
	vertical-align: text-top;
}
.LeftAlign
{
	text-align: left;
}
.Region
{
	text-align: left;
	background: MediumAquamarine;
	font-weight: bold;
}
//-->
</STYLE>
<LINK href="../style/Excalibur.css" type="text/css" rel="stylesheet" >
</HEAD>

<BODY bgcolor=White >

<%
	dim strTitleColor
	strTitleColor = "#0000cd"
	on error resume next
	strTitleColor = Request.Cookies("TitleColor")
	if strTitleColor = "" then
		strTitleColor = "#0000cd"
	end if
	on error goto 0
	
	dim cn 
	dim rs
	dim dw
	dim cm
	dim strLastRegion
	dim strLastCountry
	dim strRegion
	dim strRegionID
	dim strLocalizations
	dim strExport
	dim strRow
	dim field
	dim strGroup
	
	If Trim(Request("cboProdBrand")) = "" Then
		Response.Write "<h2>No Product Brands Provided for Comparison</h2>"
		Response.End
	End If
	
	set cn = server.CreateObject("ADODB.Connection")
	Set dw = New DataWrapper
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	Select Case Request("ReportNo")
		Case 1
			Set cm = dw.CreateCommandSP(cn, "rpt_ProdBrandCountryDiff")
			Response.Write "<h3>Country Delta Report</h3>"
		Case 2
			Set cm = dw.CreateCommandSP(cn, "rpt_ProdBrandLocalizationDiff")
			Response.Write "<h3>Localization Assignment Delta Report</h3>"
		Case 3
			Set cm = dw.CreateCommandSP(cn, "rpt_ProdBrandLocalizationSettingsDiff")
			Response.Write "<h3>Localization Definition Delta Report </h3>"
		Case Else
			Response.Write "<h2>Invalid Report Number</h2>"
			Response.End
	End Select
	
	dw.CreateParameter cm, "@p_ProdBrandIdList", adVarchar, adParamInput, 1000, Request("cboProdBrand")
	Set rs = dw.ExecuteCommandReturnRS(cm)

	If rs.EOF And rs.BOF Then
		Response.Write "<h2>Error Retrieving Data</h2>"
		Response.End
	End If
	Response.Write "<Table bgcolor=ivory width=100% border=0 cellpadding=1 cellspacing=0>"

	Response.Write "<THEAD><TR>"
	
	For Each field In rs.Fields
		If field.name = rs.Fields(0).Name Then
			Response.Write "<TH class='LeftAlign'>"
		Else
			Response.Write "<TH>"
		End If

		Response.Write field.value & "</TH>"
	Next
	
	Response.Write "</TR></THEAD><TBODY>"

	Set rs = rs.NextRecordset
	
	If rs.EOF And rs.BOF Then
		Response.Write "<TR><TD ColSpan='" & rs.Fields.Count & "'><font size=3><strong>The Brands are Itentical</strong></font></TD></TR>"
	Else
		Do Until rs.EOF

			If ucase(rs.fields(0).name) = ucase("Group") And strGroup <> rs.Fields(0).Value Then
				strGroup = rs.Fields(0).Value
				Response.Write "<TR><TD Class=region ColSpan=" & rs.Fields.Count - 1 & ">" & mid(rs.Fields(0).Value,2,len(rs.Fields(0).Value)-2) & "</TD></TR>"
			End IF

			Response.Write "<TR>"
			
			For Each field In rs.Fields
						
				If ucase(rs.fields(0).name) = ucase("Group") Then
					If field.name <> rs.Fields(0).Name Then
						If field.name = rs.Fields(1).Name Then
							Response.Write "<TD class='LeftAlign'>"
						Else
							Response.Write "<TD>"
						End If
					End If
				Else
					If field.name = rs.Fields(0).Name Then
						Response.Write "<TD class='LeftAlign'>"
					Else
						Response.Write "<TD>"
					End If
				End If
				Response.Write field.value & "</TD>"
			Next

			Response.Write "</TR>"
			
			rs.MoveNext
			
		Loop
	End If
		
	Response.Write "</TBODY></TABLE>"

	rs.Close
	set rs = nothing
	set cm = nothing
	set cn = nothing
	set dw = nothing
	
  %>
</BODY>
</HTML>

