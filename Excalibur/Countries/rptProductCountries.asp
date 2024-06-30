<%@ Language=VBScript %>
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file="../includes/DataWrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<BODY bgcolor=White>
<LINK href="../style/Excalibur.css" type="text/css" rel="stylesheet" >

<%
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
	dim strLastRegion
	dim strLastCountry
	dim strRegion
	dim strRegionID
	dim strLocalizations
	dim strExport
	dim strRow

	set cn = server.CreateObject("ADODB.Connection")
	Set dw = New DataWrapper
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	Set cm = dw.CreateCommandSP(cn, "spGetProductVersionName")
	dw.CreateParameter cm, "@ID", adInteger, adParamInput, 8, Request("ID")
	Set rs = dw.ExecuteCommandReturnRS(cm)

	Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " Product Country List</b></font><BR><BR>"

	rs.Close
	
	
	Set cm = dw.CreateCommandSP(cn, "usp_SelectProductBrandCountriesWithInactiveCountries")
	dw.CreateParameter cm, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("ID")
	Set rs = dw.ExecuteCommandReturnRS(cm)
	
'	rs.Open "spListCountriesWithLocalizations " & clng(request("OptionID")),cn,adOpenForwardOnly
	Response.Write "<Table bgcolor=ivory width=100% border=0 cellpadding=1 cellspacing=0>"
	strLastRegion = ""
	strCountry = ""
	strBrand = ""
	
	If rs.EOF Then
		Response.Write "<TR bgcolor=white><TD><font size=2><b>No Countries Defined</b></font></TD></TR>"
	End If

	do while not rs.EOF 
		if (request("List") = "" or request("List") = "All") or (request("List") = "Tablet" and rs("Tablet"))or (request("List") = "Commercial" and rs("Commercial"))or (request("List") = "Consumer" and rs("Consumer")) then
			if LastCountry <> rs("Country") then
				if LastCountry <> "" then
					If Trim(mid(strLocalizations,2)) = "" Then 
						strLocalizations = "&nbsp;"
					Else
						strLocalizations = mid(strLocalizations,2)
					End If
					response.write "<TR><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & LastCountry & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strExport & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strLocalizations & "</font></TD></TR>"
					strLocalizations=""
				end if
				LastCountry = rs("Country")
			end if
			If strBrand <> rs("Name") Then
				strBrand = rs("Name")
				Response.Write "<tr bgcolor = white><td colspan=3>&nbsp;</td></tr>"
				Response.Write "<tr bgcolor = white><td colspan=3><font size=2><strong>" & rs("Name") & "</strong></font></td></tr>"
				Response.Write "<TR bgcolor=beige><TD><font size=1 face=verdana><b>Country</b></font></TD><TD><font size=1 face=verdana><b>Export</b></font></TD><TD><font size=1 face=verdana><b>Localizations</b></font></TD></tr>"
			End If
			if strLastRegion <> rs("Region") & "" then
				strRegion = mid(left(rs("Region"),len(rs("Region"))-1),2)
				Response.Write "<TR bgcolor=MediumAquamarine><TD colspan=3 style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana><b>" & strRegion & "</b></font></TD></TR>"
				strLastRegion = rs("Region") & ""
			end if
			if rs("Export") then
				strExport = "Yes"
			else
				strExport = "No"
			end if
			strLocalizations = strLocalizations & ", " & rs("OptionConfig") & ""
		end if
		rs.MoveNext
	loop
	response.write "<TR><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & LastCountry & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strExport & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & mid(strLocalizations,2) & "</font></TD></TR>"
	rs.Close
	Response.Write "</table>"
	  
	set rs = nothing
	cn.Close
	set cn = nothing

  %>
</BODY>
</HTML>

