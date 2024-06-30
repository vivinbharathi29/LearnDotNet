<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>Localization Countries Where Used</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
<STYLE>
<!--
BODY
{
    font-family: Verdana;
    font-size:xx-small;
    background-color: White;
}
TD
{
	border-right: LightGrey 1px solid;
    font-family: Verdana;
    font-size:xx-small;
}
TH
{
	border-right: LightGrey 1px solid;
	border-left: LightGrey 1px solid;
	white-space: nowrap;
    font-family: Verdana;
    font-size:xx-small;
    background-color:Beige;
    
}
.Highlight TD
{
	background-color:LightPink;
}
//-->
</STYLE>
</HEAD>
<BODY>

<%

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetProductVersionName " & Request("ID"),cn

	Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " Product Countries Used In Localizations</b></font><BR><BR>"

	rs.Close

	rs.Open "usp_SelectProdBrandLocalizationCountries " & Request("ID"),cn,adOpenForwardOnly
	
	Response.Write "<table cellpadding=2 cellspacing=0 STYLE='border-collapse:collapse;'>"
	
	Dim sBrand
	sBrand = ""
	
	If rs.EOF Then
		Response.Write "</table><font size=2><b>No Countries Defined</b></font>"
	End If

	do while not rs.EOF
		If sBrand <> rs("Brand") Then
			sBrand = rs("Brand")
			Response.Write "<tr><td style='border:none; background-color:white' colspan=11>&nbsp;</td></tr>"
			Response.Write "<tr><td style='border:none; background-color:white' bgcolor=white colspan=11><font size=2><strong>" & rs("Brand") & "</strong></font></td></tr>"
			Response.Write "<TR><Th>Config</Th><Th>Dash</Th><Th>Country List</Th></Tr>"
		End If
		strCountries =  rs("CountryList") & ""
		Response.Write "<tr><td>" & rs("OptionConfig") & "</td><td>" & rs("Dash") & "</td><td>" & strCountries & "</td></tr>"

		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
	set cn = nothing
%>
</table>
</BODY>
</HTML>
