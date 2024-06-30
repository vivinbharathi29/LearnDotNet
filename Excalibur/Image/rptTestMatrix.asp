<%@ Language=VBScript %>
<% 
If request("FileType")= 1 Then 
Response.ContentType = "application/vnd.ms-excel" 
Else
%>
<!-- #include file = "../includes/no-cache.asp" --> 
<% End If %>
<html>
<head>
<TITLE>Product Localization Matrix</TITLE>
<% If Request("FileType") = "" Then %>
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
<style"type="text/css">
<!--
td
{
	border-right: LightGrey 1px solid;
}
th
{
    background: brown;
    color: white;
	vertical-align: middle;
	border-right: LightGrey 1px solid;
	border-left: LightGrey 1px solid;
	white-space:nowrap;
	height: 30px
}
.Highlight TD
{
	background-color:LightPink;
}
.Region
{
	text-align: left;
	background: MediumAquamarine;
	font-weight: bold;
}
.Export TD
{
 border: none;
 background-color: white;
}
//-->
</style>
<script type="text/javascript">
<!--
function Export(strID){
	window.open (window.location.href + "&FileType=" + strID);
}
//-->
</script>
<% Else %>
<style type="text/css">
th
{
	color: white;
	background-color: black;
	border-right: white thin solid;
	border-left: white thin solid;
	border-top: black thin solid;
	border-bottom: black thin solid;
	text-aling: left;
}
td
{
	border-right: black thin solid;
	border-left: black thin solid;
	border-top: black thin solid;
	border-bottom: black thin solid;
}
.Highlight TD
{
	background-color:yellow;
}
.Region
{
	text-align: left;
	background: silver;
	font-weight: bold;
}

</style>
<% End If %>
</head>
<body>
<%if request("FileType") = "" then%>
<table class="export" width=100% border=0><tr><td align=right><font size=1 face=verdana>Export: <a href="javascript: Export(1);">Excel Export</a></td></tr></table>
<%end if%>

<%
    

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetProductVersionName " & Request("ID"),cn

	Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " Image SKU Matrix</b></font><BR><BR>"

	rs.Close
	
	rs.Open "usp_SelectImageTestMatrix " & Request("ID"),cn
	
%>
<table cellpadding="2" cellspacing="0" style="border-collapse:collapse; behavior:url(../includes/client/sort.htc);">
<%
	If rs.EOF Then
		Response.Write "<TR bgcolor=white><TD colspan=13><font size=2><b>No Countries Defined</b></font></TD></TR>"
    Else
%>
    <thead><tr>
        <th>SKU Number</th>
        <th>Brand</th>
        <th>Software</th>
        <th>OS</th>
        <th>Type</th>
        <th>Status</th>
        <th>Comments</th>
        <th>Tier</th>
<%
        '<th>Localization</th>
        '<th>HP Code</th>
        '<th>Languages</th>
        '<th>Country Code</th>
        '<th>Keyboard</th>
        '<th>Power Cord</th>
%>        
        </tr></thead><tbody>
<%
	End If
	
	do while not rs.EOF
		Response.Write "<TD>" & rs("SKUNumber") & "</td><td nowrap>" & rs("BrandName") & "</td><td>" & rs("SWType") & "</td><td nowrap>" & rs("OSName") & "</td><td>" & rs("ImageType") & "</td><td>" & rs("ImageStatus") & "&nbsp;</td><td>" & rs("Comments") & "</td><td>"  & rs("Priority") &  "</td></tr>" '<td>"  & rs("Localization") &  "</td><td>"  & rs("OptionConfig") & "</td><td>"  & strLanguages & "</td><td>"  & rs("CountryCode") & "&nbsp;</td><td>" & rs("Keyboard") & "</td><td>" & rs("PowerCord") & "</td></TR>"
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
	set cn = nothing
%>
</table>
</body>
</html>
