<%@ Language=VBScript %>
<% 
If request("FileType")= 1 Then 
Response.ContentType = "application/vnd.ms-excel" 
Else
%>
<!-- #include file = "../includes/no-cache.asp" --> 
<% End If %>
<HTML>
<HEAD>
<TITLE>Product Localization Matrix</TITLE>
<% If Request("FileType") = "" Then %>
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
<STYLE>
<!--
TD
{
	border-right: LightGrey 1px solid;
}
TH
{
	border-right: LightGrey 1px solid;
	border-left: LightGrey 1px solid;
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
</STYLE>
<SCRIPT type="text/javascript">
<!--
function Export(strID){
	window.open (window.location.href + "&FileType=" + strID);
}

//-->
</SCRIPT>
<% Else %>
<STYLE>
TH
{
	color: white;
	background-color: black;
	border-right: white thin solid;
	border-left: white thin solid;
	border-top: black thin solid;
	border-bottom: black thin solid;
	text-aling: left;
}
TD
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

</STYLE>
<% End If %>
</HEAD>
<BODY>
<%if request("FileType") = "" then%>
<table class="export" width=100% border=0><tr><td align=right style="background-color:ivory"><font size=1 face=verdana>Export: <a href="javascript: Export(1);">PDD Export</a></td></tr></table>
<%end if%>

<%
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetProductVersionName " & Request("ID"),cn

	Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " Product Localization Matrix</b></font><BR><BR>"

	rs.Close

	rs.Open "usp_SelectProductBrandLocalizations " & Request("ID"),cn

	Dim sBrand
	sBrand = ""
    dim strHeader : strHeader = "" 
    dim strData : strData = ""

    If not(rs.EOF and rs.BOF) then

    dim iPowerCordSupported
    dim iDuckheadPowerCordSupported
    dim iDuckheadSupported

    iPowerCordSupported = rs.Fields("PowerCordSupported").Value
    iDuckheadPowerCordSupported = rs.Fields("DuckheadPowerCordSupported").Value
    iDuckheadSupported = rs.Fields("DuckheadSupported").Value

    dim iColumnCount : iColumnCount = 10

    if iPowerCordSupported = 1 then
       iColumnCount = iColumnCount + 1
    end if

    if iDuckheadPowerCordSupported = 1 then
       iColumnCount = iColumnCount + 1
    end if 
    
    if iDuckheadSupported = 1 then
       iColumnCount = iColumnCount + 1
    end if

    if request("FileType") = 1 then
        iColumnCount = iColumnCount - 1
    end if

    end if
%>
<table cellpadding=2 cellspacing=0 STYLE="border-collapse:collapse;">
<%

	If rs.EOF Then
		Response.Write "<TR><TD colspan=11><font size=2><b>No Countries Defined</b></font></TD></TR>"
	End If
	
	do while not rs.EOF
		 If sBrand <> rs("Brand") Then
            strHeader = ""
            
			sBrand = rs("Brand")
			Response.Write "<tr><td style='border:none; background-color: cornsilk' colspan=" & iColumnCount - 1 & "><font size=2><strong>" & rs("Brand") & "</strong></font></td></tr>"
            Response.Write "<tr><td style='border:none; background-color: ivory' colspan=" & iColumnCount - 1 & ">&nbsp;</td></tr>"

            strHeader = "<Th>Country Region</Th><Th>HP Blue Code</Th><Th>CPQ Red DASH</Th><Th>Languages</Th><Th>Country Code</Th><Th>Keyboard</Th><Th>KWL</Th>"
            If (iPowerCordSupported = 1) Then
                strHeader = strHeader & "<Th>Power Cord</Th>"
            End If
            If (iDuckheadPowerCordSupported = 1) Then
                strHeader = strHeader & "<Th>Duckhead Power Cord</Th>"
            End if
            If (iDuckheadSupported = 1) Then
                strHeader = strHeader & "<Th>Duckhead</Th>"
            End if
            strHeader = strHeader & "<Th>Doc Language</Th>"

            Response.Write strHeader
		End If
		
		If GeoID <> rs("GeoID") Then
			GeoID = rs("GeoID")
			Response.Write "<tr><td class=region colspan=" & iColumnCount - 1 & ">" & rs("Geo") & "</td></tr>"
		End If
		
		strLanguages =  rs("OSLanguage") & ""
		if trim(rs("OtherLanguage") & "") <> "" then
			strLanguages = strLanguages & "," & rs("OtherLanguage")
		end if
		if trim(rs("KWL") & "") ="" then
			strKWL = "&nbsp;"
		else
			strKWL = rs("KWL")
		end if		
		If Not IsNull(rs("NonStandard")) Then
			Response.Write "<TR Class=Highlight>"
		Else
			Response.Write "<TR>"
		End If
        
        strData = ""
        strData = "<TD>" & rs("CountryRegion") & "</td><td>" & rs("OptionConfig") & "</td><td>" & rs("Dash") & "</td><td>" & strLanguages & "</td><td>" & rs("CountryCode") & "</td><td>"  & rs("Keyboard") &  "</td><td>"  & strKWL &  "</td>"
        If (iPowerCordSupported = 1) Then
                strData = strData & "<td>" & rs("PowerCordGEO") & "</td>"
        End If
        If (iDuckheadPowerCordSupported = 1) Then
                strData = strData & "<td>" & rs("DuckheadPowerCordGEO") & "</td>"
        End if
        If (iDuckheadSupported = 1) Then
                strData = strData & "<td>" & rs("DuckheadGEO") & "</td>"
        End if
        strData = strData & "<td>"  & rs("DocKits") & "</td></tr>" 
        
        Response.Write strData

    rs.MoveNext
	loop
	rs.Close
	set rs = nothing
	set cn = nothing
%>
</table>
</BODY>
</HTML>
