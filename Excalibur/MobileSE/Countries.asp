<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function chkAll_onclick(strRegion) {
	var i;
	if (document.all("chkAll" + strRegion).checked)
		{
		for (i=0;i<document.all("chkSelected" + strRegion).length;i++)
			{
			frmCountries("chkSelected" + strRegion)(i).checked = true;
			document.all("Row" + frmCountries("chkSelected" + strRegion)(i).value).bgColor="lightsteelblue";
			}
		}
	else
		{
		for (i=0;i<document.all("chkSelected" + strRegion).length;i++)
			{
			frmCountries("chkSelected" + strRegion)(i).checked = false;
			document.all("Row" + frmCountries("chkSelected" + strRegion)(i).value).bgColor="ivory";
			}
		}
		
}

function chkMilestone_onclick(ID, RowID) {
	var strID;
	var AllChecked;
	var AllUnchecked;

	if ( event.srcElement.checked)
		document.all("Row" + ID).bgColor="lightsteelblue";				
	else
		document.all("Row" + ID).bgColor="ivory";		
		
}
function window_onload() {
		menubar.style.display = "";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
</STYLE>
<BODY bgcolor=White LANGUAGE=javascript onload="return window_onload()">
<LINK href="../style/Excalibur.css" type="text/css" rel="stylesheet" >
<LINK href="../style/wizard%20style.css" type="text/css" rel="stylesheet" >

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
	dim strLastRegion
	dim strLastCountry
	dim strRegion
	dim strRegionID
	dim strLocalizations
	dim strExport
	dim strRow

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	dim strOptionID
	if request("OptionID") = "" then
		strOptionID = ""
	else
		strOptionID = "&OptionID=" & request("OptionID")
	end if
	
	Response.Write "<font size=3 face=verdana><b>Master Country List</b></font><BR><BR>"
%>
	<Table style="Display=none" Id=menubar Class=MenuBar border=1 bordercolor=Ivory cellspacing=0 cellpadding =2 >
	<TR bgcolor=<%=strTitleColor%>>
	<%if request("List")="" or request("List")="All" then%>
		<td class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;All&nbsp;&nbsp;&nbsp</font>
	<%else%>
		<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="Countries.asp?List=All&OptionName=<%=Server.HtmlEncode(request("OptionName"))%><%=strOptionID%>">All</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if request("List")="Consumer" then%>
		<td class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;Consumer&nbsp;&nbsp;</font>
	<%else%>
		<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="Countries.asp?List=Consumer&OptionName=<%=Server.HtmlEncode(request("OptionName"))%><%=strOptionID%>">Consumer</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if request("List")="Commercial" then%>
		<td class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;Commercial&nbsp;&nbsp;</font>
	<%else%>
		<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="Countries.asp?List=Commercial&OptionName=<%=Server.HtmlEncode(request("OptionName"))%><%=strOptionID%>">Commercial</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>

	<%if request("List")="Tablet" then%>
		<td class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;Tablet&nbsp;&nbsp;</font>
	<%else%>
		<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="Countries.asp?List=Tablet&OptionName=<%=Server.HtmlEncode(request("OptionName"))%><%=strOptionID%>">Tablet</a>&nbsp;&nbsp;</font>
	<%end if%>
	</td>
	</tr></table><BR>

<%	
	'Response.Write "<font size=2 face=verdana><a href=""mailto: max.yu@hp.com;kenneth.berntsen@hp.com?Subject=Master Country List - Change Request&Body=Please make the following changes to the Master Country List and send a full extract of the matrix with changes highlighted to bob.jack@hp.com;bob.giblin@hp.com;maureen.oloughlin@hp.com"">Request Changes</a></font><BR><BR>"
	if request("OptionID") <> "" then	
		if request("OptionName") <> "" then
			Response.Write "<font size=2 face=verdana>Display: Only countries with " & Server.HtmlEncode(request("OptionName")) & " localization.</font><BR><BR>"
		end if
		Response.Write "<font size=2 face=verdana><a href=""Countries.asp"">Show All Countries</a> | <a href=""javascript:window.print();"">Print This List</a></font><BR><BR>"
	else
		Response.Write "<font size=2 face=verdana><a href=""javascript:window.print();"">Print This List</a></font><BR><BR>"
	end if
	rs.Open "spListCountriesWithLocalizations " & clng(request("OptionID")),cn,adOpenForwardOnly
	Response.Write "<Table bgcolor=ivory width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=beige><TD><font size=1 face=verdana><b>Country</b></font></TD><TD><font size=1 face=verdana><b>Export</b></font></TD><TD><font size=1 face=verdana><b>Localizations</b></font></TD></tr>"
	strLastRegion = ""
	strCountry = ""

	do while not rs.EOF 
		if (rs("Region") & "" <> "") then
			if LastCountry <> rs("Country") & "" then
				if LastCountry <> "" then
					if trim(strLocalizations) = "" then
						strLocalizationOutput = "&nbsp;"
					elseif trim(mid(strLocalizations,2)) = "" then
						strLocalizationOutput = "&nbsp;"
					else
						strLocalizationOutput = mid(strLocalizations,2)
					end if
					response.write "<TR><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & LastCountry & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strExport & "</font></TD><TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strLocalizationOutput & "</font></TD></TR>"
					strLocalizations=""
				end if
				LastCountry = rs("Country")
			end if
			if strLastRegion <> rs("Region") & "" then
				strRegion = mid(left(rs("Region"),len(rs("Region"))-1),2)
				Response.Write "<TR bgcolor=MediumAquamarine><TD colspan=3 style=""BORDER-TOP: gray thin solid""><font size=2 face=verdana><b>" & strRegion & "</b></font></TD></TR>"
				strLastRegion = rs("Region") & ""
			end if
		if (request("List") = "" or request("List") = "All") or (request("List") = "Tablet" and rs("Tablet"))or (request("List") = "Commercial" and rs("Commercial"))or (request("List") = "Consumer" and rs("Consumer")) then
			if rs("Export") then
				strExport = "Yes"
			else
				strExport = "No"
			end if
			if not isnull(rs("OptionConfig")) then
				strLocalizations = strLocalizations & ",<a href=Countries.asp?List=" & request("List") & "&OptionID=" & rs("RegionID") & "&OptionName=" & rs("OptionConfig") & ">" & rs("OptionConfig") & "</a>" 
			end if
		end if
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
