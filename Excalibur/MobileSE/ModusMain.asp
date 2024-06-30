<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<%

	dim cn, rs, strSQL, strProgram, strFamily, strVersion, WorkflowStep
	
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	'Get User
	dim ModusUser
	dim HPUser
	dim CurrentUser

	CurrentUser = lcase(Session("LoggedInUser")) 
	if left(CurrentUser,15)="excaliburweb\odmtest" or left(CurrentUser,17)="excaliburweb\moduslink" or left(CurrentUser,9)="moduslink" then
		ModusUser=true
	else
		ModusUser=false
	end if

	if left(CurrentUser,8)="americas" or left(CurrentUser,4)="emea" or left(lcase(Session("LoggedInUser")),11)="asiapacific" then 'if left(lcase(Session("LoggedInUser")),7)="excaliburweb"  or left(lcase(Session("LoggedInUser")),6)="compal" then
		HPUser=true
	else
		HPUser=false
	end if
	
%>

<html>
<head>
<meta name="VI60_defaultClientScript" content=JavaScript>
<title>DIB Information - HP Restricted</title>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oPopup = window.createPopup();

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

function PRow_onmouseover() {
   	window.event.srcElement.parentElement.style.color = "red";
	window.event.srcElement.parentElement.style.cursor = "hand";
}

function PRow_onmouseout() {
   	window.event.srcElement.parentElement.style.color = "blue";
}

//-->
</SCRIPT>
</head>

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

<% if ModusUser or HPUser then %>
<body background="../images/shadow.gif"><!-- This is the Headder Table --><!-- This is the Headder Table -->
<table border="0" cellPadding="1" cellSpacing="1" width="100%" border="0">
  <tr>
    <td nowrap width="200"><img src="../images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
	<td><font size="5" face="Tahoma" color="#006697"><b>DIB Deliverable Information</b></font></td>
  </tr>
</table>
        
<table cellPadding="1" border="0" cellSpacing="1" width="100%"> 
  <tr>
    <td nowrap width="200" style="VERTICAL-ALIGN: top" id="side">
	<h1>
	<BR>
	<font face=verdana size=4 color=yellow>Links</font><br>
 
 <BR>    
      Suggestions<br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="mailto:max.yu@hp.com?Subject=Processe Suggestion">Process</a><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><A HREF="mailto:max.yu@hp.com?Subject=SE Web Site Suggestion">Web Site</A><br>
      
      </h1>
</td>

   <td style="VERTICAL-ALIGN: top">      
    
	<table border="0" cellpadding="2"  width=100%>
		<TR><td width=10>&nbsp;</td>
		<tr><td colspan=2><table width=100% border=0 cellpadding=1 cellspacing=1>
		<tr><td colspan=2 nowrap valign=top ><b><font size=3 color=#006697>Active Project Information</font></b><br><hr width=130% color=steelblue><br></td></tr>

<%	
	strSQL = "spGetODMProducts"
	rs.Open strSQL,cn,adOpenForwardOnly

	rs.MoveFirst

	dim columncounter
	dim strSeries
	
	columncounter = 1
	Do while not rs.EOF
		if rs("Division") = 1 and rs("Name") & "" <> "Test Product"  and rs("active") then
			if columncounter =1 then
				Response.Write "<tr>"
			end if
			
			Response.Write "<td width=35% valign=top ><b><font size=3 color=#006697 FONT-FAMILY: Tahoma>" + rs("Name") + " " + rs("Version") + "</font></b><br>"

			set rs2 = server.CreateObject("ADODB.recordset")
			rs2.Open  "spListbrands4Product " & rs("ID") & ",1",cn,adOpenForwardOnly
			strSeries = ""
			do while not rs2.EOF
				if trim(rs2("SeriesSummary") & "") = "" then
					strSeries = ""
				else
					SeriesArray = split(rs2("SeriesSummary"),",")
					for i = 0 to ubound(SeriesArray)
						if rs2("StreetName") <> "" then
							strSeries =  strSeries & rs2("StreetName") & " " 
						end if
						strSeries = strSeries & seriesArray(i) 
						if trim(rs2("Suffix") & "") <> "" then
							strSeries = strSeries & " " & trim(rs2("Suffix") & "")
						end if
						strSeries = strSeries & "<BR>"
					next
				end if
				rs2.MoveNext
			loop
			rs2.Close
			set rs2 = nothing
			
			if trim(strSeries) <> "" then
				Response.write "<font color=green size=1>" & strSeries & "</font>"
			end if
			
			Response.Write "<a target=""_blank"" href=""ModusImageInfo.asp?ID=" & rs("ID") & """>Image and RTM info</a><br>"
			Response.Write "<a target=""_blank"" href=""..\image\DeliverableMatrix.asp?DIBOnly=1&ProdID=" & rs("ID") & """>DIB Deliverable List</a><br><br></td>"
			
			if columncounter = 3 then
				Response.Write "</tr>"
				columncounter = 1
			else
				columncounter = columncounter + 1
			end if
			
	end if
			rs.movenext
		loop
	%>	
	</table>

<div id="cc"><h1>HP Restricted</h1><p>Last Update <%=Date%></p></div>
<TEXTAREA style="Display:none" rows=4 cols=80 id=txtHidden name=txtHidden>

</TEXTAREA>
<%
	rs.Close

else
Response.write CurrentUser 
'Response.Redirect "../NoAccess.asp?Level=1"
end if %>
</body>
</html>
