<%@ Language=VBScript %>

<%

	  Response.Buffer = True
	  Response.ExpiresAbsolute = Now() - 1
	  Response.Expires = 0
	  Response.CacheControl = "no-cache"

%>

<html>
<STYLE>
A:link,A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>

<head>
<TITLE>Excalibur</TITLE>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
 <link rel="shortcut icon" href="favicon.ico" >

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function ShowSupportPage(){
	var NewTop;
	var NewLeft;
	
	NewLeft = (screen.width - 700)/2
	NewTop = (screen.height - 500)/2

	window.open ("mobilese/today/Support.asp","_blank","width=700,height=600,left=" + NewLeft + ",top=" + NewTop + ",menubar=no,toolbar=no,scrollbars=Yes,resizable=Yes,status=No")
}

function Recent_onclick() {
	if (window.event.srcElement.className != "")
		{
		TurnArrowsOff();
		document.all[window.event.srcElement.className].style.display = "";
		
		}
}

function TurnArrowsOff(){
	document.all["OTSArrow"].style.display = "none";
	document.all["TodayArrow"].style.display = "none";	
}

function OTS_onclick() {
	var strID;
	TurnArrowsOff();
	document.all["OTSArrow"].style.display = "";	

	window.parent.frames("RightWindow").navigate("suddenimpact/default.html");
}

function Today_onclick() {
	var strID;
	TurnArrowsOff();
	document.all["TodayArrow"].style.display = "";	

	window.parent.frames["RightWindow"].navigate ("mobilese/today/Today.asp");
}

function Today_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";

}

function Today_onmouseout() {
	window.event.srcElement.style.color = "blue";
}

function Home_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";

}

function Home_onmouseout() {
	window.event.srcElement.style.color = "blue";
}
//-->
</SCRIPT> 
</head>
<body LANGUAGE=javascript  bgcolor=Beige topMargin=0>

<%

	dim cn
	dim rs
	dim CurrentUser
	dim CurrentUserName
	dim CurrentUserID
	dim CurrentUserPartner
	dim CurrentUserDivision
    dim CurrentUserParterType

	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	'Get User
	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserName = rs("Name")
		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = trim(rs("PartnerID") & "")
	end if
	rs.Close

	set rs = nothing
	cn.close
	set cn=nothing
	
	
%>
<table ID=Recent cellSpacing="1" cellPadding="1"  border="0" LANGUAGE=javascript onclick="return Recent_onclick()">
	<TR><TD colspan=2><font face=verdana size="2"><strong>Navigate</strong></font></TD></TR>
<TR><TD colspan=2><table cellSpacing="1" cellPadding="1"  border="0">
  <tr>
    <%if request("ID") <> "" then%>
		<td width="14" nowrap><img ID="TodayArrow" style="display:none" src="images/red.gif" align="right" WIDTH="6" HEIGHT="11"></td>
    <%else%>
		<td width="14" nowrap><img ID="TodayArrow" style="display:" src="images/red.gif" align="right" WIDTH="6" HEIGHT="11"></td>
    <%end if%>
    <td nowrap><font face=verdana size="1" color="blue"><u ID=Today LANGUAGE=javascript onclick="return Today_onclick()" onmouseover="return Today_onmouseover()" onmouseout="return Today_onmouseout()">Today Page</u> </font></td></tr>
  <tr>
    <td width="14" nowrap><img ID="OTSArrow" style="display:none" src="images/red.gif" align="right" WIDTH="6" HEIGHT="11"></td>
    <td nowrap><font face=verdana size="1" color="blue"><u ID=OTS LANGUAGE=javascript onclick="return OTS_onclick()" onmouseover="return Home_onmouseover()" onmouseout="return Home_onmouseout()">Sudden Impact Info</u> </font></td></tr>
  </table>


 <br>
<font face=verdana size="2"><strong>Feedback</strong><br>
<table cellSpacing="1" cellPadding="1"  border="0">
  <tr>
    <td nowrap width="14"></td>
    <td nowrap><font size=1><a href="javascript: ShowSupportPage();">Send Email</a></font></td></tr>
  </table>
</font>

</body>
</html>
	
