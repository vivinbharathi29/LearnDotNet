<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<STYLE>
BODY
{
	FONT-SIZE: xx-small;
}
P
{
	FONT-SIZE: xx-small;
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

.MenuBar TD.ButtonSelected
{
    COLOR: black;
    BACKGROUND-COLOR: wheat
}
</STYLE>

<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script language="javascript" src="_ScriptLibrary/jsrsClient.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "_ScriptLibrary/sort.js" -->
	var oPopup = window.createPopup();

function SetMeView(value) {

	var expireDate = new Date();

	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "MeList=" + value + ";expires=" + expireDate.toGMTString() + ";";

	window.location.reload(true);

}
//-->
</SCRIPT> 
<LINK href="style/wizard%20style.css" type=text/css rel=stylesheet >
</HEAD>
<BODY>

<H3>My Information</H3>
<font size=1 face=verdana color=red>This page is under development.</font><BR><BR>
<FONT face=verdana>
	<%
		dim cn
		dim rs
		
		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")


		dim CurrentUser 
		dim CurrentUserName
		dim CurrentUserID
		dim strFavs
		dim strFavCount
		dim strTitleColor
		on error resume next
		strTitleColor = "#0000cd"
		strTitleColor = Request.Cookies("TitleColor")
		if strTitleColor = "" then
			strTitleColor = "#0000cd"
		end if
		on error goto 0
		
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
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserName = rs("Name") & ""
			CurrentUserID = rs("ID")
			strFavs = trim(rs("Favorites") & "")
			strFavCount = trim(rs("FavCount") & "")
		end if
		rs.Close

	%>

<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>
  
  <TR>
    <TD nowrap width=100 bgColor=cornsilk><STRONG><FONT 
      size=1>Name:</FONT></STRONG></TD>
    <TD><FONT size=1><%=CurrentUserName%></FONT></TD></TR>
</TABLE><BR>
<Table border=1 bordercolor=Ivory cellspacing=0 cellpadding =2 Id=menubar Class=MenuBar><TR bgcolor=<%=strTitleColor%>>
<%if strDisplayedList = "Vacation" or strDisplayedList = "" then%>
	<TD class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;Vacation&nbsp;&nbsp;</font></td>
<%else%>
	<TD><font size=1 face=verdana><a href="javascript:SetMeView('Vacation');">&nbsp;&nbsp;Vacation</A>&nbsp;&nbsp;</font></td>
<%end if%>
<%if strDisplayedList = "Status" then%>
	<td class="ButtonSelected"><font size=1 face=verdana>&nbsp;&nbsp;Status&nbsp;&nbsp;</font></td>
<%else%>
	<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="javascript:SetMeView('Travel');">Status</A>&nbsp;&nbsp;</font></td>
<%end if%>
	<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="javascript:SetMeView('Travel');">Review</A>&nbsp;&nbsp;</font></td>
	<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="javascript:SetMeView('Travel');">Reports</A>&nbsp;&nbsp;</font></td>
	<td><font size=1 face=verdana>&nbsp;&nbsp;<a href="javascript:SetMeView('Travel');">Favorites</A>&nbsp;&nbsp;</font></td>
</tr></table>

<font size=1 face=verdana><b>None</b></font>


<INPUT type="hidden" id=txtID name=txtID value=<%=request("ID")%>>
<INPUT type="hidden" id=txtUser name=txtUser value=<%=CurrentUserID%>>
<INPUT type="hidden" id=txtFavs name=txtFavs value="<%=strFavs%>">
<INPUT type="hidden" id=txtFavCount name=txtFavCount value="<%=strFavCount%>">
</BODY>
</HTML>
