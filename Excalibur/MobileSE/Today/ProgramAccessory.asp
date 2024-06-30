
<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<head>

  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		//for (i=0;i<event.srcElement.length;i++)
		for (i=event.srcElement.length-1;i>=0;i--)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}


function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


//-->
</SCRIPT>
</head>
<link rel="stylesheet" type="text/css" href="../../style/wizard%20style.css">
<body>
<font face=verdana>

<%
	dim cn
	dim rs
	dim cm
	dim p
	dim CnString
	dim strPM
	dim strProduct
	
	cnString =Session("PDPIMS_ConnectionString")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPhone
	dim CurrentUserID
	
	CurrentUser = lcase(Session("LoggedInUser"))
	CurrentUserID = 0
	
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
	Set	rs = cm.Execute 
	
	set cm=nothing	
		
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	end if
	rs.Close

	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetProductVersion"
		
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(request("ID"))
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	if not (rs.EOF and rs.BOF) then
		strPM = rs("AccessoryPMID") & ""
		strProduct = rs("DotsName") & ""
		rs.Close
	else
		strPM = ""
		strProduct = ""
	end if



	dim strPMList
	strPMList = ""
	rs.Open "spListAccessoryPMsAll",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(strPM) = trim(rs("ID")) then
			strPMList = strPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		else
			strPMList = strPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close	

set cm=nothing
set rs=nothing
set cn=nothing
	
if strProduct = "" then
	Response.Write "Unable to find the requested product."
else
	
%>
<H4><%=strProduct & " Properties"%></H4>

<FORM ACTION="ProgramAccessorySave.asp" METHOD="post" id="ProgramInput">

<table ID=tabPM border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>
  
   <tr>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Accessory&nbsp;PM:</font></strong></td>
    <td><SELECT id=cboPM name=cboPM style="WIDTH: 180px;" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value=0></OPTION>
			<%=strPMList%>			
		</SELECT>
	</td></tr>
</table>    

<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">

</FORM>
<%end if%>
</body>
</html>