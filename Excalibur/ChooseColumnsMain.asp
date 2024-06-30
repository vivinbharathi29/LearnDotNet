<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript" src="../_ScriptLibrary/jsrsClient.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    
function lstAvailable_ondblclick() {
	if (lstAvailable.selectedIndex != "undefined")
		{
			if (lstAvailable.selectedIndex != -1 )
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[lstAvailable.selectedIndex].text,lstAvailable.options[lstAvailable.selectedIndex].value);
				lstAvailable.options[lstAvailable.options.selectedIndex]=null;
			}
		}

}

function lstNew_ondblclick() {
	if (lstNew.selectedIndex != "undefined")
		{
			if (lstNew.selectedIndex != -1 )
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[lstNew.selectedIndex].text,lstNew.options[lstNew.selectedIndex].value);
				lstNew.options[lstNew.options.selectedIndex]=null;
			}
		}
}

function cmdAdd_onclick() {
	var i;
	for (i=0;i<lstAvailable.length;i++)
		{
			if (lstAvailable.options[i].selected)
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[i].text,lstAvailable.options[i].value);
			}
		}

	for (i=lstAvailable.length-1;i>-1;i--)
		{
			if (lstAvailable.options[i].selected)
			{
				lstAvailable.options[i]=null;
			}
		}
		
		
}

function cmdRemove_onclick() {
	var i;
	for (i=0;i<lstNew.length;i++)
		{
			if (lstNew.options[i].selected)
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[i].text,lstNew.options[i].value);
			}
		}

	for (i=lstNew.length-1;i>-1;i--)
		{
			if (lstNew.options[i].selected)
			{
				lstNew.options[i]=null;
			}
		}
		


}

function cmdOK_onclick() {
	var i;
	var strNewSetting="";
	
	if (lstNew.length == 0)
	    {
	    alert("You must select at least one column to display.")
	    return;
	    }
	for(i=0;i<lstNew.length;i++)
	{
		if (strNewSetting=="")
			strNewSetting = lstNew.options[i].value + ":1"  
		else
			strNewSetting = strNewSetting + "," + lstNew.options[i].value + ":1"
	}

	for(i=0;i<lstAvailable.length;i++)
	{
		if (strNewSetting=="")
			strNewSetting = lstAvailable.options[i].value + ":0"
		else
			strNewSetting = strNewSetting + "," + lstAvailable.options[i].value + ":0"
	}

	frmUpdate.txtSetting.value = strNewSetting;
	frmUpdate.submit();
}
function cmdReset_onclick() {
	frmUpdate.txtSetting.value = "";
	frmUpdate.submit();
}
function cmdCancel_onclick() {
	window.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<BR>
<table style="width:100%"><TR><TD width=10>&nbsp;</TD><TD>
<%


dim cn
dim rs
dim p
dim cm
dim ListArray
dim ListItem
dim ListPartArray


set cn = server.CreateObject("ADODB.Connection")
cn.ConnectionString = Session("PDPIMS_ConnectionString") 
cn.Open

set rs = server.CreateObject("ADODB.recordset")

dim CurrentUSer
dim CurrentUSerID

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
	end if
	rs.Close



%>


<font size=3 face=verdana><b>Choose Columns</b></font><BR><BR>
<table border=0 width="100%">
<tr>
<td nowrap width=50%>
	<font face=verdana size=2><b>Available Columns:</b></font><BR>	<SELECT style="WIDTH:100%" size=10 id=lstAvailable name=lstAvailable  multiple LANGUAGE=javascript ondblclick="return lstAvailable_ondblclick()">
		<%
		ListArray=split(request("lstColumns"),",")
		for each ListItem in ListArray
		    ListPartArray = split(Listitem,":")
			if trim(ListPartArray(1)) = "0" then
				Response.Write "<option value=""" & trim(ListPartArray(0)) & """>" & trim(ListPartArray(0)) & "</option>"
			end if
		next
		%>	</SELECT>
</td>
	<td valign=top width=30 align=middle><BR>
		<INPUT style="width:30" type="button" value=">" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()">
		<INPUT style="width:30" type="button" value="<" id=cmdRemove name=cmdRemove LANGUAGE=javascript onclick="return cmdRemove_onclick()">
	</td>
<td nowrap width=50%>
	<font face=verdana size=2><b>Selected Columns:</b></font><BR>	<SELECT style="WIDTH:100%" size=10 id=lstNew name=lstNew multiple LANGUAGE=javascript ondblclick="return lstNew_ondblclick()">
		<%
		for each ListItem in ListArray
		    ListPartArray = split(Listitem,":")
			if trim(ListPartArray(1)) = "1" then
				Response.Write "<option value=""" & trim(ListPartArray(0)) & """>" & trim(ListPartArray(0)) & "</option>"
			end if
		next
		%>	</SELECT>
</td>
</tr>
</table>
<form id=frmUpdate action="ChooseColumnsSave.asp" method=post>

<table width=100%>
<TR>
	<TD colspan=2><HR>
	</TD>
</TR>
<TR>
    <TD align=left>
    <TD align=right>
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Reset" id=cmdReset name=cmdReset LANGUAGE=javascript onclick="return cmdReset_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>
</table>

<%
set rs = nothing
cn.Close
set cn = nothing
%>
<INPUT type="text" style="display:none" id=txtEmployeeID name=txtEmployeeID value="<%=CurrentUserID%>">
<INPUT type="text" style="display:none" id=txtUserSettingsID name=txtUserSettingsID value="<%=request("UserSettingsID")%>">
<INPUT type="text" style="display:none" id=txtSetting name=txtSetting value="">
</form>
</td>
</tr>
</TABLE>
</BODY>
</HTML>
