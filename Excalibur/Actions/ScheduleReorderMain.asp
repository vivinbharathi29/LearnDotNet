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
	var strNewOrder="";
	
	for(i=0;i<lstAvailable.length;i++)
		lstAvailable.options[i].selected=true;
	cmdAdd_onclick();
	
	for(i=0;i<lstNew.length;i++)
	{
		strNewOrder = strNewOrder + lstNew.options[i].value + ","
	}
	frmUpdate.txtNewOrder.value = strNewOrder;
	frmUpdate.submit();
}


function cmdCancel_onclick() {
	window.returnValue = 0;
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

set cn = server.CreateObject("ADODB.Connection")
cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
cn.Open

set rs = server.CreateObject("ADODB.recordset")

dim strFavs
dim FavArray
dim i
dim strFav
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


<font size=3 face=verdana><b>Reorder Roadmap Items</b></font><BR><BR>
<table border=0 width=100%>
<tr>
<td nowrap width=50%>
	<font face=verdana size=2><b>Current Order:</b></font><BR>	<SELECT style="WIDTH:100%" size=10 id=lstAvailable name=lstAvailable  multiple LANGUAGE=javascript ondblclick="return lstAvailable_ondblclick()">
		<%
			rs.Open "spListActionRoadmap " & clng(request("ID")),cn,adOpenForwardOnly			do while not rs.EOF				Response.Write "<option value=""" & rs("ID") & """>" & rs("Summary") & "</option>"
				rs.MoveNext			loop			rs.Close
		%>	</SELECT>
</td>
	<td valign=top width=30 align=middle><BR>
		<INPUT style="width:30" type="button" value=">" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()">
		<INPUT style="width:30" type="button" value="<" id=cmdRemove name=cmdRemove LANGUAGE=javascript onclick="return cmdRemove_onclick()">
	</td>
<td nowrap width=50%>
	<font face=verdana size=2><b>New Order:</b></font><BR>	<SELECT style="WIDTH:100%" size=10 id=lstNew name=lstNew multiple LANGUAGE=javascript ondblclick="return lstNew_ondblclick()">
	</SELECT>
</td>
</tr>
</table>

<table width=100%>
<TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>
</table>

<%
set rs = nothing
cn.Close
set cn = nothing
%>
<form id=frmUpdate action="ScheduleReorderSave.asp" method=post>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtNewOrder name=txtNewOrder value="">
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=CurrentUserID%>">
</form>
</td>
</tr>
</TABLE>
</BODY>
</HTML>
