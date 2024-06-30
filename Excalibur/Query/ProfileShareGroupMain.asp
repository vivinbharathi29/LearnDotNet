<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!-- #include file = "../_ScriptLibrary/sort.js" -->
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
		
		for (i=0;i<event.srcElement.length;i++)
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


function optShare_onclick() {
	AccessRow.style.display="";
}

function optStopShare_onclick() {
	AccessRow.style.display="none";
}

function window_onload() {
	if (typeof(frmMain.txtName) != "undefined")
		frmMain.txtName.focus();
}

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
    frmMain.cmdOK.disabled = true;
    frmMain.cmdCancel.disabled = true;
	var blnFailed = false;
	var i;
	var blnFound = false;
	
	if (typeof(frmMain.txtName) != "undefined")
		if (frmMain.txtName.value == "")
			{
            frmMain.cmdOK.disabled = false;
            frmMain.cmdCancel.disabled = false;
			alert("Group name is required");
			blnFailed = true;
			frmMain.txtName.focus();
			}
	else if (typeof(frmMain.chkEmployee) != "undefined")
        {
        for (i=0;i<frmMain.chkEmployee.length;i++)
            {
            if (frmMain.chkEmployee[i].checked)
                {
                blnFound = true;
                break;
                }
            } 
        if (!blnFound)
            {
            frmMain.cmdOK.disabled = false;
            frmMain.cmdCancel.disabled = false;
            alert("You must select at least one member");
            blnFailed = true;
            frmMain.txtName.focus();
            }
        }
	if (!blnFailed)
		{
		frmMain.submit();
		}
	
	
}

function mouseover_Column(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function mouseout_Column(){
	window.event.srcElement.style.color="black";
}


function DeleteGroup(ID){
    if (confirm("Deleting this group will make it unavailable for all profiles.\r\rAre you sure you want to delete this group?") )
        {
        frmMain.chkDelete.checked = true;
        frmMain.submit();
        }
  
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" bgcolor=Ivory>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">


<%	if trim(request("GroupID")) <> "" then %>
	<Table cellpadding=0 cellspacing=0 width=100%><tr><td><font size=3 face=verdana><b>Edit Group</b></font></td><td align=right><a href="javascript: DeleteGroup(<%=clng(request("GroupID"))%>);">Delete Group</a></td></tr></Table>
    <font size=1 face=verdana color=red><br>Note: Any changes made to this group will update all of the Report Profiles where it is used.</font>
<%else%>
    <font size=3 face=verdana><b>Add Group</b></font>
<%end if%>

<form id= frmMain action="ProfileShareGroupSave.asp" method=post>
<input style="display:none" id="chkDelete" name="chkDelete" type="checkbox" value=1>
<%
	
	dim cn, rs
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


	set rs = server.CreateObject("ADODB.recordset")



	'Get User
    dim CurrentUser
	dim CurrentDomain
	dim CurrentUserID
	
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

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
    else
        Response.Write("Pulsar User Not Found. Please submit a Pulsar Support request for assistance.")
        Response.End()
	end if
	rs.Close

	dim strProfileName
	dim strProfileMembers
	dim GroupArray
	
	if trim(request("GroupID")) <> "" then
	    rs.open "spGetSharedProfile " & request("GroupID"),cn,adOpenStatic
	    if rs.eof and rs.bof then
            strProfileName = ""
	        strProfileMembers = ""
	    else
    	    GroupArray = split(rs("Setting") & "","|")
	        strProfileName = groupArray(0)
	        strProfileMembers = "," & replace(groupArray(1)," ","") & ","
	    end if
        rs.close    
    else
        strProfileName = ""
        strProfileMembers = ""
    end if	
	



%>
<table border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>
  <tr>
    <td valign=top width="160"><strong><font size=2>Group&nbsp;Name:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>
        <input style="width:100%" id="txtName" name="txtName" type="text" value="<%=strProfileName%>">
    </td>
  </tr>
  <tr>
    <td valign=top width="160"><strong><font size=2>Members:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>

<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; OVERFLOW-X: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; WIDTH: 280px;HEIGHT: 320px; BACKGROUND-COLOR: white" id=DIV1>
	<TABLE ID=TableEmployee width="100%">
	<THEAD><TR style="POSITION: relative;TOP: expression(document.getElementById('DIV1').scrollTop-2)">
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 0 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Name&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 2 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Group&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 3 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Division&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 4 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Manager&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 5 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Email&nbsp;</TD>
	</TR></THEAD>
	<%

	dim strDivision
	dim strNotCurrentMembers
	dim strRow
	
	rs.Open "splistEmployees",cn,adOpenStatic
	do while not rs.EOF
	
		if trim(rs("PartnerID")) <> "1" then
			strDivision =  rs("Partner") & ""
		else
			select case trim(rs("Division") & "")
			case "1"
				strDivision = "Mobile"
			case "2"
				strDivision = "bPC"
			case "3"
				strDivision = "Mobile"
			case "4"
				strDivision = "Workstations"
			case "5"
				strDivision = "cPC"
			case else
				strDivision = "Other"
			end select
		end if
		if rs("Active") = 1 and instr(strProfileMembers,"," & trim(rs("ID")) & ",") > 0  then
      	    response.Write "<TR class=""Row"" Email=""" & rs("Email") & """ ID=""" & rs("Name") & """><TD nowrap><input id=""chkEmployee"" name=""chkEmployee"" style=""height:16px;width:16px"" type=""checkbox"" checked value=""" & rs("ID") & """></TD>"
			response.Write "<TD nowrap>" & rs("Name") & "</TD>"
			response.Write "<TD nowrap>" & rs("Workgroup") & "</TD>"
			response.Write "<TD nowrap>" & strDivision & "</TD>"
			response.Write "<TD nowrap>" & rs("Manager") & "</TD>"
			response.Write "<TD nowrap>" & rs("Email") & "</TD>"
			response.Write "</TR>"
		end if
		rs.MoveNext
	loop
	rs.Close

	rs.Open "splistEmployees",cn,adOpenStatic
	do while not rs.EOF
	
		if trim(rs("PartnerID")) <> "1" then
			strDivision =  rs("Partner") & ""
		else
			select case trim(rs("Division") & "")
			case "1"
				strDivision = "Mobile"
			case "2"
				strDivision = "bPC"
			case "3"
				strDivision = "Mobile"
			case "4"
				strDivision = "Workstations"
			case "5"
				strDivision = "cPC"
			case else
				strDivision = "Other"
			end select
		end if
		if rs("Active") = 1 and clng(currentuserid) <> clng(rs("ID")) and instr(strProfileMembers,"," & trim(rs("ID")) & ",") = 0  then
      	    response.Write "<TR class=""Row"" Email=""" & rs("Email") & """ ID=""" & rs("Name") & """><TD nowrap><input id=""chkEmployee"" name=""chkEmployee"" style=""height:16px;width:16px"" type=""checkbox"" value=""" & rs("ID") & """></TD>"
			response.Write "<TD nowrap>" & rs("Name") & "</TD>"
			response.Write "<TD nowrap>" & rs("Workgroup") & "</TD>"
			response.Write "<TD nowrap>" & strDivision & "</TD>"
			response.Write "<TD nowrap>" & rs("Manager") & "</TD>"
			response.Write "<TD nowrap>" & rs("Email") & "</TD>"
			response.Write "</TR>"
		end if
		rs.MoveNext
	loop
	rs.Close
	%>
	
	</TABLE>
            
    </div>
          			

    </td>
  </tr>
  </table>
<table width=100%><TR><TD align=right style="width:100%">
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
</TD></TR></TABLE>
<%


	set rs = nothing
	cn.Close
	set cn = nothing
	

%>
<INPUT type="hidden" id=txtGroupID name=txtGroupID value="<%=trim(request("GroupID"))%>">
</form>
</BODY>
</HTML>
