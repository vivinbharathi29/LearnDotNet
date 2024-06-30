<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../../_ScriptLibrary/sort.js" -->

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
	if (typeof(frmMain.cboEmployee) != "undefined")
		frmMain.cboEmployee.focus();
}

function cmdCancel_onclick() {
	window.parent.close();
}

function cmdOK_onclick() {
	var blnFailed = false;
	
	if (typeof(frmMain.cboEmployee) != "undefined")
		if (frmMain.cboEmployee.selectedIndex ==0)
			{
			alert("Name is Required");
			blnFailed = true;
			frmMain.cboEmployee.focus();
			}
		else
			cboEmployee_onchange();

	if (!blnFailed)
		{
		if (typeof(frmMain.cboEmployee) != "undefined")		
			frmMain.txtAction.value= "1"; //Action 1=Add, 2=Remove, 3=Update
		else if(frmMain.optStopShare.checked)		
			frmMain.txtAction.value= "2"; //Action 1=Add, 2=Remove, 3=Update
		else if(frmMain.optShare.checked)		
			frmMain.txtAction.value= "3"; //Action 1=Add, 2=Remove, 3=Update
		frmMain.submit();
		}
	
	
}

function cmdNewGroup_onclick(){
	var strResult;
	strResult = window.showModalDialog("ProfileShareGroup.asp","","dialogWidth:450px;dialogHeight:480px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
        {
			frmMain.cboEmployee.options[frmMain.cboEmployee.length] = new Option(strResult[0],strResult[1]);
			frmMain.cboEmployee.selectedIndex = frmMain.cboEmployee.length - 1
            frmMain.cmdEditGroup.disabled = false;
        }
		//window.location.href =  "ProfileShare.asp?ID=" + txtProfileID.value;
}

function cmdEditGroup_onclick(){
	var strResult;
	var GroupID;
	
	GroupID = frmMain.cboEmployee.options[frmMain.cboEmployee.selectedIndex].value;
	
	strResult = window.showModalDialog("ProfileShareGroup.asp?GroupID=" + GroupID,"","dialogWidth:450px;dialogHeight:520px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
	    {
	    if(strResult[1]=="0") 
	        {
            frmMain.cboEmployee.options[frmMain.cboEmployee.selectedIndex] = null;
	        frmMain.cboEmployee.selectedIndex=0;
            frmMain.cmdEditGroup.disabled = true;
	        }
	    else
            frmMain.cboEmployee.options[frmMain.cboEmployee.selectedIndex].text = strResult[0];
        }
}

function EditGroup(ID){
	var strResult;
	var GroupID;
	
	strResult = window.showModalDialog("ProfileShareGroup.asp?GroupID=" + ID,"","dialogWidth:450px;dialogHeight:520px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
	    {
	    if(strResult[1]=="0") 
	        {
            frmMain.optStopShare.checked = true;
            cmdOK_onclick();
	        }
	    else
            GroupNameCell.innerText = strResult[0];
        }
}

function cboEmployee_onchange() {
	//KeyString = "";
	frmMain.txtEmployeeName.value = frmMain.cboEmployee.options[frmMain.cboEmployee.selectedIndex].text;
	frmMain.txtEmployeeID.value = frmMain.cboEmployee.options[frmMain.cboEmployee.selectedIndex].value;
    if (frmMain.AddType.value == "2")
        {
        if (frmMain.cboEmployee.selectedIndex > 0)
            frmMain.cmdEditGroup.disabled = false;
        else
            frmMain.cmdEditGroup.disabled = true;
        }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" bgcolor=Ivory>
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font size=3 face=verdana><b>
<%	if trim(request("EmployeeID")) <> "0" then %>
	Edit Permissions
<%elseif trim(request("AddType")) <> "2" then%>
	Add Person
<%else%>
	Add Group
<%end if%>
</b></font>

<%
	
	dim cn, rs
	dim blnEmployeeFound
	dim strEmployeeName
	dim EditYesValue
	dim EditNoValue
	dim DeleteYesValue
	dim DeleteNoValue
	dim strNameLabel
	dim GroupArray
	
	blnEmployeeFound = true

    if trim(request("AddType")) <> "2" then
        strNameLabel = "Name:"
    else
        strNameLabel = "Group:"
    end if

	if request("CanEdit") = "True" then
		EditYesValue = "checked"
		EditNoValue = ""
	else
		EditYesValue = ""
		EditNoValue = "checked"
	end if

	if request("CanDelete") = "True" then
		DeleteYesValue = "checked"
		DeleteNoValue = ""
	else
		DeleteYesValue = ""
		DeleteNoValue = "checked"
	end if
	
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
    


	if trim(request("AddType") <> "2") then
	    if trim(request("EmployeeID")) <> "0" and isnumeric(trim(request("EmployeeID"))) then 
		    rs.Open "spGetEmployeeByID " & clng(request("EmployeeID")),cn,adOpenStatic
		    if rs.EOF and rs.BOF then
			    blnEmployeeFound = false
		    else
			    strEmployeeName = rs("Name") & ""
		    end if
		    rs.Close
	    elseif trim(request("EmployeeID")) <> "0" then
		    blnEmployeeFound = false
	    end if	
    else
	    if trim(request("EmployeeID")) <> "0" and isnumeric(trim(request("EmployeeID"))) then 
		    rs.Open "spGetEmployeeUserSetting " & clng(request("EmployeeID")),cn,adOpenStatic
		    if rs.EOF and rs.BOF then
			    blnEmployeeFound = false
		    else
			    strEmployeeName = left(rs("Setting") & "", instr(rs("Setting") & "","|")-1)
		    end if
		    rs.Close
	    elseif trim(request("EmployeeID")) <> "0" then
		    blnEmployeeFound = false
	    end if	
    end if
	if not blnEmployeeFound then
		Response.Write "Not enough information supplied."
	else
%>
<form id= frmMain action="ProfileSharePropertiesSave.asp" method=post>
<table border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>
  <tr>
    <td valign=top width="160"><strong><font size=2><%=strNameLabel%></font><font color=red size=1> *</font></strong></td>
    <td valign=top>


<%
	if trim(request("EmployeeID")) <> "0" then 
		if trim(request("AddType") <> "2") then
    		Response.write strEmployeeName
	    else
	        response.Write "<Table cellpadding=0 cellspacing=0 width=""100%""><TR><TD width=""100%"" id=""GroupNameCell"">" & strEmployeeName & "</TD><TD valign=top>&nbsp;&nbsp;<a href=""javascript: EditGroup(" & clng(request("EmployeeID")) & ");"">Edit</a></TD></TR></table>"
	    end if
	else
		if trim(request("AddType") <> "2") then
		    Response.Write "<SELECT id=cboEmployee name=cboEmployee LANGUAGE=javascript onkeydown=""return combo_onkeydown()"" onfocus=""return combo_onfocus()"" onchange=""return cboEmployee_onchange()"" onkeypress=""return combo_onkeypress()"" onclick=""return combo_onclick()"">"
		    rs.Open "spGetEmployees",cn,adOpenStatic
		    Response.Write "<OPTION selected value=""""></Option>"
		    do while  not rs.EOF
    			if rs("Active") = 1 then
				    Response.Write "<OPTION value=" & trim(rs("ID")) & ">" & rs("Name") & "</OPTION>"
			    end if
			    rs.MoveNext
		    loop
		    rs.Close
		    Response.Write "</SELECT>"
        else
		    Response.Write "<SELECT style=""width:250"" id=cboEmployee name=cboEmployee LANGUAGE=javascript onkeydown=""return combo_onkeydown()"" onfocus=""return combo_onfocus()"" onchange=""return cboEmployee_onchange()"" onkeypress=""return combo_onkeypress()"" onclick=""return combo_onclick()"">"
		    rs.Open "spGetEmployeeUserSettings " & clng(Currentuserid) & ",4",cn,adOpenStatic
		    Response.Write "<OPTION selected value=""""></Option>"
		    do while  not rs.EOF
			    GroupArray = split(rs("Setting"),"|")
			    Response.Write "<OPTION value=" & trim(rs("ID")) & ">" & GroupArray(0) & "</OPTION>"
			    rs.MoveNext
		    loop
		    rs.Close
		    Response.Write "</SELECT>"
		    response.Write "&nbsp;<input id=""cmdNewGroup"" name=""cmdNewGroup"" type=""button"" value=""Add"" Language=javascript onclick=""cmdNewGroup_onclick();"">"
		    response.Write "&nbsp;<input disabled id=""cmdEditGroup"" name=""cmdEditGroup"" type=""button"" value=""Edit"" Language=javascript onclick=""cmdEditGroup_onclick();"">"
	    end if
	end if
	

%>
    </td>
  </tr>
<%	if trim(request("EmployeeID")) <> "0" then %>
  <tr>
 <%else%>
  <tr style=display:none>
 <%end if%>
    <td valign=top width="160"><strong><font size=2>Action:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>
    <INPUT type="radio" checked id=optShare name=optAction LANGUAGE=javascript onclick="return optShare_onclick()"> Share&nbsp;&nbsp;&nbsp;
    <INPUT type="radio" id=optStopShare name=optAction LANGUAGE=javascript onclick="return optStopShare_onclick()"> Stop Sharing
    </TD>
	</TR>
  <tr ID=AccessRow>
    <td valign=top width="160"><strong><font size=2>Permissions:</font><font color=red size=1> *</font></strong></td>
    <td valign=top>
    <table><TR><TD>
    <b>Edit&nbsp;Profile:</b>
    </TD>
    <TD>
    <INPUT type="radio" <%=EditYesValue%> id=optEdit name=optEditPermission value="1"> Yes
    <INPUT type="radio" <%=EditNoValue%> id=optNoEdit name=optEditPermission value="0"> No
    </TD></TR>
	<TR><TD>
    <b>Delete&nbsp;Profile:</b></TD>
    <TD>
    <INPUT type="radio" <%=DeleteYesValue%> id=optDelete name=optDeletePermission value="1"> Yes
    <INPUT type="radio" <%=DeleteNoValue%> id=optNoDelete name=optDeletePermission value="0"> No<BR>
    </TD></TR></Table>
    </TD>
	</TR>
  </table>
<table width=100%><TR><TD align=right style="width:100%">
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
</TD></TR></TABLE>
<%
	end if
	set rs = nothing
	cn.Close
	set cn = nothing
	

%>
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=trim(request("EmployeeID"))%>">
<INPUT type="hidden" id=txtEmployeeName name=txtEmployeeName value="<%=strEmployeeName%>">
<INPUT type="hidden" id=txtAction name=txtAction value="">
<INPUT type="hidden" id=AddType name=AddType value="<%=trim(request("AddType"))%>">
<INPUT type="hidden" id=txtProfileID name=txtProfileID value="<%=request("ProfileID")%>">
</form>
</BODY>
</HTML>
