<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<TITLE>Choose Product Group</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../../includes/bundleConfig.inc" -->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() {

	if (Reassign.txtNotes.value == "" && Reassign.chkUrgent.checked)
		{
			alert("You must enter notes for Urgent requests.");
			Reassign.txtNotes.focus();
			return;
		}
	Reassign.txtName.value = Reassign.cboOwner.options[Reassign.cboOwner.selectedIndex].text;
	Reassign.submit();
}

function chkUrgent_onclick() {
	if (Reassign.chkUrgent.checked)
		{
		ReqNotes.style.display = "";
		Reassign.txtNotes.focus();
		}
	else
		ReqNotes.style.display = "none";
}


function AddCycle(){

	var NewTop;
	var NewLeft;
	var strResult;

	NewLeft = (screen.width - 655) / 2;
	NewTop = (screen.height - 650)/2;

	modalDialog.open({ dialogTitle: 'Add New Product Group', dialogURL: '../../program/program.asp', dialogHeight: 400, dialogWidth: 500, dialogResizable: false, dialogDraggable: true });

	/*strResult = window.showModalDialog("../../program/program.asp","","dialogWidth:655px;dialogHeight:480px;edge: Sunken;maximize:Yes;;center:Yes; help: No;resizable: Yes;status: No") 
	if (typeof(strResult) != "undefined")
	{
		//window.location.reload(true);
        window.location.href=window.location.href;
	}*/
}

//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  10/06/2016 - Harris, Valerie - Change dialogs to JQuery dialogs     
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();
}
//-->
</SCRIPT>
</HEAD>
<BODY onload="window_onload();" bgcolor="ivory">
<%

if request("ID") = ""  then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim strLoaded
	
	strLoaded = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b>Choose Product Group for 
<%
    rs.open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
    if rs.eof and rs.bof then
        response.write "Product"    
    else
        response.write rs("Name")  
    end if
    rs.close
%>
</b></font>
<form ID=frmMain method=post action="ProductCycleSave.asp">
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="150" nowrap valign=top><b>Product Group:</b>&nbsp;</td>
		<td width="100%"><a href="javascript: AddCycle();">Add New Product Group</a><BR>
        <%  
            dim strName
            rs.open "spListProgramsForProductAll " & clng(request("ID")),cn,adOpenForwardOnly
            do while not rs.eof
                strname = rs("FullName")
                
                if trim(rs("OnProduct")) = "1" then
                    response.write "<input checked id=""chkCycle"" name=""chkCycle"" type=""checkbox"" value=""" & rs("ID") & """ CycleName=""" & strname &  """> " & strname & "<BR>"
                    strLoaded = strLoaded & "," & rs("ID")
                else
                    response.write "<input id=""chkCycle"" name=""chkCycle"" type=""checkbox"" value=""" & rs("ID") & """ CycleName=""" & strName & """> " & strname & "<BR>"
                end if
                rs.movenext
            loop
            rs.close
            if trim(strLoaded) <> "" then
                strLoaded = mid(strLoaded,2)
            end if
        %>
		</td>
	</tr>
</table>
    <INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
    <input id="txtCycleList" name="txtCycleList" type="hidden" value="">
    <input id="txtCycleLoaded" name="txtCycleLoaded" type="hidden" value="<%=strLoaded%>">
</form>
<%

	set rs = nothing
	set cn = nothing
end if


%>

</BODY>
</HTML>
