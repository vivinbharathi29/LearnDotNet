<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	txtRevDisplay.focus();
}


function cmdImageButton_onmouseover() {
	window.event.srcElement.style.cursor = "default";
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
}

function cmdImageButton_onmouseout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
}

function cmdImageButton_onmousedown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";

}

function cmdImageButton_onmouseup() {
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	ImageButton_Pressed(window.event.srcElement.name);
}

function cmdImageButton_onkeydown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";
}

function cmdImageButton_onkeyup() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	if (window.event.keyCode !=9)
		ImageButton_Pressed(window.event.srcElement.name);
}

function ImageButton_Pressed(ID){
	if (txtRevDisplay.value=="")
		txtRevDisplay.value="2";
	else if (isNumeric(txtRevDisplay.value))
		txtRevDisplay.value=parseInt(txtRevDisplay.value)+1;
	else
		alert("Internal Rev must be an integer.");
	txtRevDisplay.focus();
}

function cmdImageButton_onfocusout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
	window.event.srcElement.style.backgroundColor = "gainsboro";

}
function isNumeric(sText)
{
   var ValidChars = "-0123456789";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }
function cmdOK_onclick() {
	if (txtRevDisplay.value=="")
		{
	    frmUpdate.txtOSCode.value = txtOSCodeDisplay.value;
	    frmUpdate.txtPrepStatus.value =  cboPrepStatus.options[cboPrepStatus.options.selectedIndex].value;
	    frmUpdate.txtRev.value = txtRevDisplay.value;
		frmUpdate.submit();
		}
	else if (! isNumeric(txtRevDisplay.value))
		{
		alert("Internal Rev number be an integer.");
		txtRevDisplay.focus();
		}
	else if (parseInt(txtRevDisplay.value)< 1)
		{
		alert("Internal Rev number be a positive integer.");
		txtRevDisplay.focus();
		}
else {
        frmUpdate.txtOSCode.value = txtOSCodeDisplay.value;
        frmUpdate.txtPrepStatus.value = cboPrepStatus.options[cboPrepStatus.options.selectedIndex].value;
        frmUpdate.txtRev.value = txtRevDisplay.value;
		frmUpdate.submit();
		}	
}

function cmdCancel_onclick() {
	window.parent.close();
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

<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<LINK href="../style/wizard%20style.css" type=text/css rel=stylesheet>

<%
	dim cn
	dim rs
	dim strRev
    dim CurrentUserID
    dim CurrentUser
    dim CurrentUserPINGroup
    dim strGroupName
    dim strOSCode
    dim strPreinstallPrepStatus
    	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
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
	
    CurrentUserPINGroup = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
        if rs("workgroupid")= 15 then
            CurrentUserPINGroup = 1
            strGroupName = "Houston "
        elseif rs("workgroupid")= 22 then
            CurrentUserPINGroup = 2
            strGroupName = "Taiwan "
        else
            CurrentUserPINGroup = 1
            strGroupName = "Houston "
	    end if
	end if
	rs.Close

    rs.open "spGetPreinstallDeliverableProperties " & clng(request("ID")),cn,adOpenStatic
    if rs.eof and rs.bof then
        strOSCode = "0"
        strPreinstallPrepStatus = ""
    else
        strOSCode = trim(rs("OSCode") & "")
        strPreinstallPrepStatus= trim(rs("PreinstallPrepStatus") & "")
    end if	
    rs.Close

	rs.Open "spGetInternalRev " & clng(request("ID")) & "," & clng(CurrentUserPINGroup),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strRev = ""
	else
		strRev = trim(rs("PreinstallInternalRev") & "")
	end if
%>
    <font size=3 face=verdana><b>Preinstall Properties</b></font><BR>
	<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr>
            <td  width=120><b>Version&nbsp;ID:</b></td>
            <td><%=clng(request("ID"))%></td>
        </tr>
        <TR><TD><b>Internal&nbsp;Rev:&nbsp;&nbsp;</b></TD><TD width="100%"><table cellpadding=0 cellspacing=0 bordor=0><tr><td><INPUT id=txtRevDisplay name=txtRevDisplay style="MARGIN-TOP: -10px; VERTICAL-ALIGN: middle; WIDTH: 58px; HEIGHT: 22px" 
      size=7 value="<%=strRev%>">&nbsp;<input type="image" src="../images/PLUS2.GIF" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; MARGIN-TOP: 1px; BORDER-LEFT: thin outset; WIDTH: 18px; BORDER-BOTTOM: thin outset; TOP: 0px; HEIGHT: 18px; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()"             
      id="cmdRev" name ="cmdRev" title ="Increment Internal Rev" LANGUAGE="javascript" 
      onmouseover="return cmdImageButton_onmouseover()" 
      onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" 
      onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" 
      onkeyup="return cmdImageButton_onkeyup()" 
      WIDTH="18" HEIGHT="18"></TD><td valign=middle nowrap>&nbsp;-&nbsp;<%=strGroupName%></TD></tr></table></TD></TR>
     <TR><TD width=120><b>Prep&nbsp;Status:&nbsp;&nbsp;</b></TD>
        <TD width="100%">
            <select style="width:120px" id="cboPrepStatus" name="cboPrepStatus">
                <%if strPreinstallPrepStatus="0" then%>
                    <option selected value=0></option>
                <%else%>
                    <option value=0></option>
                <%end if %>
                <%if strPreinstallPrepStatus="1" then%>
                    <option value=1 selected>Complete</option>
                <%else%>
                    <option value=1>Complete</option>
                <%end if %>
                <%if strPreinstallPrepStatus="2" then%>
                    <option value=2 selected>Not Required</option>
                <%else%>
                    <option value=2>Not Required</option>
                <%end if %>
            </select>
        </TD>
     </TR>
     <TR><TD width=120><b>OS&nbsp;Code:&nbsp;</b></TD>
        <TD width="100%"><input id="txtOSCodeDisplay" name="txtOSCodeDisplay" style="width:120px" type="text" maxlength=5 value="<%=strOSCode%>"></TD>
     </TR>
	</table>
	<table width=100%>
		<TR><TD width=100% align=right>
			<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
			<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
			</td></tr>
	</table>
<%
	set rs = nothing
	cn.Close
	set cn = nothing
%>
<form id="frmUpdate" method="post" action="InternalRevSave.asp">
<INPUT style="display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT style="display:none" type="text" id=txtRev name=txtRev value="">
<INPUT style="display:none" type="text" id=txtTeam name=txtTeam value="<%=CurrentUserPINGroup%>">
<input style="display:none" id="txtOSCode" name="txtOSCode" type="text" value="">
<input style="display:none" id="txtPrepStatus" name="txtPrepStatus" type="text" value="">

</form>
</BODY>
</HTML>
