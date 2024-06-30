<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function copyClipboard(Type) {
	if (!document.all)
		return; // IE only
	
	var Clip;
	
	if (Type==1)
		{
		clipboardData.clearData("text");
		Clip = frmMain.txtFormEncoded.createTextRange();
		Clip.execCommand("Copy");
		}
	else if (Type==2)
		{
		if (frmMain.txtPage.value=="http://16.81.19.70//global.asa")
			alert("Error page was glabal.asa");
		else
			{
			clipboardData.clearData("text");
			Clip = frmMain.txtPage.createTextRange();
			Clip.execCommand("Copy");
			}
		}
	else
		{
		if (frmMain.txtPage.value=="http://16.81.19.70//global.asa")
			alert("Error page was glabal.asa");
		else
			{
			if (Type==7) //Stop Impersonating
			    {
			    if (frmMain.txtErrorUserID.value != "" && frmMain.txtErrorUserID.value != "0" && frmMain.txtCurrentUserID.value != "" && frmMain.txtCurrentUserID.value != "0")
    	    	       {
    	    	        strID = window.showModalDialog("ChangeUserSave.asp?cboEmployee=0&txtEmployeeID=" + frmMain.txtCurrentUserID.value,"_blank","dialogWidth:655px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
    			        alert("You are no longer impersonating this user.")
				        }
                else
			        alert("You can not stop impersonating this user from here.  Try using the function on the Today page.")
			    }
			else if (Type==6) //Impersonate
			    {
			    if (frmMain.txtErrorUserID.value != "" && frmMain.txtErrorUserID.value != "0" && frmMain.txtCurrentUserID.value != "" && frmMain.txtCurrentUserID.value != "0")
    	    	    {
    	    	    strID = window.showModalDialog("ChangeUserSave.asp?cboEmployee=" + frmMain.txtErrorUserID.value + "&txtEmployeeID=" + frmMain.txtCurrentUserID.value,"_blank","dialogWidth:655px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
    			    alert("You are now impersonating this user.")
                    }
                else
			        alert("Could not impersonate this user")
			    }
			else if (Type==4)
				window.open (frmMain.txtFull.value);
			else
				{
				clipboardData.clearData("text");
				Clip = frmMain.txtFull.createTextRange();
				Clip.execCommand("Copy");
				}
			}
		}
		
}

function DefaultError(ID){
	if (ID==1)
		frmMain.txtCause.value="User entered invalid text into a field or parameter.";
	else if (ID==2)
		frmMain.txtCause.value="Network Issue.";
	else if (ID==3)
		frmMain.txtCause.value="Server Rebooted.";
	else if (ID==4)
		frmMain.txtCause.value="Normal Timeout.";
	else if (ID == 5)
	    frmMain.txtCause.value = "Unknown. Could Not Duplicate";
	else if (ID == 6)
	    frmMain.txtCause.value = "A service was restarted on the server.";
	else if (ID == 7)
	    frmMain.txtCause.value = "The Sudden Impact validator found and corrected these issues.";
	else if (ID == 8)
	    frmMain.txtCause.value = "The user did not apply enough filters to the report.";
}

function window_onload(){
    frmMain.txtForm.value = frmMain.txtFormLoaded.value;
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory language=javascript onload=window_onload()>
<LINK href="../../style/wizard style.css" type=text/css rel=stylesheet >
<%
	dim cn 
	dim rs
	dim strForm
	dim strColumn
	dim strLine
	dim strReferrer
	dim strFull
	dim CurrentUserID
	dim ErrorUser
	dim ErrorUserID
	dim ServerVariables
	dim strCopy
    dim strSIIDList
    dim strErrorDesc
	
    strSIIDList = ""

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	
	CurrentUserID = 0

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

	rs.CursorType = adOpenStatic
	'rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
	end if
	rs.Close

	ServerVariables = ""
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetAppError " & clng(request("ID")),cn,adOpenForwardOnly
	if not(rs.EOF and rs.BOF) then
		strForm = trim(rs("RequestForm") & "")
		if strForm = "" then
			strForm = trim(rs("RequestQueryString") & "")
		else
			strForm = trim(rs("RequestQueryString") & "") & "&" & trim(rs("RequestForm") & "")	
		end if 
		if left(strForm,1) = "&" then
			strForm = mid(strForm,2)
		end if
		strFile = trim(rs("ErrFile") & "")
		strLine = trim(rs("ErrLine") & "")
		strReferrer = trim(rs("Referrer") & "")
        strErrorDesc = trim(rs("ErrDescription") & "")
    	ErrorUser = trim(rs("AuthUser") & "")
	    ServerVariables = trim(rs("ServerVariables") & "")
	end if
	rs.Close

	if lcase(left(strForm,11)) = "500;http://" then
	    strFull = mid(strForm,5)
	elseif strForm <> "" then
		strFull = "http://16.81.19.70" & strFile & "?" & strForm
	else
		strFull = "http://16.81.19.70" & strFile
	end if
    if 	ErrorUser <> "" then
	    CurrentUser = lcase(ErrorUser)
	    if instr(currentuser,"\") > 0 then
		    CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		    Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
        else
            CurrentDomain = "americas"
            Currentuser = "mhamilton"
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

	    rs.CursorType = adOpenStatic
	    'rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 

	    set cm=nothing	
	    if not (rs.EOF and rs.BOF) then
		    ErrorUserID = rs("ID")
	    end if
	    rs.Close
    
    else
        ErrorUserID = 0
    end if

	set rs = nothing
	cn.Close
	set cn = nothing

  '  if left(ErrorUser,15) = "SiDataValidator" then
   '     strSIIDList = replace(strErrorDesc,",",", ")
  '  end if

    strCopy = request("CopyTo")
    if instr(strCopy,"," & trim(request("ID"))) > 0 then
        strCopy = replace(strCopy,"," & trim(request("ID")),"")
    elseif instr(strCopy, trim(request("ID")) & ",") > 0 then
        strCopy = replace(strCopy,trim(request("ID") & ","),"")
    elseif strCopy=trim(request("ID")) then
        strCopy = ""
    end if

%>
<form ID=frmMain action="AppErrorSave.asp" method=post>
<table width="100%"  cellspacing=0 cellpadding=2>
<TR>
	<TD valign=top><b>Cause:</b></TD><TD width="100%">
	<font size=1 face=verdana color=green>This field will be saved to <u>all</u> open error records with matching page, error, and line number.</font>
	<TEXTAREA id=txtCause style="WIDTH: 100%" name=txtCause rows=4></TEXTAREA>
    <table width="100%">
        <tr>
            <td rowspan=2 valign=top>
        	Common Causes:
            </td>
            <td>
            	<a href="javascript: DefaultError(1);">Invalid User Input</a><br>
        	    <a href="javascript: DefaultError(7);">SI Validator</a>&nbsp;&nbsp;&nbsp; 
            </TD> 
            <td>
        	    <a href="javascript: DefaultError(3);">Server Rebooted</a><br>
	            <a href="javascript: DefaultError(6);">Service Restarted</a>&nbsp;&nbsp;&nbsp; 
	        </TD>
            <td>
                <a href="javascript: DefaultError(2);">Network Issue</a><BR> 
	            <a href="javascript: DefaultError(4);">Normal Timeout</a>&nbsp;&nbsp;&nbsp; 
            </TD>
            <td>
	            <a href="javascript: DefaultError(8);">Buffer Size</a>&nbsp;&nbsp;&nbsp;<br> 
	            <a href="javascript: DefaultError(5);">Unknown</a> 
            </td>
        </tr>
    </table>
	</TD>
</TR>
<TR>
	<TD valign=top><b>Location:</b></TD><TD><TEXTAREA id=txtLocation style="WIDTH: 100%; BACKGROUND-COLOR: ivory" name=txtLocation rows=4 readOnly><% Response.Write "File: " & strFile & vbCrLf & "Line: " & strLine & vbCrLf & "Column: " & strColumn& vbCrLf & "Referrer: " & strReferrer %></TEXTAREA></TD>
</TR>
<TR>
	<TD valign=top><b>Input:</b></TD><TD><!--<font size=2 face=verdana>Copy&nbsp;To&nbsp;Clipboard:&nbsp;<a href="javascript:copyClipboard(1);">Encoded Input String</a>&nbsp;|&nbsp;<a href="javascript:copyClipboard(2);">Page URL</a>&nbsp;|&nbsp;<a href="javascript:copyClipboard(3);">Full URL</a>&nbsp;|&nbsp;--><a href="javascript:copyClipboard(4);">Reproduce This Error</a> | <a href="javascript:copyClipboard(6);">Impersonate User</a> | <a href="javascript:copyClipboard(7);">Stop Impersonating This User</a><BR></font><TEXTAREA id=txtForm style="WIDTH: 100%; HEIGHT: 250px; BACKGROUND-COLOR: ivory" name=txtForm rows=11 readOnly></TEXTAREA>
	<TEXTAREA id=txtFormLoaded style="display:none;WIDTH: 100%; HEIGHT: 360px; BACKGROUND-COLOR: ivory" name=txtFormEncoded rows=11 readOnly><%=replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(strForm & vbcrlf & "----------------------------------------------------------------------------------------" & vbcrlf & ServerVariables,"%20"," "),"%94",chr(148)),"%B4",chr(180)),"%28","("),"%29",")"),"%93","$"),"%0D%0A",vbcrlf),"%24","$"),"%3A",":"),"%25","%"),"%2F","/"),"%3D","="),"%22",""""),"%27","'"),"%3C","<"),"%3E",">"),"%3B",";"),"%5C","\"),"%2C",","),"+"," "),"&",vbcrlf) & strSIIDList%></TEXTAREA></TD>
	<TEXTAREA id=txtFormEncoded style="display:none;WIDTH: 100%; HEIGHT: 360px; BACKGROUND-COLOR: ivory" name=txtFormEncoded rows=11 readOnly><%=strForm%></TEXTAREA></TD>
</TR>
<TR>
	<TD valign=top nowrap><b>Also Close:</b></TD>
	<%if trim(strCopy) = "" then%>
	    <TD><INPUT type="text" readonly style="WIDTH: 100%; BACKGROUND-COLOR: ivory" ID=CopyMsg value="All other errors with the same error, page, and line number as this one."><INPUT type="text" style="display:none" id="txtCopyTo" name="txtCopyTo" value=""></td>
    <%else%>
	    <TD><INPUT type="text" style="WIDTH: 100%; BACKGROUND-COLOR: ivory" readonly id="txtCopyTo" name="txtCopyTo" value="<%=strCopy%>"></td>
    <%end if%>
</TR>
</table>
<INPUT type="hidden" id="txtID" name="txtID" value="<%=request("ID")%>">
<INPUT type="hidden" id="txtErrorUserID" name="txtErrorUserID" value="<%=ErrorUserID%>">
<INPUT type="hidden" id="txtPage" name="txtPage" value="<%="http://16.81.19.70" & strFile%>">
<INPUT type="hidden" id="txtFull" name="txtFull" value="<%=strFull%>">
<INPUT type="hidden" id="txtCurrentUserID" name="txtCurrentUserID" value="<%=CurrentUserID%>">

</form>
</BODY>
</HTML>
