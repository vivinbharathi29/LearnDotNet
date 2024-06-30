<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdGo_onclick() {
if (optEmail.checked)
	{
	location.href="mailto:pulsar.support@hp.com?SUBJECT=Feedback";
	window.opener=self;
	window.close();
	}
else if (optNewUser.checked)
	{
	if (txtNewUser.value == "")
		{
		alert("The Name or Email address of the new user is required.");
		txtNewUser.focus();
		}
	else
		{
	    location.href = "mailto:" + txtNewUser.value + "?SUBJECT=Excalibur Registration Instructions&Body=To use Excalibur all you need to do is open this link:%0a%09http://<%=Application("Excalibur_ServerName")%>/Excalibur/Excalibur.asp%0a%0aThe first time you open Pulsar it will tell you that it can't find your user account.  Just click 'Register Now' to setup your account.";
		window.opener=self;
		window.close();
		}
	}
else if (optEmergency.checked)
	{
	if (frmMain.txtJustification.value == "")
		{
		alert("Please explain why this is an emergency.");
		frmMain.txtJustification.focus();
		}
	else if (frmMain.txtMessage.value == "")
		{
		alert("Please enter a short text message to send to the Excalibur Support team.");
		frmMain.txtMessage.focus();
		}
	else
		{
		frmMain.submit();
		}
	}
}

function optNewUser_onclick() {
	cmdGo.value = "  Create Email   "
}

function optEmergency_onclick() {
	cmdGo.value = "Send Message"
}

function optEmail_onclick() {
	cmdGo.value = "  Create Email   "
}

function txtNewUser_onkeypress() {
	optNewUser.checked=true;
	optNewUser_onclick();
}

function txtJustification_onkeypress() {
	optEmergency.checked=true;
	optEmergency_onclick();
}

function txtMessage_onkeypress() {
	optEmergency.checked=true;
	optEmergency_onclick();
}

function window_onload() {
	if (txtStartOption.value=="2")
		{
		optNewUser.checked=true;
		optNewUser_onclick();
		txtNewUser.focus();
		}
}

//-->
</SCRIPT>
</HEAD>
<Style>
body{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}

td{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}
</STYLE>
<BODY bgColor=ivory LANGUAGE=javascript onload="return window_onload()">
<%
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn

	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPartner
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
	Set	rs = cm.Execute 
	
	set cm=nothing	
		
	if not (rs.EOF and rs.BOF) then
		CurrentUserPartner = rs("PartnerID") & ""
		CurrentUserID = rs("ID") & ""
	end if
	rs.Close


    set rs = nothing
    cn.Close
    set cn = nothing
    
    if trim(currentuserpartner) = "1" then
        strAddUserRowDisplay = ""
    else
        strAddUserRowDisplay = "none"
    end if
%>


<P>

<font size=3 face=verdana><b>Contact the Excalibur Support Team</b><BR><BR></font>There are&nbsp;three contact options 
available:<BR>
     <BR>
<table bgcolor=lavender bordercolor=lightsteelblue border=1 cellpadding=2 cellspacing=0>
	<TR><TD valign=top>
		<INPUT type=radio checked id=optEmail name=optType value=1 LANGUAGE=javascript onclick="return optEmail_onclick()"></TD><TD><b>Normal Email</b><BR>Used for bug reports, requests, suggestions, feedback, questions, etc.<BR><BR>     </TD></TR>
	
	<TR style="display:<%=strAddUserRowDisplay%>">
	
	<TD valign=top><INPUT type=radio id=optNewUser name=optType value=2 LANGUAGE=javascript onclick="return optNewUser_onclick()"></TD><TD><b>Request&nbsp;New User Account <font color=red></b>- For HP Employees and Contractors only</font><BR>Send an email to the person or people 
      who need a new user account.&nbsp; The email contains instructions for 
      registering with Excalibur.
      
      <TABLE width="100%">
		<TR>
			<TD nowrap><b>New User's Name or Email Address:&nbsp;<font size=2 color=red>*</font></b></TD>
			<TD width=100%><INPUT style="WIDTH: 100%" id=txtNewUser name=txtNewUser LANGUAGE=javascript onkeypress="return txtNewUser_onkeypress()"></TD>
		</TR>
      </TABLE>

                       </TD></TR>
	<%
		dim strSupportTime
		dim TimeParts
		dim TimeColor
		strSupportTime = formatdatetime(now(),vbshorttime)
		TimeParts = split(strSupportTime,":")
		if cint(TimeParts(0)) > 17 or cint(TimeParts(0)) < 8 or cint(Weekday(Date)) = 1 or cint(Weekday(Date)) = 7 then
			TimeColor= "red"
		else
			timecolor = "black"
		end if
		if cint(TimeParts(0)) > 12 then
			strSupportTime = TimeParts(0)-12 & ":" & Timeparts(1)
		elseif cint(TimeParts(0)) = 0 then
			strSupportTime = "12:" & Timeparts(1)
		else
			strSupportTime = TimeParts(0) & ":" & Timeparts(1)
		end if
		if cint(TimeParts(0)) >= 12 then
			strSupportTime	= "<font color=" & timecolor & ">" & strSupportTime & " " & "PM</font>"
		else
			strSupportTime	= "<font color=" & timecolor & ">" & strSupportTime & " " & "AM</font>"
		end if

        if (cint(TimeParts(0)) > 20 or cint(TimeParts(0)) < 7) and lcase(trim(CurrentDomain)) <> "americas" then
            response.write "<tr style=""display:none"">"
        else
            response.write "<tr>"
        end if
	%>
	
	<TD valign=top><INPUT type=radio id=optEmergency name=optType value=3 LANGUAGE=javascript onclick="return optEmergency_onclick()"></TD><TD><b>Emergency Support Request</b> - <font color=red>For urgent requests only.</font><br>This option sends a text message to the Excalibur Support Team 
      cell phones.&nbsp;&nbsp;Support Team Local Time: <b><%=strSupportTime%></b>
      <form id=frmMain method=post action="SupportSend.asp">
      <TABLE width="100%">
		<TR>
			<TD><b>Issue Type:&nbsp;<font size=2 color=red>*</font></b></TD>
    		<TD>
                <input name="optIssueType" type="radio" value=1> ExCoP
                <input checked name="optIssueType" type="radio" value=0> Other
            </TD>
		</TR>

		<TR>
			<TD><b>Justification:&nbsp;<font size=2 color=red>*</font></b></TD>
			<TD><INPUT style="WIDTH: 100%" id=txtJustification name=txtJustification LANGUAGE=javascript onkeypress="return txtJustification_onkeypress()"></TD>
		</TR>
		<TR>
			<TD nowrap><b>Text Message:&nbsp;<font size=2 color=red>*</font></b></TD>
			<TD width="100%"><INPUT style="WIDTH: 100%" id=txtMessage name=txtMessage maxlength=100 LANGUAGE=javascript onkeypress="return txtMessage_onkeypress()"></TD>
		</TR>
      </TABLE>
      </form>
                        </TD></TR>
</table></P>
<TABLE width="100%">
<TR>
	<TD align=right>
		<INPUT type="button" value="  Create Email   " id=cmdGo name=cmdGo LANGUAGE=javascript onclick="return cmdGo_onclick()">
		<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
	</TD>
</TR>
</TABLE>

     <INPUT type="hidden" id=txtStartOption name=txtStartOption value="<%=request("StartOption")%>">


</BODY>
</HTML>
