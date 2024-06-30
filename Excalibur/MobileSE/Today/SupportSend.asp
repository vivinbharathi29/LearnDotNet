<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value == "1")
		{
		alert("Text message sent successfully.");
		window.opener=self;
		window.close();
		}
	else
		{
		alert("An error occurred while sending this text message.  Please try again or select a different contact option.");
		window.history.back();
		}
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
body{
	FONT-FAMILY: verdana;
	FONT-SIZE: x-small;
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()" bgColor=Ivory>
<b>Sending Text Message</b><BR><BR>
Justification: <%=request("txtJustification")%><BR>
Message:<%=request("txtMessage")%><BR><BR>
Attempting to send message...

<%
	dim strSuccess
	strSuccess = "0"

	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open
	
	dim CurrentUser
	dim CurrentDomain
	dim CurrentUserEmail
	
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if
	

	set rs = server.CreateObject("ADODB.Recordset")
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
		
	if (rs.EOF and rs.BOF) then
		CurrentUserEmail  = "max.yu@hp.com"
	else
		CurrentUserEmail = rs("Email") & ""
	end if	
	rs.close
	set rs = nothing
	cn.Close
	set cn = nothing

    strExtraNotifications = ""
    if trim(request("optIssueType")) = "1" then
        strExtraNotifications = ";8327883159@tmomail.net"
    end if

	on error resume next
	Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	'Set oMessage.Configuration = Application("CDO_Config")
	oMessage.From = Currentuseremail
	oMessage.To= "pulsar.support@hp.com;releaseteam@hp.com;"
	oMessage.Subject = "Excalibur Emergency" 
	oMessage.HTMLBody = request("txtMessage") & "<BR>Justification: " & request("txtJustification")
	oMessage.Send 
	Set oMessage = Nothing 	

	Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	'Set oMessage.Configuration = Application("CDO_Config")
	oMessage.From = Currentuseremail
	oMessage.To= "ExcaliburSupport@cingularme.com;2818020711@vtext.com;2817873444@txt.att.net" & strExtraNotifications
	oMessage.Subject = "Excalibur Emergency" 
	oMessage.HTMLBody = request("txtMessage") 
	oMessage.Send 
	if err.number = 0 then
		strSuccess= "1"
		Response.Write "Done"
	else
		Response.Write "Failed"
	end if
	Set oMessage = Nothing 	
	
	
	
%>


<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
