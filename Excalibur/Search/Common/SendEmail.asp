<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<!-- #include file="../../includes/emailwrapper.asp" -->
<html>
<head>
	<title>Send Email</title>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
		function window_onload() {
			if (typeof (txtSuccess) != "undefined") {
				if (txtSuccess.value == "1") {
					alert("Email Sent.");
					window.parent.opener = 'X';
					window.parent.open('', '_parent', '')
					window.parent.close();
				}
				else
					alert("Error sending email.");
			}
			else
				alert("Error sending email.");

		}
//-->
	</script>
	<style type="text/css">
		td
		{
			font-family: Verdana;
			font-size: xx-small;
		}
		body
		{
			font-family: Verdana;
			font-size: xx-small;
		}
	</style>
</head>
<body onload="window_onload();">
	Sending Email...<br />
	<br />
	<%

	dim strStyle
	strStyle="<STYLE>td{FONT-FAMILY: Verdana;FONT-SIZE:xx-small;vertical-align:top;}thead{background-color:beige;}h1{FONT-FAMILY: Verdana;FONT-SIZE:x-small;}body{FONT-FAMILY: Verdana;FONT-SIZE:xx-small;}A:link{COLOR: Blue;}A:visited{COLOR: Blue;}A:hover{COLOR: red;}</STYLE>"

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	cn.CommandTimeout = 180

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserEmail


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

		CurrentUserID = rs("ID")
		CurrentUserEmail = rs("Email") & ""
	else
		Response.Redirect "Excalibur.asp"
	end if
	rs.Close


	if CurrentUserEmail = "" then
		Response.Write "<BR><BR><font size=2 face=verdana>Email is only available for registered Excalibur users. Click <a href=""pm.asp"">here</a> to launch Excalibur and register for a user account.</font"
	else

		dim EmailsToSend
		dim strTo
		dim strCC
		dim strBody
		dim strSubject
		dim i
		dim strID
		dim strReciept

		strReciept = ""
		EmailsToSend = split(request("chkTo"),",")

		for each strID in EmailsToSend
			strTo = trim(request("txtTo" & trim(strID)))
			strCC = trim(request("chkCC" & trim(strID)))
			strSubject = request("txtSubject")
			if trim(request("txtNotes")) = "" then
				strBody = request("txtEmailTable" & trim(strID+1))
			else
				strBody = request("txtNotes") & "<BR><BR>" & request("txtEmailTable" & trim(strID+1))
			end if


			Set oMessage = New EmailWrapper
			oMessage.From = CurrentUserEmail
'			if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'				oMessage.To= "matt.hamilton@hp.com"
'				if strCC <> "" then
'					strBody = "<BR/>CC=" & strCC & "<BR/>" & strBody
'				end if
'			else
				oMessage.To= strTO '"max.yu@hp.com"
				if strCC <> "" then
					oMessage.cc = strCC '"max.yu@hp.com"
				end if
'			end if
			oMessage.Subject = strSubject
			oMessage.HTMLBody = strStyle & strBody
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send
			Set oMessage = Nothing


			strReciept = strReciept & "From: " & CurrentUserEmail & "<BR>"
			strReciept = strReciept & "To: " & strTo & "<BR>"
			if strCC <> "" then
				strReciept = strReciept & "Cc: " & strCC & "<BR>"
			end if
			strReciept = strReciept & "Sent: " & Now & "<BR><BR>"
			strReciept = strReciept & "Subject: " & strSubject & "<BR><BR>"
			strReciept = strReciept & strBody & "<BR><HR>"

		next

		response.write strReciept

		if trim(request("chkCopyMe")) = "1" and trim(strReciept) <> "" then
			Set oMessage = New EmailWrapper
			oMessage.From = CurrentUserEmail
'			if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'				oMessage.To= "matt.hamilton@hp.com"
'			else
				oMessage.To= CurrentUserEmail '"max.yu@hp.com"
'			end if
			oMessage.Subject = request("txtSubject")
			oMessage.HTMLBody = strStyle & strReciept
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send
			Set oMessage = Nothing
		end if
	end if

	set rs = nothing
	cn.Close
	set cn = nothing

	%>
	<input id="txtSuccess" type="text" value="1" />
</body>
</html>
