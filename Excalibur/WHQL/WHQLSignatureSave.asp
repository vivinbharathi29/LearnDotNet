<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = 1;
	window.parent.close();
}


//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

dim oMessage
dim strBody

strBody = ""


	'Get FROM mail address.
	dim strConnect
	dim cn
	dim CurrentUserID
	dim CurrentUserName
	dim CurrentUserEmail
	
	CurrentUserEmail = ""
	
	strConnect = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = strConnect
	cn.CommandTimeout = 60
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


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
		CurrentUserName = rs("Name")
		CurrentUserEmail = rs("Email")
	end if
	rs.Close
	cn.close
	set rs = nothing
	set cn = nothing

	Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	'Set oMessage.Configuration = Application("CDO_Config")		
	if CurrentUserEmail = "" then
		oMessage.From = "bruce.ridings@hp.com"
	else
		oMessage.From = CurrentUserEmail
	end if

	oMessage.To = "Testsigrequest@hp.com" 
	
	if request("chkCCEmail")="on" and CurrentUserEmail <> "" then
		oMessage.cc = CurrentUserEmail
	end if
		
	oMessage.Subject = "Test Sig Request: " & request("txtVendorName") & " " & request("txtDriverCategory") & " " & request("txtVersionPass") & " - " & request("txtPlatform")
	
	strBody = "<b><font size=3 face=verdana>" & "WHQL Test Signature Request" & "</font></b>"
	strBody = strBody & "<br><br><TABLE>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Vendor Name: " & "</td><td>" & request("txtVendorName") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Driver Category: " & "</td><td>" & request("txtDriverCategory") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Operating System: " & "</td><td>" & request("txtOperatingSystem") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Version/Pass: " & "</td><td>" & request("txtVersionPass") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Link to Driver: " & "</td><td>" & "<a href=""" & request("txtLinktoDriver") & """>" & request("txtLinktoDriver") & "</a></td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "File Name: " & "</td><td>" & request("txtFileName") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Platform(s) Supported: " & "</td><td>" & request("txtPlatform") & "</td></tr>"
	strBody = strBody & "<tr><td valign=top width=220>" & "Date Needed: " & "</td><td>" & request("txtDateNeeded") & "</td></tr>"
	strBody = strBody & "</TABLE>"
		
	oMessage.HTMLBody = "<font size=2 face=verdana>" & strBody & "</font>"
						
	oMessage.Send 
	Set oMessage = Nothing 			

%>				
</BODY>
</HTML>
