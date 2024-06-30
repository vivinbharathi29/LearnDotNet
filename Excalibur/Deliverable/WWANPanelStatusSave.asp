<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
   <script src="../Scripts/jquery-1.10.2.js"></script>
 <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	    if (txtSuccess.value == "1") {
	        window.returnValue = 1;
	        if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.UpdatePanelStatusCallbackReloadCallback(1);
	            ClosePulsarPlusPopup();
	        }
	        else {
	            window.parent.close();
	        }
	    }
	    else
	        document.write("<BR><font size=2 face=verdana>Unable to update WWAN TSS.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update WWAN TSS.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim RowsChanged
	dim strSuccess
	dim cn
	dim cm
	
	strSuccess = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spUpdateWWANTTS"
		
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TTS", 200, &H0001, 10)
	p.Value = left(request("cboTTS"),10)
	cm.Parameters.Append p
	

	cm.Execute RowsChanged
	Set cm=nothing
	'if RowsChanged <> 1 or cn.Errors.count > 0 then : The sp is a simple update statement and as there are triggers for the deliverable table
    ' the rows affected will be changed so RowsChanged here is not a good indication. As we will have new pulsar plus code so just not use this value here to sokve the issue.
    if cn.Errors.count > 0 then	
		cn.RollbackTrans
	else
		cn.CommitTrans
		strSuccess = "1"
	end if


	set rs = server.CreateObject("ADODB.recordset")
	
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
	Set rs = cm.Execute 
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("email") & ""
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
	end if
	rs.Close




	'Notify Developer that the review is complete
	if strSuccess = "1" then
		'Get Deliverable Version Properties and Developer Email Address
		
		rs.Open "spGetVersionProperties4Web " &  clng(request("txtID")),cn,adOpenStatic		
		
		strTo = trim(rs("DeveloperEmail") & "")
		
		strVersion = rs("Version") & ""
		if rs("Revision")&"" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass")&"" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if		
		strBody = "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
		strBody = strBody &  "<TD>" & rs("VersionVendor") & "</TD>"
		strBody = strBody &  "<TD>" & rs("Name") & "</TD>"
		strBody = strBody &  "<TD nowrap>" & strVersion & "&nbsp;</TD>"
		strBody = strBody &  "<TD>" & rs("Modelnumber") & "&nbsp;</TD>"
		strBody = strBody &  "<TD nowrap>" & rs("Partnumber") & "&nbsp;</TD>"
		strBody = strBody &  "</TR>"
		strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;} A:visited{COLOR: blue} A:hover{COLOR: red}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD></tr>" & strBody & "</table><BR><font size=2 face=verdana><a href=""http://16.81.19.70/Excalibur.asp"">Open Pulsar</a> | <a href=""http://16.81.19.70/Release.asp?Action=1&ID=" & rs("VersionID") & """>Release This Deliverable</a></font>"
		strBody = "<font size=2 face=verdana color=black>The WWAN TTS for the following deliverable has been set to ""<b>" & request("cboTTS") & "</b>"" by the WWAN Engineer:<BR><BR></font>" & strBody 
		
		rs.Close
	
		'Send Mail
	
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		
		oMessage.From = CurrentUserEmail
		
		if strTo="" then
			oMessage.To = "max.yu@hp.com"
			oMessage.CC = CurrentUserEmail
		else
			oMessage.To = strTo 
			oMessage.CC = CurrentUserEmail & ";max.yu@hp.com"
		end if
		
		oMessage.Subject = "Deliverable TTS has been " & request("cboTTS") & " by WWAN Engineer"
		oMessage.HTMLBody = strBody
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 			
	end if

	set rs = nothing
	cn.Close
	set cn = nothing



%>


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">


</BODY>
</HTML>
