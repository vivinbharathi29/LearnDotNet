<%@ Language=VBScript %>
<!-- #include file="includes/emailwrapper.asp" -->
<%

	  Response.Buffer = True
	  Response.ExpiresAbsolute = Now() - 1
	  Response.Expires = 0
	  Response.CacheControl = "no-cache"
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=mistyrose>

<font size=2 face=verdana><b>Suspicious activity has been detected.</b><br><br>
Your Excalibur access has been suspended and the system administrators have been notified.<br><br>Please contact <a href="mailto:max.yu@hp.com?Subject=Suspended Excalibur Account">Max Yu</a> if you have questions or need assistance.</font>
<%
	
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain
    dim CurrentUser
    dim AuthUser
	AuthUser = lcase(Session("LoggedInUser"))

	if instr(AuthUser,"\") > 0 then
		CurrentDomain = left(AuthUser, instr(AuthUser,"\") - 1)
		Currentuser = mid(AuthUser,instr(AuthUser,"\") + 1)
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
	end if
	rs.Close

    if currentuserid <> 31 and currentuserid <> 8 and currentuserid <> 5016  and currentuserid <> 1396 then

        'suspend their account if they have one
        if CurrentUserID <> 0 then
            cn.execute "spUpdateUserAccountSuspended " & clng(currentuserid)
        end if

	    Set oMessage = New EmailWrapper

        oMessage.From = "max.yu@hp.com" 
	    oMessage.To = "max.yu@hp.com;pulsar.support@hp.com"
        oMessage.Subject = "Suspicious Excalibur activity detected"
        if CurrentUserID <> 0 then
    	    oMessage.HTMLBody = "<font size=2 face=verdana color=black>Account suspended: " & Currentusername & "</font>"
	    else
    	    oMessage.HTMLBody = "<font size=2 face=verdana color=black>Suspicious activity detected from domain account: " & AuthUser & "<BR><BR>This person does not currently have an excalibur account.</font>"
        end if
        oMessage.DSNOptions = cdoDSNFailure
	    oMessage.Send 
	    Set oMessage = Nothing 	
   
        set rs = nothing
        cn.Close
        set cn = nothing
    	
        session.Abandon
    end if

%>
</BODY>
</HTML>
