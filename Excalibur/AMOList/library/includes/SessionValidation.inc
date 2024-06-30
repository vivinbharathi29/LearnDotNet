<%
function Check_Expire()	
	if (Session("LoggedInUser") = "") then
		Check_Expire = true
	else
		Check_Expire = false
	end if
end function

function SessionValidation(ByVal sLink)
	if (Check_Expire) then
		Response.Redirect Session("PulsarServerPath")
		'Response.Redirect sLink
	end if
end function

sub SessionValidation2()
	if (Check_Expire) then
		Response.Redirect Session("PulsarServerPath")
		'Response.Redirect sLink
	end if
	if (not isobject(Session("rsAMORoles"))) then
		Response.Redirect Session("PulsarServerPath")
	end if
	'Session("Owner") = Session("LoggedInUser")
end sub

sub ValidateSession()
	if (Check_Expire) then
		Response.Redirect Session("PulsarServerPath")
	end if
end sub

sub RestoreSession()
	dim strCookie
	dim objErr

	if (Check_Expire) then
		' Initialize user session
		Session_Init

		' get the user's information and populate session variables rsGroups, FullName, IsODM, rsUser.
		if (Session("AMOUserID") = "") then
			'get the user's roles and UserID
			Call SetPermission()

			strCookie = GetDBCookie("AMO AMOUserID")
			if strCookie = "" then
				Session("AMOUserID") = 0
			else
				Session("AMOUserID") = GetDBCookie("AMO AMOUserID")
			end if

			strCookie = GetDBCookie("AMO IsODM")
			if strCookie = "" then
				Session("IsODM") = 0
			else
				Session("IsODM") = GetDBCookie("AMO IsODM")
			end if

			strCookie = GetDBCookie("AMO FullName")
			if strCookie = "" then
				session("FullName") = ""
			else
				session("FullName") = GetDBCookie("AMO FullName")
			end if

			strCookie = GetDBCookie("AMO BusinessSegmentIDs")
			if strCookie = "" then
				Session("BusinessSegmentIDs") = ""
			else
				Session("BusinessSegmentIDs") = GetDBCookie("AMO BusinessSegmentIDs")
			end if

		end if
	end if
end sub

function Session_Init()
	'Note: had to hard code /Excalibur folder in front of variables path in functions in /library/scripts/Calendar.js
	'Check for PRP and create the session variable if needed
	if Request.ServerVariables("HTTP_X_FORWARDED_HOST") <> "" then
		Session("ServerPath") = "/Excalibur"	'PRP automatically throws this on at the beginning
		Session("PrefixFullPath") = "https://" & Request.ServerVariables("HTTP_X_FORWARDED_HOST") & "/Excalibur/"	'example: https://prp.atlanta.hp.com/irs/
		Session("ServerURL") = "https://" & Request.ServerVariables("HTTP_X_FORWARDED_HOST") & "/Excalibur/"	'example: https://prp.atlanta.hp.com/irs/
		
		if Request.ServerVariables("HTTP_HOST") = "pulsarweb-pro.prp.ext.hp.com" then
                	Session("PulsarServerPath") = "https://pulsarweb-pro.prp.ext.hp.com/Excalibur"
            	else
                	Session("PulsarServerPath") = "https://virtual-prp010-itg.houston.hp.com/excalibur"
		end if
	else
		Session("ServerPath") = "/Excalibur"
		Session("PrefixFullPath") = "http://" & Request.ServerVariables("Server_Name") & "/Excalibur/"	'example: http://pulsarweb.usa.hp.com/irs/
		Session("ServerURL") = "http://" & Request.ServerVariables("Server_Name")	'example: http://pulsarweb.usa.hp.com

		
            	dim Host
		Host = Request.ServerVariables("HTTP_HOST")
		
            	if InStr(Host,"localhost") > 0 then
                	Session("PulsarServerPath") = "http://localhost/Excalibur"
				elseif InStr(Host,lcase(Session("LoggedInUser"))) > 0 then
                	Session("PulsarServerPath") = "http://localhost/Excalibur"
				elseif InStr(Host,"pulsartest") > 0 then
                	Session("PulsarServerPath") = "http://pulsartest.usa.hp.com/Excalibur"
				elseif InStr(Host,"pulsarsandbox") > 0 then
                	Session("PulsarServerPath") = "http://pulsarsandbox.usa.hp.com/Excalibur"
                else
					Session("PulsarServerPath") = "http://pulsarweb.usa.hp.com/Excalibur"
            	end if

	end if

	if (Check_Expire) then
		dim sUser, strEmail, strNTUserDomainID
		if Request.ServerVariables("HTTP_SMUSERDN") <> "" then	'is SiteMinder turned on?
			'SiteMinder
			strNTUserDomainID = Request.ServerVariables("HTTP_NTUSERDOMAINID")
			if strNTUserDomainID <> "" and ucase(left(strNTUserDomainID,4)) <> "AUTH" and ucase(left(strNTUserDomainID,8)) <> "AMERICAS" and _
			  ucase(left(strNTUserDomainID,11)) <> "ASIAPACIFIC" and ucase(left(strNTUserDomainID,4)) <> "EMEA" and _ 
			  ucase(left(strNTUserDomainID,9)) <> "AMERICAS2" and ucase(left(strNTUserDomainID,9)) <> "AMERICAS3" then	'logged in with domain\username
				sUser = Request.ServerVariables("HTTP_NTUSERDOMAINID")
				Session("LoggedInUser") = mid(sUser, Instr(1, sUser, ":") + 1)
				Session("Domain") = left(sUser, Instr(1, sUser, ":")-1)
				Session("DomainNTUserName") = Session("Domain") & "\" & Session("LoggedInUser")
			else	'have to use email to authenticate
				if Request.ServerVariables("HTTP_USERID") <> "" then
					strEmail = Request.ServerVariables("HTTP_USERID")
					Session("Domain") = mid(strEmail, Instr(1, strEmail, "@") + 1)	'after the @ in the email, such as hp.com
					Session("LoggedInUser") = left(strEmail, Instr(1, strEmail, "@")-1)		'before the @ in the email, such as my.name
					Session("DomainNTUserName") = strEmail
				end if
			end if
		else	'no SiteMinder turned on
			sUser = Request.ServerVariables("AUTH_USER")
			Session("LoggedInUser") = mid(sUser, Instr(1, sUser, "\") + 1)
			Session("Domain") = left(sUser, Instr(1, sUser, "\")-1)
			Session("DomainNTUserName") = Session("Domain") & "\" & Session("LoggedInUser")
		end if
		'UpdateLastLogonTime
	end if
end function

function UpdateLastLogonTime()
	dim oServer, oErr, sRepository, sNTUserName, sDomainName
	
	sRepository = Application("Repository")
	sNTUserName = Session("LoggedInUser")
	sDomainName = Session("Domain")
	set oServer = Server.CreateObject("JF_S_General.ISGeneral")
	set oErr = oServer.UpdateLastLogonTime(sRepository, sNTUserName, sDomainName)
	set oErr = nothing 'ignore any error
	set oServer = nothing
end function
%>