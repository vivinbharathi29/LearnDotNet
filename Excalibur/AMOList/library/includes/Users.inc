<%
' ========================================================================================
' Include file for User and Group functions
' ========================================================================================

' ========================================================================================
' function Get_Users_Info
' This function creates the session variable recordsets rsGroup and rsUser
' 
' Return
' oErr (object): the returned error object
' ========================================================================================
function Get_Users_Info()
	Get_Users_Info = 0
	'get the user's groups
	Get_Users_Info = Users_CreateRsGroups() 'returns 1 if no records found for teams and/or bus. segments for user
	if Get_Users_Info = 1 then
		'get the user's information
		set Get_Users_Info = Users_Info()
	end if

end function

' ========================================================================================
' function Users_Info
' This function creates the session variable recordset rsUser that contains the user's information
' 
' Return
' oErr (object): the returned error object
' ========================================================================================
Function Users_Info()
	dim bNotFound
	dim oRecords, oRs
	dim CurrentUser, CurrentDomain, CurrentUserName
	
	CurrentUser = lcase(Session("LoggedInUser"))
	if instr(CurrentUser,"\") > 0 then
		CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
		CurrentUserName = mid(CurrentUser,instr(CurrentUser,"\") + 1)
	end if

	bNotFound = False
	set Users_Info = nothing
	
	if Session("AMOUserID") = 0 then
		'need a valid User ID in order to look-up User 
		bNotFound = True
	else
		'set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
		set oRecords = New ISGeneral
		set oRS = oRecords.GetOneUserInfo( Application("Repository"), clng(Session("AMOUserID")))
		if not (oRs is nothing) then
			if oRs.RecordCount = 0 then
				'need a valid User ID in order to look-up User 
				bNotFound = True
			else
				set Session("rsUser") = oRs

				'Get User's ODM Status: --
				'In additional to storing values in Session, save in DB Cookie just in case Session timeouts
				if oRs("IsODMUser").Value = 1 then
					Session("IsODM") = 1	'ODM user
					Call SaveDBCookie("AMO IsODM", 1)
				else
					Session("IsODM") = 0	'HP user
					Call SaveDBCookie("AMO IsODM", 0)
				end if
			end if
		end if
		set oRecords = nothing
	end if
	
	'If User's isn't found in UserInfo Table, create record for the User: ---	
	if bNotFound then
		'can only fill in recordset info from server variables
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Fields.Append "IRSUserID", 3, 4
		oRs.Fields.Append "NTUserName", 129, 30
		oRs.Fields.Append "Domain", 129, 30
		oRs.Fields.Append "FirstName", 129, 30
		oRs.Fields.Append "LastName", 129, 30
		oRs.Fields.Append "Email", 129, 64
		oRs.Fields.Append "Phone", 129, 18
		oRs.Fields.Append "Pager", 129, 18
		oRs.Fields.Append "FullName", 129, 65
		oRs.Fields.Append "NTUserFullName", 129, 100
		oRs.Fields.Append "DomainNTUserName", 129, 65
		oRs.Fields.Append "FirstLast", 129, 65
		oRs.Fields.Append "PFAreaID", 3, 4
		oRs.Fields.Append "PFArea", 129, 64
		oRs.Fields.Append "ShowFullMenu", 11
		oRs.Fields.Append "IsODMUser", 11
		oRs.CursorLocation = 2
		oRs.Open

		oRs.AddNew
		oRs("IRSUserID").Value = 0
		oRs("NTUserName").Value = CurrentUserName
		oRs("Domain").Value = CurrentDomain
		oRs("FirstName").Value = ""
		oRs("LastName").Value = ""
		oRs("Email").Value = ""
		oRs("Phone").Value = ""
		oRs("Pager").Value = ""
		oRs("FullName").Value = ""
		oRs("NTUserFullName").Value = ""
		oRs("DomainNTUserName").Value = CurrentUser
		oRs("FirstLast").Value = ""
		oRs("PFAreaID").Value = 0
		oRs("PFArea").Value = ""
		oRs("ShowFullMenu").Value = 0
		oRs("IsODMUser").Value = 1	'default to say they are an ODM user so they don't have many rights
		oRs.MoveNext

		set session("rsUser") = oRs
		session("BusinessSegments") = ""
		session("FullName") = ""
		session("IsODM") = 1
		Call SaveDBCookie("AMO IsODM", 1)
	end if
end function

' ========================================================================================
' function Users_CreateRsGroups
' This function creates the session variable recordset rsGroup that contains all
' the group names the user belongs to.
' Also creates the session variables ODMIDs, ODMs, SiteIDs, BusinessSegmentIDs, BusinessSegments, DASHNetIDs
' 
' Return
' oErr (object): the returned error object
' ========================================================================================
function Users_CreateRsGroups()
	dim oSvr, oRs, oErr, bErr
	dim ODMIDs, ODMs, SiteIDs, BusinessSegmentIDs, BusinessSegments, DASHNetIDs
	ODMIDs = ""
	ODMs = ""
	SiteIDs = ""
	BusinessSegmentIDs = ""
	BusinessSegments = ""
	DASHNetIDs = ""
	bErr = 0

	'set oSvr = Server.CreateObject("JF_S_Permission.ISUser")
	set oSvr = New ISGeneral
	if Session("AMOUserID") = "" then
		Session("AMOUserID") = 0
	end if
		
	set oRs = oSvr.GetGroupsByUser( Application("Repository"), clng(Session("AMOUserID")))
	if oRs.RecordCount = 0 then
		bErr = 1
	else
		if oRs.RecordCount > 0 then
			set Session("rsAMOGroups") = oRs
			session("FullName") = oRs.Fields("FullName")
			Call SaveDBCookie("AMO FullName", oRs.Fields("FullName"))
		end if
	end if

	if bErr = 0 then
		set oSvr = New ISWebCategory
		'Get Team-to-BusinessSegment Name and IDs for Current User, replaces GrouptoDivision
		set oRs = oSvr.wUser_BusinessSegment(Application("Repository"), clng(Session("AMOUserID"))) 
		if oRs is nothing then
			bErr = 1
		else
			Do While Not oRs.EOF
				If IsNull(oRs("Disabled")) = True Or IsEmpty(oRs("Disabled")) = True Or oRs("Disabled") = "" Then
					if BusinessSegmentIDs <> "" then
						BusinessSegmentIDs = BusinessSegmentIDs + ", "
					end if
					if BusinessSegments <> "" then
						BusinessSegments = BusinessSegments + ", "
					end if
					
					BusinessSegmentIDs = BusinessSegmentIDs & cstr(oRs("BusinessSegmentID"))
					BusinessSegments = BusinessSegments & oRs("BusinessSegment")
				End If
				oRs.MoveNext
			loop
			session("BusinessSegmentIDs") = BusinessSegmentIDs
			Call SaveDBCookie("AMO BusinessSegmentIDs", BusinessSegmentIDs)
			session("BusinessSegments") = BusinessSegments	'can not save bus. segment names to Cookie, characters too long			
		end if
	end if

	Users_CreateRsGroups = bErr
	set oSvr = nothing
end function

' ========================================================================================
' function Users_CreateUserInfo
' Creates nt user name & domain followed by all groups or logon group for user.
' 
' Input:
' bLogon - True if for logon or False if for all groups
'
' Returns:
' string to display
' ========================================================================================
function Users_CreateUserInfo( bLogon )
	dim sHTML, sUsersGroups
	sHTML = ""
	
	if session("Domain") = "" then
		'bypass all because the rest is not implemented yet
	else
		if session("Domain") <> "" or session("User") <> "" then
			sHTML = sHTML & "User = "
			if session("Domain") <> "" then
				sHTML = sHTML & session("Domain")
				if session("User") <> "" then
					sHTML = sHTML & "\"
				end if
			end if
			if session("User") <> "" then
				sHTML = sHTML & session("User")
			end if
		end if
		if sHTML <> "" then
			sHTML = sHTML & ", "
		end if

		if bLogon then
			sHTML = sHTML & Users_DisplayLogonGroup
		else
			sUsersGroups = Users_DisplayAllGroups
			sHTML = sHTML & Users_DisplayLogonGroup
			if sUsersGroups <> "" then
				if right(sHTML, 2) = ", " then
					sHTML = sHTML & sUsersGroups
				else
					sHTML = sHTML & ", " & sUsersGroups
				end if
			end if
		end if
		
		if right(sHTML, 2) = ", " then
			sHTML = left(sHTML, len(sHTML)-2)
		end if
	end if
	
	Users_CreateUserInfo = sHTML
end function

' ========================================================================================
' function Users_DisplayAllGroups
' Displays all groups for user
' 
' These are the rules:
' - if session("group") is not "" then display that as the logon group at the top of the home page like it is today
' - if session("group") = "" and Session("AMOUserID") <> "" then that means it's a non-guest so just display their username with no logon group at the top of the home page
' - if session("group") = "" and Session("AMOUserID") = 0 then that means it's a guest so display their username along with "Guest/View Only" at the top of the home page
' - in the global footer, display ALL the groups the user is in. This is true for the home page too whereas today that information is not displayed in the global footer if you're on the home page.
'
' Returns:
' string of HTML to display
' ========================================================================================
function Users_DisplayAllGroups()
	dim sHTML
	dim oRs
	
	sHTML = ""
	if CLng(Session("AMOUserID")) = 0 then
		'don't display this because it is being taken care of in Users_DisplayLogonGroup
		'sHTML = "User Group = Guest/View Only"
	elseif not isobject(Session("rsAMOGroups")) then
		'recordset doesn't exist so don't display anything
	else
		set oRs = Session("rsAMOGroups")
		if oRs.State <> 1 then 
			'is the recordset opened?
		else
			oRs.Filter = ""
			oRs.Movefirst
			'get rid of IRS Admin group if there
			'oRs.Filter = "GroupType<>4"
			if oRs.RecordCount > 0 then
				oRs.Sort = "GroupName"
				sHTML = "User Groups = "
				while not oRs.eof
					sHTML = sHTML & oRs.Fields("GroupName").Value & ", "
					oRs.MoveNext
				wend
				if right(sHTML, 2) = ", " then
					sHTML = left(sHTML, len(sHTML)-2)
				end if
			end if
		end if
		set oRs = nothing
		
	end if

	Users_DisplayAllGroups = sHTML
end function

' ========================================================================================
' function Users_DisplayLogonGroup
' Displays logon group for user
' 
' These are the rules:
' - if session("group") is not "" then display that as the logon group at the top of the home page like it is today
' - if session("group") = "" and Session("AMOUserID") <> "" then that means it's a non-guest so just display their username with no logon group at the top of the home page
' - if session("group") = "" and Session("AMOUserID") = 0 then that means it's a guest so display their username along with "Guest/View Only" at the top of the home page
' - in the global footer, display ALL the groups the user is in. This is true for the home page too whereas today that information is not displayed in the global footer if you're on the home page.
'
' Returns:
' string of HTML to display
' ========================================================================================
function Users_DisplayLogonGroup()
	dim sHTML
	dim oRs
	
	if CLng(Session("AMOUserID")) = 0 then
		sHTML = "Logon User Group = Guest/View Only"
	elseif session("group") <> "" then
		sHTML = "Logon User Group = " & session("Group")
	else
		sHTML = ""
	end if

	Users_DisplayLogonGroup = sHTML
end function

' ========================================================================================
' function IsUserInGroup(lngGroupID)
' This function returns true if the current user is in the input group or 
' IRS admin, and false otherwise.
' ========================================================================================
function IsUserInGroup(byval lngGroupID)
	dim bAdminCreate, bAdminUpdate, bAdminView, bAdminDelete
	dim rs

	IsUserInGroup = false
	GetRights2 Application("ADMIN"), bAdminCreate, bAdminView, bAdminUpdate, bAdminDelete

	if bAdminCreate or bAdminUpdate or bAdminDelete then
		IsUserInGroup = true
	else
		if isobject(Session("rsAMOGroups")) and not IsNUll(lngGroupID) then
			set rs = Session("rsAMOGroups")
			rs.Filter = "GroupID=" & cstr(lngGroupID)
			if rs.RecordCount > 0 then
				IsUserInGroup = true
			end if
			rs.Filter = ""
		end if
	end if
end function

' ========================================================================================
' function IsUserInGroupType(lngGroupType)
' This function returns true if the current user is in the specified User Group type
' and false otherwise.
' ========================================================================================
function IsUserInGroupType(byval lngGroupType)
	dim rs
	
	IsUserInGroupType = false
	if isobject(Session("rsAMOGroups")) then
		set rs = Session("rsAMOGroups")
		rs.Filter = "GroupType=" & cstr(lngGroupType)
		if rs.RecordCount > 0 then
			IsUserInGroupType = true
		end if
		rs.Filter = ""
	end if
end function

' ========================================================================================
' function GetUsersInfoFromLDAP(strError)
' This function returns true if the current user is in GAL and false otherwise.
' Session variables for user's information is created.
' If an error occurs, the error message will be returned in the strError variable.
' ========================================================================================
function GetUsersInfoFromLDAP(byRef strError)
	dim strFirstName, strLastName, strEmail, strPhone, strPager
	dim oRecords, objErr

	strError = ""
	set oRecords = Server.CreateObject("JF_S_Permission.ISUser")
	set objErr = oRecords.GetUserInfoFromADByNT( Session("Domain") & ":" & Session("User"), _
			strFirstName, strLastName, strEmail, strPhone, strPager )
	if not objErr is nothing then
		'fyi, LDAP goes down occasionally
		strError = Errors_GetHTML(objErr)
		if strError = "" then
			strError = "LDAP information not found"
		end if
		'didn't find user
		GetUsersInfoFromLDAP = False
	else
		if (strFirstName = "") or (strLastName = "") or (strEmail = "") then
			GetUsersInfoFromLDAP = False
		else
			'create temporary sessions since they will be abandoned at the end
			session("chk-firstname") = strFirstName
			session("chk-lastname") = strLastName
			session("chk-email") = strEmail
			session("chk-phone") = strPhone
			session("chk-pager") = strPager
			GetUsersInfoFromLDAP = True
		end if
	end if
	set objErr = nothing
	set oRecords = nothing

end function

'********************************************************************************
'* Purpose: Get roles of user in ISIS user group and return recordset of ISIS user group
'* 
'* Parameters:
'*  oRsISISUsers  = recordset of all ISIS users and their respective NT Groups
'*  bISIS         = if the user is in the ISIS user group
'*  bIRSDeveloper = if the user is an IRS Developer
'*  bIRSTester    = if the user is an IRS Tester
'*  bIRSPM        = if the user is an IRS PM
'*  bIRSRelease   = if the user is an IRS Release Coordinator
'*
'* Returns:
'*  Fills in all passed variables and returns any error
'********************************************************************************
function GetISISRoles( byref oRsISISUsers, byref bISIS, byref bIRSDeveloper, byref bIRSTester, byref bIRSPM, byref bIRSRelease )
	dim oRecords, oErr
	dim sErr
	
	sErr = ""
	set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
	'get all the users in ISIS (groupid=100) and their respective NT Groups
	set oErr = oRecords.ViewUsersByGroup (Application("Repository"), 100, oRsISISUsers )
	if not oErr is nothing then
		sErr = Errors_GetHTML(oErr)
	else
		bISIS = false
		bIRSDeveloper = false
		bIRSTester = false
		bIRSPM = false
		bIRSRelease = false

		oRsISISUsers.Filter = "IRSUserID=" & Session("AMOUserID")
		if oRsISISUsers.RecordCount > 0 then
			bISIS = true
		end if
		oRsISISUsers.Filter = ""
		
		oRsISISUsers.Filter = "NTGroupName='IRS Developer' and IRSUserID=" & Session("AMOUserID")
		if oRsISISUsers.RecordCount > 0 then
			bIRSDeveloper = true
		end if
		oRsISISUsers.Filter = ""
		
		oRsISISUsers.Filter = "NTGroupName='IRS Tester' and IRSUserID=" & Session("AMOUserID")
		if oRsISISUsers.RecordCount > 0 then
			bIRSTester = true
		end if
		oRsISISUsers.Filter = ""
		
		oRsISISUsers.Filter = "NTGroupName='IRS PM' and IRSUserID=" & Session("AMOUserID")
		if oRsISISUsers.RecordCount > 0 then
			bIRSPM = true
		end if
		oRsISISUsers.Filter = ""
		
		oRsISISUsers.Filter = "NTGroupName='IRS Release Coordinator' and IRSUserID=" & Session("AMOUserID")
		if oRsISISUsers.RecordCount > 0 then
			bIRSRelease = true
		end if
		oRsISISUsers.Filter = ""
	end if

	set oRecords = nothing
	
	GetISISRoles = sErr
end function

'********************************************************************************
'* Purpose: Get the User GroupIDs for a given user for the given RoleType
'* 
'* Parameters:
'*  intIRSUserID         = IRSUserID of user
'*  intRoleCode       = RoleCode to search. This would be an application variable from global.asa such as Application("TPlan")
'*  chrGroupTypes     = Optional way to filter the user group types searched. Enter an empty string if not wanting to filter.
'*  chrSearchRoleName = The role name to filter, i.e. create, update, delete, view
'*
'* Returns:
'*  String of comma delimited User GroupIDs of which the given user has the given RoleType
'*  oRsTestGroups = Entire recordset retrieved in case user wants to do more manipulation on it
'* NOT USED IN PULSAR
'********************************************************************************
function Users_ViewUserRoleCodes( byval intIRSUserID, byval intRoleCode, byval chrGroupTypes, byval chrSearchRoleName, byref oRsTestGroups)
	dim oRecords, oErr, oRs
	dim strGroupIDs
	
	strGroupIDs = ""
	set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
	set oErr = oRecords.ViewUserRoleCodes (Application("Repository"), intIRSUserID, intRoleCode, oRs, chrGroupTypes)
	if not oErr is nothing then
		HandleError(oErr)
	else
		if chrSearchRoleName <> "" then
			oRs.Filter = "ShortRoleName = '" & chrSearchRoleName & "'"
		end if
		if oRs.RecordCount > 0 then
			do while not oRs.eof
				if strGroupIDs <> "" then
					strGroupIDs = strGroupIDs & ","
				end if
				strGroupIDs = strGroupIDs & cstr(oRs("UserGroupID"))
				oRs.movenext
			loop
		end if
		oRs.Filter = ""
		set oRsTestGroups = oRs
		set oRs = nothing
	end if
	
	set oRecords = nothing
	
	Users_ViewUserRoleCodes = strGroupIDs
end function

'********************************************************************************
'* Purpose: Get the User GroupIDs for the current user filtered by group type
'* 
'* Parameters:
'*  chrGroupTypes = filter the user group types searched. Enter an empty string if not wanting to filter.
'*
'* Returns:
'*  String of comma delimited User GroupIDs for the group type
'********************************************************************************
function Users_UserGroupIDs( byval chrGroupTypes)
	dim strGroupIDs, strGroupID, strFilter
	dim oRsGroups
	dim arrGroupTypes
	
	strGroupIDs = ""
	set oRsGroups = Session("rsAMOGroups")
	
	if chrGroupTypes <> "" then
		arrGroupTypes = split(chrGroupTypes, ",")
		strFilter = ""
		for each strGroupID in arrGroupTypes
			if strFilter <> "" then
				strFilter = strFilter & " or "
			end if
			strFilter = strFilter & "GroupType=" & strGroupID
		next

		oRsGroups.filter = strFilter
	end if

	if oRsGroups.RecordCount > 0 then
		do while not oRsGroups.eof
			if strGroupIDs <> "" then
				strGroupIDs = strGroupIDs & ","
			end if
			strGroupIDs = strGroupIDs & cstr(oRsGroups("GroupID"))
			oRsGroups.movenext
		loop
		oRsGroups.movefirst
	end if
	oRsGroups.Filter = ""
	set oRsGroups = nothing

	Users_UserGroupIDs = strGroupIDs
end function
%>