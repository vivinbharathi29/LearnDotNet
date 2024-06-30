<!-- #include file="Users.inc" -->
<%
' ========================================================================================
' Script to get roles
' ========================================================================================
' GetRoles		: get the roles for component/test/project management
' ========================================================================================

' ========================================================================================
' function GetAllRoles2()
' This function caches the role permissions for the current log on user.
'
' Session("rsAMORoles"): the recordset that contains the roles
'			Fields in the recordset: 
'			Type (long): beantype
'			Create, View, Update, Delete (boolean): the permission
' Return
' oErr (object): the returned error object
' ========================================================================================
' function GetAllRoles2()
' This function returns the role permissions for objects in IRS system.
' Currently used in module marketing implementation. The rest of IRS is not ready to use this yet.
'
' Return
' oErr (object): the returned error object
' ========================================================================================
if (Session("AMOUserID") = "") then
	RestoreSession()
end if

'*************************************************************************************
'* Function		: GetUserInfo
'* Description	: Get Current User's ID or Email Address
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************
Function GetUserInfo(sType)
	Dim oPermission
	Dim oRSUser, oRSUserImpersonate, sRepository, CurrentUser, CurrentUserID, CurrentUserEmail, ImpersonateUserID

	'Set Repository value: ---
	sRepository = Application("Repository")

	set oRSUser = nothing
	set oRSUserImpersonate = nothing
	
	'Get Actual User Network login name: ---
	CurrentUser = lcase(Session("LoggedInUser"))
	If Session("LoggedInUser") <> "" Then
		CurrentUser = lcase(Session("LoggedInUser"))
		Call SaveDBCookie("AMO LoggedInUser", CurrentUser)
	Else
		CurrentUser = GetDBCookie("AMO LoggedInUser")
	End If
	
	'Set Permissions Object: ---
	set oPermission = New ISRole
	
	'Get Actual User's Email and User ID: ---	
	set oRSUser = oPermission.GetActualUserInfo(sRepository, CurrentUser)
	if not (oRSUser is Nothing) then
		if not (oRSUser.EOF) then
			CurrentUserID = oRSUser("ID") 
			CurrentUserEmail = oRSUser("Email") 
		else
			CurrentUserID = 0
			CurrentUserEmail = ""
		end if
	end if

	'Get Impersoante User's User ID: ---	
	set oRSUserImpersonate = oPermission.GetImpersonateUser(sRepository, CurrentUser)
	if not (oRSUserImpersonate is Nothing) then
		if not (oRSUserImpersonate.EOF) then
			ImpersonateUserID = oRSUserImpersonate("ImpersonateID") 
		else
			ImpersonateUserID = 0
		end if
	else
		ImpersonateUserID = 0
	end if

	'Close Permission Object: --
	set oPermission = nothing

	If sType = "Email" then
		GetUserInfo = CurrentUserEmail
		Exit Function
	Else
		If ImpersonateUserID <> 0 Then
			Session("AMOUserID") = ""
			GetUserInfo = ImpersonateUserID
			Session("AMOUserID") = ImpersonateUserID
		Else
			Session("AMOUserID") = ""
			GetUserInfo = CurrentUserID
			Session("AMOUserID") = CurrentUserID
		End If
		Exit Function
	End If
End Function

'*************************************************************************************
'* Function		: SetPermission (modified version of IRS' GetAllRoles2)
'* Description	: Return Current User's Roles/Permissions or ID
'* Modified		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************
sub SetPermission()
	dim oPermission, oRS, sRepository, intUserID, CurrentUser, sUserRoleNames
	
	set Session("rsAMORoles") = Nothing

	'Set Repository value: ---
	sRepository = Application("Repository")

	'Get Actual Current User
	CurrentUser = lcase(Session("LoggedInUser"))

	'Get either Actual or Impersonate User ID
	intUserID = clng(GetUserInfo("UserID"))

	'Set Role Names Session: ---
	if Session("AMOUserRoleNames") = "" then
		set oPermission = New ISRole
		sUserRoleNames = oPermission.GetUserRoleNames(sRepository, intUserID)
		Session("AMOUserRoleNames") = sUserRoleNames 'can not save user's role names to Cookie, characters too long for users assigned multiple roles	
		set oPermission = nothing
	end if

	'Set UserID Session: ----
	if IsNumeric(intUserID) = True then
		Session("AMOUserID") = intUserID
		Call SaveDBCookie("AMO AMOUserID", intUserID)
	else
		Session("AMOUserID") = 0
		Call SaveDBCookie("AMO AMOUserID", 0)
	end if

	'Set User's Info
	Get_Users_Info()

	'Set ODM 
	ODM_Users_Info()
end sub

' ========================================================================================
' sub ODM_Users_Info
' This sub creates the session and cookie for ODM Validation
' 
' Return
' oErr (object): the returned error object
' ========================================================================================
sub ODM_Users_Info()
	dim bNotFound
	dim oRecords, oODMRs, intUserID

	bNotFound = False
	set oODMRs = nothing

	If Session("AMOUserID") = "" Then
		intUserID = GetDBCookie("AMO AMOUserID")
		Session("AMOUserID") = intUserID
	Else
		intUserID = Session("AMOUserID")
	End If

	set oRecords = New ISGeneral
	set oODMRs = oRecords.GetOneUserInfo( Application("Repository"), clng(Session("AMOUserID")))
	if not (oODMRs is nothing) then
		if oODMRs.RecordCount = 0 then
			'need a valid User ID in order to look-up User 
			bNotFound = True
		else
			'Get User's ODM Status: --
			'In additional to storing values in Session, save in DB Cookie just in case Session timeouts
			if cint(oODMRs("IsODMUser").Value) = 1  then
				Session("IsODM") = 1	'ODM user
				Call SaveDBCookie("AMO IsODM", 1)
			else
				Session("IsODM") = 0	'HP user
				Call SaveDBCookie("AMO IsODM", 0)
			end if
		end if
	end if
	set oRecords = nothing	
end sub



' ========================================================================================
' function GetRights(byval lngBeanType, byref oRS, byref bCreate, byref bView, byref bUpdate, byref bDelete)
' This function returns the role permissions for a bean type.
'
' lngBeanType (number): the bean type to get permissions
' oRS: the recordset that contains the roles
' Output:
' bCreate, bView, bUpdate, bDelete: the permissions
' ========================================================================================
function GetRights(byval lngBeanType, byref oRS, byref bCreate, byref bView, byref bUpdate, byref bDelete)
	bCreate = false
	bView = false
	bUpdate = false
	bDelete = false

	if oRS is nothing then exit function	'is the recordset defined?
	if oRS.State <> 1 then exit function	'is the recordset opened?
	oRS.Filter = ""
	oRS.Movefirst
	oRS.Filter = "RoleName=" & lngBeanType	'find the bean type
	if oRS.RecordCount > 0 then				'found -> get the permissions
		bCreate = oRS.Fields("Create")
		bView = oRS.Fields("View")
		bUpdate = oRS.Fields("Update")
		bDelete = oRS.Fields("Delete")
	end if
	oRS.Filter = ""
end function

' ========================================================================================
' function GetRights2(byval lngBeanType, byref bCreate, byref bView, byref bUpdate, byref bDelete)
' This function returns the role permissions for a specific permission.
' Knows that there is only one recordset for the roles
'
' strPermissionType (string): The permission assigned to the role
' Output:
' bCreate, bView, bUpdate, bDelete: the permissions
' ========================================================================================
function GetRights2(byval strPermissionType, byref bCreate, byref bView, byref bUpdate, byref bDelete)
	'set oRS = Session("rsAMORoles"); lcase(Session("LoggedInUser"))
	
	dim oRS, oPermission, sRepository, intUserID, CurrentUser

	'Set Repository value: ---
	sRepository = Application("Repository")

	'Get User
	CurrentUser = lcase(GetUserInfo("Email"))

	'Get Permission Status for Role assigned specific Permission: ---	
	set oPermission = New ISRole
	set oRS = oPermission.GetAllRoles(sRepository, CurrentUser, cstr(strPermissionType))

	bCreate = false
	bView = false
	bUpdate = false
	bDelete = false

	if oRS is nothing then exit function	'is the recordset defined?
	if oRS.State <> 1 then 
		'is the recordset opened?
	else
		oRS.Filter = ""
		oRS.Movefirst
		'oRS.Filter = "View = 1 Or Update = 1 or Create = 1 Or Delete = 1 " 	'find the Role where Permission is 
		oRs.Filter = "View = 1"
		if oRS.RecordCount > 0 then											'found -> get the permissions
			bView = True 'oRS.Fields("View")
		end if	

		oRs.Filter = "Update = 1"
		if oRS.RecordCount > 0 then											'found -> get the permissions
			bUpdate = True 'oRS.Fields("Update")
		end if	

		oRs.Filter = "Create = 1"
		if oRS.RecordCount > 0 then											'found -> get the permissions
			bCreate = True 'oRS.Fields("Create")
		end if	

		oRs.Filter = "Delete = 1"
		if oRS.RecordCount > 0 then											'found -> get the permissions
			bDelete = True 'oRS.Fields("Delete")
		end if	

		'if oRS.RecordCount > 0 then											'found -> get the permissions
		'	bCreate = oRS.Fields("Create")
		'	bView = oRS.Fields("View")
		'	bUpdate = oRS.Fields("Update")
		'	bDelete = oRS.Fields("Delete")
		'end if

		oRS.Filter = ""
	end if
	set oRS = nothing
	
end function

' ========================================================================================
' function GetRightsByRecordset
' This function returns the role permissions for a bean type using the given recordset.
' The recordset must be set up like Session("rsAMORoles"). One way to get a recordset for the
' user filtered by division is to use the function GetAllRolesByDivision.
' Input:
' 	lngBeanType (number) = the bean type to get permissions
' 	oRS                  = Recordset that contains the rights for the user.
' Output:
' bCreate, bView, bUpdate, bDelete: the permissions
' ========================================================================================
function GetRightsByRecordset(byval lngBeanType, byref bCreate, byref bView, byref bUpdate, byref bDelete, byval oRS)
	bCreate = false
	bView = false
	bUpdate = false
	bDelete = false

	if oRS is nothing then exit function			'is the recordset defined?
	if oRS.State <> 1 then 
		'is the recordset opened?
	else
		oRS.Filter = ""
		oRS.Movefirst
		oRS.Filter = "RoleName=" & lngBeanType		'find the bean type
		if oRS.RecordCount > 0 then					'found -> get the permissions
			bCreate = oRS.Fields("Create")
			bView = oRS.Fields("View")
			bUpdate = oRS.Fields("Update")
			bDelete = oRS.Fields("Delete")
		end if
		oRS.Filter = ""
	end if
end function

' ========================================================================================
' function GetAllRolesByDivision
' This function returns the role permissions for objects in IRS system for a given user in a given division.
' Input:
'		lngUserID      = UserID of user. Can use session("UserID")
' 	strDivisionIDs = comma delimited business segment ID's
' Return:
' 	oRS = recordset set up just like Session("rsAMORoles")
' 	oErr (object): the returned error object
' ========================================================================================
function GetAllRolesByDivision(byval lngUserID, byval strDivisionIDs, byref oRS)
	dim oPermission
	
	'set oPermission = Server.CreateObject("JF_S_Permission.ISRole")
	set oPermission = New ISRole
	set GetAllRolesByDivision = oPermission.GetAllRolesByDivision(Application("Repository"), lngUserID, strDivisionIDs)
	set oPermission = nothing
end function

' ========================================================================================
' function GetAllRolesByUserGroup
' This function returns the role permissions for objects in IRS system for a given user in a given division.
' Input:
'		lngUserID       = UserID of user. Can use session("UserID")
' 	strUserGroupIDs = comma delimited Usergroup ID's
' Return:
' 	oRS = recordset set up just like Session("rsAMORoles")
' 	oErr (object): the returned error object
' ========================================================================================
function GetAllRolesByUserGroup(byval lngUserID, byval strUserGroupIDs, byref oRS)
	dim oPermission
	
	'set oPermission = Server.CreateObject("JF_S_Permission.ISRole")
	set oPermission = New ISRole
	set GetAllRolesByUserGroup = oPermission.GetAllRolesByUserGroup(Application("Repository"), lngUserID, strUserGroupIDs)
	set oPermission = nothing
end function

' ========================================================================================
' original function
function GetGroupsForRole(byval sRoleName, byval sPermissionName, byval bCreate, byval bView, byval bUpdate, byval bDelete, byval bExcludeIRS)
	set GetGroupsForRole = GetGroupsForRole_Base(sRoleName, sPermissionName, bCreate, bView, bUpdate, bDelete, bExcludeIRS, true)
end function

' same as original function but do not check Admin rights. That way only the user's groups for the particular role will be correct.
function GetGroupsForRole2(byval sRoleName, byval sPermissionName, byval bCreate, byval bView, byval bUpdate, byval bDelete, byval bExcludeIRS, byval bCheckAdmin)
	set GetGroupsForRole2 = GetGroupsForRole_Base(sRoleName, sPermissionName, bCreate, bView, bUpdate, bDelete, bExcludeIRS, bCheckAdmin)
end function

function GetGroupsForRole_Base(byval sRoleName, byval sPermissionName, byval bCreate, byval bView, byval bUpdate, byval bDelete, byval bExcludeIRS, byval bCheckAdmin)
	dim oSvr, objErr, sDomainUserName, oRS, intUserID 
	dim bAdminCreate, bAdminUpdate, bAdminView, bAdminDelete

	'sDomainUserName = Session("DomainNTUserName")
	sDomainUserName = lcase(Session("LoggedInUser"))
	intUserID = clng(GetUserInfo("UserID"))

	if bCheckAdmin then
		GetRights2 Application("ADMIN"), bAdminCreate, bAdminView, bAdminUpdate, bAdminDelete
		if bAdminCreate or bAdminUpdate then
			'if admin, do not get groups for a specific user, just get all the groups under the role
			sDomainUserName = ""
		end if
	end if
	'set oSvr = Server.CreateObject("JF_S_Permission.ISRole")
	set oSvr = New ISRole
	set oRS = oSvr.GetRoleGroups(Application("Repository"), sRoleName, sPermissionName, bCreate, bUpdate, bView, bDelete, intUserID )

	if (oRS is nothing) then
		set GetGroupsForRole_Base = Nothing
	else
		if not (oRS is nothing) then
			'if bExcludeIRS then
				'oRS.Filter = "GroupCode <> 'IRS'"
			'end if
			if (oRS.RecordCount > 1) then
				oRS.Sort = "GroupName"
			end if
		end if
		set GetGroupsForRole_Base = oRS
	end if
	set oSvr = nothing
end function



%>