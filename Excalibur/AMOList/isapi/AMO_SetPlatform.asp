<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetPlatform.asp
'* Description	: AMO List - Set Platform Value
'* Modified     : Harris, Valerie - 08/31/2016 - Rename file and declare intAction variable
'*************************************************************************************
' --- GLOBAL & OPTIONAL INCLUDES: --- 
%>
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<%
'--- DECLARE LOCAL VARIABLES: ---
Dim strSaveStatus
Dim intAction
Dim intModuleID
Dim intPlatformID
Dim intSetStatus
Dim strValue
Dim strFullName
Dim intUserID
Dim strRepository
Dim oSvr
Dim oErr
Dim strError

'--- DEFINE LOCAL VARIABLES: ---
'--- querystring values: --
intAction = Request.QueryString("RGS")
intModuleID = Request.QueryString("ModuleID")
intPlatformID = Request.QueryString("PlatformID")
intSetStatus = Request.QueryString("SetStatus")
strValue = Request.QueryString("Value")
strFullName = Request.QueryString("FullName")
intUserID = Request.QueryString("UserID")

'---initialize variables: ---
strSaveStatus = ""
strRepository = Application("REPOSITORY")

If Not IsNumeric(intAction) Then
	strSaveStatus = "Invalid call to Ajax"
Else
	intAction = CInt(intAction)
End If

If strSaveStatus = "" Then
	If Not IsNumeric(intModuleID) Then
		strSaveStatus = "Invalid ModuleID"		
	Else
		intModuleID = CLng(intModuleID)
	End If
End If

If strSaveStatus = "" Then
	If Not IsNumeric(intPlatformID) Then
		strSaveStatus = "Invalid PlatformID"
	Else
		intPlatformID = CLng(intPlatformID)
	End If
End If

if strSaveStatus = "" then
	if Not IsNumeric(intSetStatus) then
		strSaveStatus = "Invalid Set Status"
    else
        intSetStatus = CLng(intSetStatus)
	end if
end if

if strSaveStatus = "" then
    'In remote scripting, Value was by default empty so not applying validation
	strValue = request.QueryString("Value")
    if strValue <> "" then
		strValue = Trim(strValue)
    else
        strValue = ""
	end if
end if

If strSaveStatus = "" Then
	If strFullName = "" Then
		strSaveStatus = "Invalid Field"
    Else
        strFullName = Trim(strFullName)
	End If
End If

If strSaveStatus = "" Then
	If intUserID = "" Then
		strSaveStatus = "Invalid UserID"
    Else
        intUserID = Trim(intUserID)
	End If
End If

if strSaveStatus = "" then
	On Error Resume Next
	set oSvr = New ISAMO
	if Err.Number <> 0 then
		strSaveStatus = Err.Description 
	else
		select case intAction
			case 1	'set Platform Save Status              
				strError = oSvr.SavePlatformStatus(CStr(strRepository), intModuleID, intPlatformID, intSetStatus, CStr(strValue), strFullName, intUserID)
		end select

		if Err.number <> 0 then 
			strSaveStatus = Err.Description 
		end if 
		if strError <> "True" then
			strSaveStatus = strError
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end if

if strSaveStatus <> ""  then
	response.write(strSaveStatus)
else
    response.write("success")
end if
response.End
%>
