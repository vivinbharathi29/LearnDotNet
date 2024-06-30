<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetBulkStatusValue.asp
'* Description	: AMO List - Set Bulk Status Value
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
Dim intAction
Dim strSaveStatus
Dim strModuleIDs
Dim strFullName
Dim intUserID
Dim strRepository
Dim intAMORASReview
Dim intAMORASUpdate
Dim intAMOComplete
Dim intAMOReject
Dim oSvr
Dim oErr
Dim strError

'--- DEFINE LOCAL VARIABLES: ---
'--- querystring values: --
intAction = Request.QueryString("RGS")
strModuleIDs = Request.QueryString("ModuleID")
strFullName = Request.QueryString("FullName")
intUserID = Request.QueryString("UserID")

'---initialize variables: ---
strSaveStatus = ""
intAMORASReview = CInt(Application("AMO_RASREVIEW"))
intAMORASUpdate = CInt(Application("AMO_RASUPDATE"))
intAMOComplete = CInt(Application("AMO_COMPLETE"))
intAMOReject = CInt(Application("AMO_REJECT"))
strRepository = Application("REPOSITORY")

If Not IsNumeric(intAction) Then
	strSaveStatus = "Invalid call to Ajax"
Else
	intAction = CInt(intAction)
    if intAction <> 1 and intAction <> 2 and intAction <> 3 and intAction <> 4 then
        strSaveStatus = "Invalid call to Ajax"
    end if
End If

if strSaveStatus = "" then
	if strModuleIDs = "" then
		strSaveStatus = "Invalid ModuleID"
    else
        strModuleIDs = Trim(strModuleIDs)
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
			case 1	'set AMO RASREVIEW
				strError = oSvr.UpdateBulkStatus(CStr(strRepository), CInt(intAMORASReview), CStr(strModuleIDs), strFullName, intUserID)
			case 2	'set AMO RASUPDATE
				strError = oSvr.UpdateBulkStatus(CStr(strRepository), CInt(intAMORASUpdate), CStr(strModuleIDs), strFullName, intUserID)
            case 3	'set AMO COMPLETE
				strError = oSvr.UpdateBulkStatus(CStr(strRepository), CInt(intAMOComplete), CStr(strModuleIDs), strFullName, intUserID)
            case 4	'set AMO REJECT
				strError = oSvr.UpdateBulkStatus(CStr(strRepository), CInt(intAMOReject), CStr(strModuleIDs), strFullName, intUserID)
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


