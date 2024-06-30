<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetComparability.asp
'* Description	: AMO List - Set Comparability Value
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
Dim intModuleID 
Dim intDivisionID
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
intDivisionID = Request.QueryString("DivisionID")
intSetStatus = Request.QueryString("SetStatus")
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
	If Not IsNumeric(intDivisionID) Then
		strSaveStatus = "Invalid DivisionID"
	Else
		intDivisionID = CLng(intDivisionID)
	End If
End If

if strSaveStatus = "" then
	if Not IsNumeric(intSetStatus) then
		strSaveStatus = "Invalid Set Status"
    else
        intSetStatus = CLng(intSetStatus)
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
'	On Error Resume Next
	set oSvr = New ISAMO
	if Err.Number <> 0 then
		strSaveStatus = Err.Description 
	else
		select case intAction
			case 1	'set Comparability          
				strError = oSvr.SaveComparabilityStatus(CStr(strRepository), intModuleID, intDivisionID, intSetStatus, strFullName, intUserID)
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
