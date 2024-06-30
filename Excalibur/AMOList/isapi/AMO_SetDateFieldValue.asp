<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetDateFieldValue.asp
'* Description	: AMO List - Set Date Field Value
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
Dim intRegionID
Dim strField
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
intRegionID = Request.QueryString("RegionID")
strField = Request.QueryString("Field")
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
	If Not IsNumeric(intRegionID) Then
		strSaveStatus = "Invalid RegionID"
	Else
		intRegionID = CLng(intRegionID)
	End If
End If

If strSaveStatus = "" Then
	If strField = "" Then
		strSaveStatus = "Invalid Field"
    Else
        strField = Trim(strField)
	End If
End If

If strSaveStatus = "" Then
	If strValue = "" Then
		strSaveStatus = "Invalid Value"
	Else
        'fix problem with passing in + characters
		strValue = Replace(strValue,"%2B","+")
        strValue = Trim(strValue)
	End If
End If

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

If strSaveStatus = "" Then
	On Error Resume Next
	set oSvr = New ISAMO
	If Err.Number <> 0 Then
		strSaveStatus = Err.Description 
	Else
	
		select case intAction 
			case 1	
				strError = oSvr.SaveDateFieldValue(CStr(strRepository), intModuleID, intRegionID, strField, CStr(strValue), strFullName, intUserID)
		End select

		if Err.number <> 0 then 
			strSaveStatus = Err.Description 
		end if 
		if strError <> "True" then
			strSaveStatus = strError
		end if
	End If 
	set oErr = nothing
	set oSvr = nothing
End If

If strSaveStatus <> ""  Then
	response.write(strSaveStatus)
Else
    response.write("success")
End If
response.End
%>

