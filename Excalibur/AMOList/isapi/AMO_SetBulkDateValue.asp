<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetBulkDateValue.asp
'* Description	: AMO List - Set Date Value for multiple AMO Features
'*************************************************************************************
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
Dim strModuleRegionIDs
Dim strCPLBlindDate
Dim strBOMRevADate
Dim strRasDisconDate
Dim strObsoleteDate
Dim strGlobalseriesdate
Dim strFullName
Dim intUserID
Dim strRepository
Dim oSvr
Dim oErr
Dim strError

'--- DEFINE LOCAL VARIABLES: ---
'--- querystring values: --
intAction = Request.QueryString("RGS")
strModuleRegionIDs = Request.QueryString("ModuleRegionID")
strCPLBlindDate = Request.QueryString("CPLBlindDate")
strBOMRevADate = Request.QueryString("BOMRevADate")
strRasDisconDate = Request.QueryString("RasDisconDate")
strObsoleteDate = Request.QueryString("ObsoleteDate")
strGlobalseriesdate = Request.QueryString("Globalseriesdate")
intUserID = Request.QueryString("UserID")

'---initialize variables: ---
strSaveStatus = ""
strRepository = Application("REPOSITORY")


If Not IsNumeric(intAction) Then
	strSaveStatus = "Invalid call to Ajax"
Else
	intAction = CInt(intAction)
End If

if strSaveStatus = "" then
	if strModuleRegionIDs = "" then
		strSaveStatus = "Invalid Module Region ID"
    Else
        strModuleRegionIDs = Trim(strModuleRegionsIDs) 
	end if
end if

if strSaveStatus = "" then	
	if strCPLBlindDate = "" then
		strSaveStatus = "Invalid CPLBlindDate"
	else
		if not IsDate(strCPLBlindDate) then
			strSaveStatus = "Invalid CPLBlindDate"
		end if
	end if
end if

if strSaveStatus = "" then	
	if strBOMRevADate = "" then
		strSaveStatus = "Invalid BOMRevADate"
	else
		if not IsDate(strBOMRevADate) then
			strSaveStatus = "Invalid BOMRevADate"
		end if
	end if
end if


if strSaveStatus = "" then
	if strRasDisconDate = "" then
		strSaveStatus = "Invalid RasDisconDate"
	else
		if not IsDate(strRasDisconDate) then
			strSaveStatus = "Invalid RasDisconDate"
		end if
	end if
end if

if strSaveStatus = "" then
	if strObsoleteDate = "" then
		strSaveStatus = "Invalid ObsoleteDate"
	else
		if not IsDate(strObsoleteDate) then
			strSaveStatus = "Invalid ObsoleteDate"
		end if
	end if
end if

if strSaveStatus = "" then
	if strGlobalseriesdate = "" then
		strSaveStatus = "Invalid Globalseriesdate"
	else
		if not IsDate(strGlobalseriesdate) then
			strSaveStatus = "Invalid Globalseriesdate"
		end if
	end if
end if

If strSaveStatus = "" Then
	If strFullName = "" Then
		strSaveStatus = "Invalid Field"
    Else
        strFullName = CStr(strFullName)
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
			case 1	
				strError = oSvr.UpdateBulkDate(strRepository, strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate, strFullName, intUserID)
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
