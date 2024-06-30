<%@ LANGUAGE="VBScript" %>
<%OPTION EXPLICIT
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SetBulkStatusEmailValue.asp
'* Description	: AMO List - Send Status Email for multiple Modules
'*************************************************************************************
' --- GLOBAL & OPTIONAL INCLUDES: --- 
%>
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../includes/AMO.inc" -->
<!-- #include file="../data/openDBConnection.asp" -->
<%
'--- DECLARE LOCAL VARIABLES: ---
Dim strSaveStatus
Dim intAction
Dim intStatusID
Dim strModuleIDs
Dim strRegionIDs
Dim oSvr
Dim oErr

'--- DEFINE LOCAL VARIABLES: ---
'--- querystring values: --
intAction = Request.QueryString("RGS")
strModuleIDs = Request.QueryString("ModuleID")
strRegionIDs = Request.QueryString("RegionID")
intStatusID = Application("AMO_REJECT")

'---initialize variables: ---
strSaveStatus = ""

If Not IsNumeric(intAction) Then
	strSaveStatus = "Invalid call to Ajax"
Else
	intAction = CInt(intAction)
End If

if strSaveStatus = "" then
	if strModuleIDs = "" then
		strSaveStatus = "Invalid ModuleID"
    else
        strModuleIDs = Trim(strModuleIDs)
	end if
end if

if strSaveStatus = "" then
	if strRegionIDs = "" then
		strSaveStatus = "Invalid RegionID"
    else
        strRegionIDS = Trim(strRegionIDs)
	end if
end if

if strSaveStatus = "" then
	if not IsNumeric(intStatusID) then
		strSaveStatus = "Invalid StatusID"
    else
        intStatusID = CLng(intStatusID)
	end if
end if

if strSaveStatus = "" then
	
	'	On Error Resume Next
	Call SendStatusChangeEmail (True,False,intStatusID,strModuleIDs,strRegionIDs)
		
end if

if strSaveStatus <> ""  then
	response.write(strSaveStatus)
else
    response.write("success")
end if
%>
<!-- #include file="../data/closeDBConnection.asp" -->
<%
response.End
%>


