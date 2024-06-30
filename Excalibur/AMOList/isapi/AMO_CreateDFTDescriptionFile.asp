<%@  Language=VBScript %>
<% OPTION EXPLICIT %>
<!------------------------------------------------------------------- 
'Description: AMO DATA
'----------------------------------------------------------------- //-->    
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataAVL.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
'printrequest
call validateSession

dim sHeader, sErr, sBusSegIDs, sOwnerIDs
dim nUserID
dim oSvr, oErr, oRsOptions

sHeader = "After Market Option List - Create File"
sErr = ""

nUserID = Session("AMOUserID")

'get the cookies from the Generate Export filter.
sBusSegIDs = GetDBCookie( "AMO Export_chkBusSeg")
sOwnerIDs = 0

if sBusSegIDs = "" or sOwnerIDs = "" then
	sErr = "No Business Segment values passed, AMO_CreateDFTDescriptionFile.asp"
	Response.Write(sErr)
    Response.End()
end if

if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	set oRsOptions = oSvr.AMO_GetDFTDescriptionData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID)
end if

if sErr = "" then
	Response.clear()
	Response.ContentType = "text/plain"
	Response.AddHeader "content-disposition", "attachment; filename=AMO_DFTDesc.txt"

	response.write oRsOptions("UserName") & vbCrLf
	response.write "MSK" & vbTab
		response.write "DESC" & vbTab & "M" & vbTab & "PROD_NBR" & vbTab 
		response.write "LANG_CD" & vbTab & "DESC_CD" & vbTab & "DESC_TXT" & vbTab
		response.write "START_EFF_DT" & vbCrLf
	
	do while not oRsOptions.EOF
		response.write "DESC" & vbTab & oRsOptions("M") & vbTab & oRsOptions("PROD_NBR") & vbTab 
		response.write "99" & vbTab & oRsOptions("DESC_CD_Common") & vbTab & stripFancyChars(oRsOptions("CD_DESC")) & vbTab
		response.write tomorrow() & vbCrLf

		oRsOptions.MoveNext
	loop

	oRsOptions.MoveFirst

	do while not oRsOptions.EOF
		response.write "DESC" & vbTab & oRsOptions("M") & vbTab & oRsOptions("PROD_NBR") & vbTab 
		response.write "99" & vbTab & oRsOptions("DESC_CD_Quote") & vbTab & stripFancyChars(oRsOptions("QU_DESC")) & vbTab
		response.write tomorrow() & vbCrLf

		oRsOptions.MoveNext
	loop

	response.write "END" & vbCrLf

	response.end
end if
%>
<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>
</HEAD>
<BODY>
<FORM name=thisform method=post>
<%

'Response.Write BuildHelp(sHeader, "")

Response.Write sErr


%>
</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->