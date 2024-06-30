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
call ValidateSession

dim sHeader, sErr, sBusSegIDs, sOwnerIDs
dim intDummyCount
dim nUserID, nAutoloadfile
dim bHaveLocalized
dim oSvr, oErr, oRsOptions, oRsLocalOptions
dim bLocalizedrow

sHeader = "After Market Option List - Create PHweb NPI Autoload File"
sErr = ""

nUserID = Session("AMOUserID")



'get the cookies from the Generate Export filter.
sBusSegIDs = GetDBCookie( "AMO Export_chkBusSeg")
sOwnerIDs = GetDBCookie( "AMO Export_chkGroupOwner")

if sBusSegIDs = "" then
	sErr = "No Business Segment or Owned By values passed, AMO_CreateAutoloadFile.asp"
	Response.Write(sErr)
    Response.End()
end if

if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	'Get Base Part Number file data
	set oRsOptions = oSvr.AMO_GetAutoloadFileData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID, 0)
	if (oRsOptions is nothing) then
		sErr = "No Business Segment or Owned By values passed, AMO_CreateAutoloadFile.asp"
		Response.Write(sErr)
        Response.End()
	end if
end if

'if sErr = "" then
	'Get Localized file data
'	set oRsLocalOptions = oSvr.AMO_GetAutoloadFileData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID, 1)
'	if (oRsLocalOptions is nothing) then
'		sErr = "No Business Segment or Owned By values passed, AMO_CreateAutoloadFile.asp"
'		Response.Write(sErr)
'       Response.End()
'	end if
'end if

if sErr = "" then
	Response.clear()
	Response.ContentType = "text/plain"
	Response.AddHeader "content-disposition", "attachment; filename=AMO_Autoload.txt"
	CommonHeader
	response.write vbTab & vbTab
	response.write "Initial MLORW Target Cost" & vbCrLf

	'oRsOptions.MoveFirst
	intDummyCount = 0
  
	do while not oRsOptions.EOF
		'see if there is any localized data for the base part number
		'oRsLocalOptions.Filter = "ModuleID=" & cstr(oRsOptions("ModuleID"))
		if oRsOptions("HasLocalized") = "Yes" then
			bHaveLocalized = true           
		else
			bHaveLocalized = false
		end if
        if cint(oRsOptions("Localized")) = 1 then
			bLocalizedrow = true           
		else
			bLocalizedrow = false
		end if
    
		'row 1
        if not bLocalizedrow then
		    response.write "ADD" & vbTab & oRsOptions("ParentPin") & vbTab & "dummy" & cstr(intDummyCount) & vbTab & oRsOptions("RowLevel1") & vbTab
    
		    BaseCommonFields(False)
		    response.write oRsOptions("MLORW") & vbCrLf
		    'row 2
		    response.write "ADD" & vbTab & "dummy" & cstr(intDummyCount) & vbTab  & oRsOptions("Pin") & vbTab & oRsOptions("RowLevel2") & vbTab
		    BaseCommonFields(bHaveLocalized)
		    response.write vbCrLf
	    
		    intDummyCount = intDummyCount + 1
        else
        
		   ' if bHaveLocalized then
			    'there is localized data for that module
			'    do while not oRsLocalOptions.EOF
				response.write "ADD" & vbTab & oRsOptions("ParentPin") & vbTab & oRsOptions("Pin") & vbTab & oRsOptions("Level") & vbTab
				response.write oRsOptions("ODMCode") & vbTab & oRsOptions("FormatCode") & vbTab & oRsOptions("CategoryGroup") & vbTab 
				response.write stripFancyChars(oRsOptions("ShortDescription")) & " " & oRsOptions("LocalCountrification") & vbTab & oRsOptions("BrandName") & vbTab & stripFancyChars(oRsOptions("ShortDescription")) & vbTab 
				response.write oRsOptions("LocalCountrification") & vbTab & oRsOptions("LCStatusLocal") & vbTab & makeDateFormat(oRsOptions("AvailDate")) & vbTab 
				response.write makeDateFormat(oRsOptions("EOLDate")) & vbTab & makeDateFormat(oRsOptions("DiscBuild")) & vbTab & makeDateFormat(oRsOptions("OBSDate")) & vbTab 
				response.write oRsOptions("SDFFlagLocal")
				response.write vbCrLf
				
			'	oRsLocalOptions.MoveNext
			'loop
		end if
		'oRsLocalOptions.Filter = ""
		
		oRsOptions.MoveNext
	loop
	response.End
   
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

<%
function CommonHeader()
	response.write "Action" & vbTab & "Parent Pin" & vbTab & "Pin" & vbTab & "Level" & vbTab
	response.write "ODM Code" & vbTab & "Format Code" & vbTab & "Category Group" & vbTab
	response.write "Short Description" & vbTab & "Brand Name" & vbTab & "Long Description" & vbTab
	response.write "Countrification" & vbTab & "LC Status" & vbTab & "Avail. Date" & vbTab
	response.write "EOL Date" & vbTab & "Disc Build" & vbTab & "OBS Date" & vbTab
	response.write "SDF Flag"' & vbTab & vbTab
end function

function BaseCommonFields(byVal bHaveLocalized)
	response.write oRsOptions("ODMCode") & vbTab & oRsOptions("FormatCode") & vbTab & oRsOptions("CategoryGroup") & vbTab 
	response.write stripFancyChars(oRsOptions("ShortDescription")) & vbTab & oRsOptions("BrandName") & vbTab & stripFancyChars(oRsOptions("ShortDescription")) & vbTab 
	response.write oRsOptions("Countrification") & vbTab
	if bHaveLocalized then
		response.write oRsOptions("LCStatusBaseWithLocal")
	else
		response.write oRsOptions("LCStatus")
	end if
	response.write vbTab & makeDateFormat(oRsOptions("AvailDate")) & vbTab 
	response.write makeDateFormat(oRsOptions("EOLDate")) & vbTab & makeDateFormat(oRsOptions("DiscBuild")) & vbTab & makeDateFormat(oRsOptions("OBSDate")) & vbTab 
	if bHaveLocalized then
		response.write oRsOptions("SDFFlagBaseWithLocal") & vbTab & vbTab
	else
		response.write oRsOptions("SDFFlag") & vbTab & vbTab
	end if
end function
%>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->