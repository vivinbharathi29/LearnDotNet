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
dim nUserID, strError
dim oSvr, oErr, oRs, oRsOptions
dim bBase, bLocalized, bNoPhWeb

sHeader = "After Market Option List - PHweb NPI Autoload Files"
sErr = ""
bBase = False
bLocalized = False
bNoPhWeb = False

nUserID = Session("AMOUserID")

'get the cookies from the Generate Export filter.
sBusSegIDs = GetDBCookie("AMO Export_chkBusSeg")
sOwnerIDs = 0

if sBusSegIDs = "" then
	strError = "No Business Segment values passed, AMO_AutoloadFiles.asp"
	Response.Write(strError)
    Response.End()
end if


if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	set oRs = oSvr.AMO_AnyAutoloadFileData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs)
end if

if sErr = "" then
     Dim intBase, intLocalized
     If Not (oRs.EOF) Then
        do while not oRs.EOF
	        intBase = cint(oRs("Base"))
            intLocalized = cint(oRs("Localized"))
		    oRs.MoveNext
	    loop
    End If
end if

if sErr = "" then
    ' see if there are any module caetgories that do not have corresponding PHWeb categories
	if cint(intBase) = 2 or cint(intLocalized) = 2 then
		bNoPhWeb = True
		' get the list of module categories that do not have corresponding PHWeb categories
		set oRsOptions = oSvr.AMO_GetAutoloadFileData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID, 2)
	end if	   


	if (cint(intBase) = 1 or cint(intLocalized) = 1) and bNoPhWeb = False then
		bBase = True
	end if
end if
%>
<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>

<script type="text/javascript">
<% if bNoPhWeb = False then %>
function baseAutoload() {
	form1.action = "AMO_CreateAutoloadFile.asp";
	form1.submit();
}
<% end if %>

function RASReview() {
	form1.action = "AMO_ModuleList_RAS.asp";
	form1.submit();
}
</script>

</HEAD>
<BODY>
<form method="post" name="form1">
<%

if sErr = "" then
	if bNoPhWeb = True then
	%>
		<p>The following module categories do not have a corresponding AMO PHweb Category. Please assign one in "<a href="<%=Application("IRSWebServer") %>"irsplus/admin/AfterMarketOptionList/AMOPHWebCategories_Main.aspx" target="_blank">System Admin \ After Market Option List \ 
		PHweb Categories</a>" before the PHweb NPI Autoload file can be created.</p>
		<%
		do while not oRsOptions.EOF
			response.Write "<li>" & oRsOptions("ModuleCategory") & "<br>"
		
			oRsOptions.MoveNext
		loop
		oRsOptions.close
		set oRsOptions = nothing
		%>
	<%
	else
	%>
		Please select the PHweb NPI Autoload link to create the file.

		<p><%
		if bBase then
			Response.Write "<a href=""../nj.asp"" onClick=""baseAutoload(); return false;"">AMO_Autoload.txt</a>"
		else
			response.write "<b>No Options are available to create an AMO_Autoload.txt file.</b>"
		end if
		%>
		</p>
	<% end if %>
	
	<p>&nbsp;</p>
	<input type="button" style="width:150px;" name="rasreview" id="rasreview" value="Return to RAS Review" onClick="RASReview();">
<%
else
	Response.Write sErr
end if


%>
</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->