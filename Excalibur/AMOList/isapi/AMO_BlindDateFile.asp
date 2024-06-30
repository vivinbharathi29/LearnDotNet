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

dim sHeader, sErr, strError, sBusSegIDs, sOwnerIDs
dim nUserID
dim oSvr, oErr, oRsOptions
dim bCreateFile

sHeader = "After Market Option List - GPSy DFT Select Availability (SA) File"
sErr = ""
bCreateFile = False

nUserID = Session("AMOUserID")

'get the cookies from the Generate Export filter.
sBusSegIDs = GetDBCookie("AMO Export_chkBusSeg")
sOwnerIDs = 0

if sBusSegIDs = "" then
	strError = "No Business Segment values passed, AMO_BlindDateFile.asp"
	Response.Write(strError)
    Response.End()
end if

if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	set oRsOptions = oSvr.AMO_AnyBlindDateFileData(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs)	
end if

if sErr = "" then
    If Not oRsOptions Is Nothing then 
        do while not oRsOptions.EOF
		   if cint(oRsOptions("Base")) > 0 or cint(oRsOptions("Localized")) > 0 then
		        bCreateFile = True
                exit do
	        end if
		    oRsOptions.MoveNext
	    loop
    End If
end if
%>
<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>
<script type="text/javascript">
function BlindDate() {
	form1.action = "AMO_CreateBlindDateFile.asp";
	form1.submit();
}

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
%>
	Please select the Select Availability (SA) link to create the file.

	<p><%
	if bCreateFile then
		Response.Write "<a href=""../nj.asp"" onClick=""BlindDate(); return false;"">AMO_SelectAvailability.txt</a>"
	else
		response.write "<b>No Options are available to create an AMO_SelectAvailability.txt file.</b>"
	end if
	%>
	</p>
	
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
