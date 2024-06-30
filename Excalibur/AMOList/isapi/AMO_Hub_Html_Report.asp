<%@ language=vbscript %>
<%Server.ScriptTimeout = 6000 %>
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
	
dim sErr, strEOLDate
dim sFilter
dim oSvr, oErr
dim oRsAMOModules
dim sBusSegIDs ,strKeyWord
	strEOLDate = Request.querystring("nHubEOLDate")
	sBusSegIDs = Request.querystring("BusSeg")

	Call SaveDBCookie( "AMO txtHubEOLDate", strEOLDate )
	Call SaveDBCookie( "AMO Hub BusSelected", sBusSegIDs )

	sFilter = "  and (R.SCMID = 1 Or R.SCMID = null) and O.SCMID = 1 and S.StatusID = 172"
	if sBusSegIDs <> "" then 
		sFilter = sFilter & " and AFD.DivisionID in (" & sBusSegIDs & ")" 
	end if
		
	if strEOLDate <> "" then
		sFilter = sFilter & " and (R.moduleID in (select distinct moduleID from AMO_Region where SCMID = 1 And RASDiscontinueDate >= '" & strEOLDate & "'))"
	end if
		
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO	
	strKeyWord=""
	set oRsAMOModules = oSvr.AMOModule_Search(Application("REPOSITORY"), sFilter, strKeyWord, sBusSegIDs, null)
	RecordReportUsage 76, Session("AMOUserID")	
	if not oRsAMOModules is nothing then
	    sErr = "Error Generating Report, AMO_Hub)Excel)Report.asp"
		Response.Write(sErr)
        Response.End()
	else
		oRsAMOModules.Sort = "Category ASC, Description ASC"
	end if
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<title><%=sH1%></title>
<script language="JavaScript" src="../library/scripts/formChek.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

//-->
</script>

</head>

<body bgcolor="#FFFFFF" LANGUAGE="javascript">
           
<form NAME="frmHW_FeaturesReport">

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1" style='color:blue'>
<%	

	if oRsAMOModules.recordcount> 0 then
		Response.Write  "<h1>Factory/Hub SKU Localization Table Report</h1><br>" 
		Response.Write  "<strong>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</strong>" 
		WriteAMO_SCM_HUB_ReportGridHtml oRsAMOModules, 1, "", sBusSegIDs, "" 
	Else
%>
	<tr>
		<td><strong>No After Market Option found that matches filter</strong></td>
	</tr>                   
	<%end if%>


	</table>
</form>
<%  
InsertConfidentialFooter 
'free the report object
oRsAMOModules.Close
set oRsAMOModules = nothing
Server.ScriptTimeout = 6000
%>
</body>
</html>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->