<%@ language=vbscript %>
<% 
    Server.ScriptTimeout = 6000
    Response.Expires = 0
%> 
<!-- #include file="../library/includes/ErrHandler.inc" -->
<%
     'Pass Values to IPulsar AMO SCM Comparison Report at page load: -----
	dim sSCMIDS, sFilter, sBusSegIDs , strEOLDate, sErr
	dim iCompareCurrent
	
    sSCMIDS = Request.Form("chkCompareSCM")
	
	sBusSegIDs = Request.Form("chkBusSeg1")
	
	iCompareCurrent = Request.QueryString("CompareCurrent")
	
	strEOLDate = Request.Form("txtEOLDate")
	

    'if sBusSegIDs <> "" then 
		'sFilter = sFilter & " and AFD.DivisionID in (" & sBusSegIDs & ")" 
	'end if
    'if strEOLDate <> "" then
		'sFilter = sFilter & " and (O.RASDiscontinueDate >= '" & strEOLDate & "' or O.RASDiscontinueDate is null) "
	'end if

    'Response.Write("?SCMIDS=" & sSCMIDS & "&chkBusSeg1=" & sBusSegIDs & "&CompareCurrent=" & CompareCurrent & "&txtEOLDate=" & strEOLDate)
    'Response.End()

	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    'set oSvr = New ISAMO				
	'if len(CompareCurrent) > 0 then
		'set rsSCMCompare = oSvr.AMO_SCM_CompareCurrent(Application("REPOSITORY"),sSCMIDS, sFilter)
		'if rsSCMCompare is nothing then
			'sErr = "Missing required parameters.  Unable to complete your request."
		    'Response.Write(sErr)
		    'Response.End()
		'end if
		
		'set RsTwoAMOPublishs  = oSvr.ViewTwo_AMO_SCM_Publish(Application("REPOSITORY"), sSCMIDS)
		'if RsTwoAMOPublishs  is nothing then
			'sErr = "Missing required parameters.  Unable to complete your request."
		    'Response.Write(sErr)
		    'Response.End()
		'end if
	'else
	
		'set rsSCMCompare = oSvr.AMO_SCM_Comparison(Application("REPOSITORY"),sSCMIDS)
		'if rsSCMCompare is nothing then
			'sErr = "Missing required parameters.  Unable to complete your request."
		    'Response.Write(sErr)
		    'Response.End()
		'end if	

		'set RsTwoAMOPublishs = oSvr.ViewTwo_AMO_SCM_Publish(Application("REPOSITORY"), sSCMIDS)
		'if RsTwoAMOPublishs is nothing then
			'sErr = "Missing required parameters.  Unable to complete your request."
		    'Response.Write(sErr)
		    'Response.End()
		'end if	
	'end if
	
%>
<!DOCTYPE html>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title>AMO SCM Comparison - Report</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/wizard%20style.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--
    function window_onload() {
        window.location.href ='/IPulsar/Reports/AMO/AMO_SCM_Comparison.aspx?SCMID=<%=sSCMIDS%>&BusSeg=<%=sBusSegIDs%>&CompareCurrent=<%=iCompareCurrent%>&EOLDate=<%=strEOLDate%>';
    }
//-->
</SCRIPT>
</head>
<BODY LANGUAGE=javascript  onload = "return window_onload();">

</BODY>
</html>

