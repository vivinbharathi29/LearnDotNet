<%@ Language=VBScript %>
<% Option Explicit
'*************************************************************************************
'* Version		: 1.0
'* FileName		: AMO_SaveDate.asp
'* Description	: AMO List - Set Date Value for one AMO Feature
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
	Call ValidateSession

	dim oSvr, oErr, htmlErrors, sErr
	dim strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate
		
	On Error Resume Next
	htmlErrors = ""
	strModuleRegionIDs = ""
	
	strModuleRegionIDs = Request.Form("nModuleIdRegionIds")
	
	strCPLBlindDate = Request.Form("txtCPLBlindDate")
	strBOMRevADate = Request.Form("txtBOMRevADate")
	strRasDisconDate = Request.Form("txtRasDisconDate")
	strObsoleteDate = Request.Form("txtObsoleteDate")
	strGlobalseriesdate = Request.Form("txtGlobalseriesdate")
		
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	
	sErr = oSvr.UpdateBulkDate(Application("REPOSITORY"), strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate, Session("FullName"), Session("AMOUserID"))
	
	if sErr <> "True" then
	  htmlErrors = "The submitted form was missing required parameters.  Data processing was unable to complete successfuly. AMO_SaveDate.asp"
	end if
	
	set oErr = nothing
	set oSvr = nothing
	
%>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title></title>
</HEAD>
<BODY onload="window.close();" LANGUAGE=javascript>
<FORM name=thisform method=post>
<%
    if htmlErrors <> "" then 
	    Response.Write htmlErrors 
    end if
%>
</FORM>
</BODY>
</HTML>

