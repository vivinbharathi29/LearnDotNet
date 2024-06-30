<%@ Language="VBScript" %>
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

	dim sModuleIdRegionIds, oSvr
	
	sModuleIdRegionIds = Request.QueryString("nModuleIdRegionIds")

%>

<!DOCTYPE html>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="AMO Bulk Date Selection" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title>AMO - Add Bulk Date</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT type="text/javascript">
<!--
function btnSave_Click() {
	var discontinueDate;
	var	mrAvailableDate; 
	var globalSeriesDate; 
    
	if(thisform.txtCPLBlindDate.value == "" && thisform.txtBOMRevADate.value == "" && thisform.txtRasDisconDate.value == "" && thisform.txtObsoleteDate.value == "" && thisform.txtGlobalseriesdate == "")
	{
		alert("Please enter at least one of the dates before proceeding.");
		return false;
	}
	
	if (document.thisform.txtGlobalseriesdate.value != "" && document.thisform.txtRasDisconDate.value == "") {
		alert("Global Series Config EOL will be removed without End of Manufacturing (EM)");
	  document.thisform.txtGlobalseriesdate.value = "";
	  return true;
	}
	  
	if (document.thisform.txtGlobalseriesdate.value != "" && document.thisform.txtBOMRevADate.value == "") {
	    alert("Global Series Config EOL will be removed without PHweb (General) Availability (GA)")
	  document.thisform.txtGlobalseriesdate.value = "";
	  return true;
	}
	
	mrAvailableDate = new Date(document.thisform.txtBOMRevADate.value)
	discontinueDate = new Date(document.thisform.txtRasDisconDate.value)
	globalSeriesDate = new Date(document.thisform.txtGlobalseriesdate.value)
		
	if(globalSeriesDate < mrAvailableDate || globalSeriesDate > discontinueDate) {
	    alert("Global Series Config EOL must fall between the PHweb (General) Availability (GA) and the End of Manufacturing (EM)");
		return false;	
	}

	if (confirm("If you\'re using Bulk Date Change to change only one of these dates, (PHweb (General) Availability (GA); End of Manufacturing (EM); Global Series Config EOL), you may break the business rule (Global Series Config EOL has to be in the range of PHweb (General) Availability (GA) and End of Manufacturing (EM)).\nAre you sure you want proceed ?")) {
		thisform.nModuleIdRegionIds.value = "<%=sModuleIdRegionIds%>";
		thisform.action = "AMO_SaveDate.asp";
		thisform.submit();
		window.opener.location.reload(true);
	}
}

function calculateCPLBlindDate(RASDate) {
	var somedate = new Date(RASDate)
	var themonth = somedate.getMonth()
	var theday = somedate.getDate()
	var theyear = somedate.getFullYear()
	//1 day prior to GA date : SUG 9763,Vinutha	
	somedate = new Date(theyear, themonth, theday-1)
	return (somedate.getMonth() + 1) + '/' + somedate.getDate() + '/' + somedate.getFullYear();
}


function calculateObsoleteDate(RASDate) {
	var somedate = new Date(RASDate);
	var themonth = somedate.getMonth();
	var theday = somedate.getDate();
	var theyear = somedate.getFullYear();
	// add 3 month to date	
	var newdate = new Date(theyear, themonth+3, theday);
	var thenewmonth = newdate.getMonth();
	var thenewyear = newdate.getFullYear();
	
	var timeA = new Date(thenewyear, thenewmonth+1, 1);
	var timeB = new Date(timeA - (60*60*24*1000)); // subtract 1 day
	var daysInMonth = timeB.getDate();
	
	somedate = new Date(thenewyear, thenewmonth, daysInMonth)
	
	return (somedate.getMonth() + 1) + '/' + somedate.getDate() + '/' + somedate.getFullYear();
}

function BOMRevADate_Update(thefield) {
    if (!checkDate(thefield, "PHweb (General) Availability (GA)", true)) {
		return false;
	}
	if (thefield.value.length > 0) {
		// calculate Select Availability (SA) only if not an empty date
		var cplObject = document.getElementById("txtCPLBlindDate")
		cplObject.value = calculateCPLBlindDate(thefield.value)
	}
}

function RASDiscontinueDate_Update(thefield) {
	if (!checkDate(thisform.txtRasDisconDate, "End of Manufacturing (EM)", true))
			return false;
		if (thefield.value.length > 0) {
			// calculate Obsolete only if not an empty date
			var obdObject = document.getElementById("txtObsoleteDate")
			obdObject.value = calculateObsoleteDate(thefield.value)
	}
}

function CPLBlindDate_Update(thefield) {
	if (!checkDate(thisform.txtCPLBlindDate, "Select Availability (SA)", true))
			return false;
}

function ObsoleteDate_Update(thefield) {
	if (!checkDate(thisform.txtObsoleteDate, "End of Sales (ES)", true))
			return false;
}

function Globalseriesdate_Update(thefield) {
		if (!checkDate (thisform.txtGlobalseriesdate, "Global Series Config EOL", true))
			return false;
}
//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="gray">
<!-- #include file="../library/includes/popup.inc" -->
	<FORM name=thisform method=post>
	<table width="100%" border="0" cellspacing="8" cellpadding="3">
		<TR>
		<TD>Select Availability (SA)</TD>
		<TD><input type="text" id=txtCPLBlindDate name=txtCPLBlindDate size=15 maxlength=10 value="" class="filter-dateselection" onBlur="CPLBlindDate_Update(this)">			
			(MM/DD/YYYY)</td></TR>
	<TR>
		<TD>PHweb (General) Availability (GA)</TD>
		<TD><input type="text" name="txtBOMRevADate" id="txtBOMRevADate" value="" size="15" maxlength="10" class="filter-dateselection" onBlur="BOMRevADate_Update(this)">
			(MM/DD/YYYY)</td></TR>

	<TR><TD>End of Manufacturing (EM)</TD>
		<TD><INPUT type="text" id=txtRasDisconDate name=txtRasDisconDate size=15 maxlength=10 value="" class="filter-dateselection" onBlur="RASDiscontinueDate_Update(this)" >
			(MM/DD/YYYY)</td></TR>		
	<TR>
		<TD>End of Sales (ES)</TD>
		<TD><input type="text" id=txtObsoleteDate name=txtObsoleteDate size=15 maxlength=10 value="" class="filter-dateselection" onBlur="ObsoleteDate_Update(this)">
			(MM/DD/YYYY)</td></TR>
			
	<TR>
		<TD>Global Series Config EOL</TD>
		<TD><input type="text" id=txtGlobalseriesdate name=txtGlobalseriesdate size=15 maxlength=10 value="" class="filter-dateselection" onBlur="Globalseriesdate_Update(this)">
			(MM/DD/YYYY)</td></TR>
		
	</table>
	
	<table>
	<tr>
		<td>
		<INPUT id=btnSave name=btnSave type=button value="Save" LANGUAGE=javascript onclick="return btnSave_Click();">
		<INPUT id=btnCancel name=btnCancel type=button value="Cancel"  LANGUAGE=javascript onClick="window.close();">
		</td>
	</tr>
	<tr>
		<td><p>&nbsp;</p><font size=1 color=red><i>Warning : Using the bulk change to update the PHweb (General) Availability (GA), End of Manufacturing (EM) or Global Series Config EOL may break a business rule. Please use the validate data option on the report tab to validate the date before publishing the SCM.</i></font></td>
	</tr>
	</table>
<INPUT type="hidden" id=nModuleIdRegionIds name=nModuleIdRegionIds value="">
<input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</FORM>
</BODY>
</HTML>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->