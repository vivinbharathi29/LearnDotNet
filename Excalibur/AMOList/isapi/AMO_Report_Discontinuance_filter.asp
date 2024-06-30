<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%Response.Expires = -1%>
<!-- #include file="../library/includes/SessionValidation.inc" -->
<% SessionValidation2 %>

<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/DLbx_Category.inc" -->
<!-- #include file="../library/includes/DualListboxRs.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../includes/AMO_Discontinuance_Report.inc" -->
<%
dim sErr, sAuthor
dim IAreaID	
dim  rsSuggestions, oErr
dim bViewRight, bEditright, bCreate, bDelete
dim bISIS, bIRSTest
dim rsAllDevelopers, sRequestor, sReleaseDate
dim sStatus, oRsAMOstatus, oRsBusSeg
on error resume next

sErr = ""

if sErr = "" then
'	set oErr = objAMO.AMO_ViewAllStatus (application("repository"), oRsAMOstatus)
    set oRsAMOstatus = GetMOLCategory(24)
    if oRsAMOstatus is nothing then  
	    sErr = "Missing required parameters.  Unable to complete your request."
	    Response.Write(sErr)
	    Response.End()
    end if
end if


if sErr = "" then
	'set oErr = GetMOLCategory(oRsBusSeg, 28)
    set oRsBusSeg = GetMOLCategory(34)	
    if oRsBusSeg is Nothing then
	    Response.Write("Recordset error: oRsBusSeg")
	    Response.End()
    end if
end if 	


%>
<html>
<!DOCTYPE html>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sHeader%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sHeader%> - AMO Properties</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script language="JavaScript" src="../library/scripts/cancel.js"></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

function btnSearch_onclick() {
	//SelectAll(thisform.lbxSelectedStatus);
	var i;
	
	if (thisform.lbxSelectedBusSeg.options.length==0)
	{
		alert(" Please select business segment" );
		
		return (false);
	}
	
	if ((thisform.txtDiscDateFrom.value=="") && (thisform.txtDiscDateTo.value==""))
	{
		alert(" Please enter Discontinue Dates." );
		thisform.txtDiscDateFrom.focus();
		return (false);
	}
	
		
	if (!checkDate (thisform.txtDiscDateFrom, "Discontinue Date (From)", true))
		return false;

	if (!checkDate (thisform.txtDiscDateTo, "Discontinue Date (To)", true))
		return false;
	if ((thisform.txtDiscDateFrom.value!="") && (thisform.txtDiscDateTo.value !=""))
	{	var startdate = newDate(thisform.txtDiscDateFrom.value)
		var enddate = newDate(thisform.txtDiscDateTo.value)
		if (( startdate.getTime() - enddate.getTime()) > 0) {
			alert("The Discontinue Date (From) is older than the Discontinue Date (To), please re-enter the dates." );
			thisform.txtDiscDateFrom.focus();
			return (false);
		}
	}	
    	for (i=0; i < thisform.lbxSelectedStatus.options.length; i++)
		{thisform.lbxSelectedStatus.options[i].selected = true;}
			for (i=0; i < thisform.lbxSelectedBusSeg.options.length; i++)
		{thisform.lbxSelectedBusSeg.options[i].selected = true;}
	SelectAll(thisform.lbxSelectedBusSeg)
	thisform.target="_blank"
	thisform.submit();
	
}


//-->
</SCRIPT>
</head>

<BODY bgcolor="#FFFFFF">

<h2>AMO Discontinuance Report-Filter</h2>

<form NAME="thisform" METHOD="post" action="AMO_Report_Discontinuance.asp">
<table BORDER="0" CELLSPACING="0" CELLPADDING="0" width = 100%>

</table>
<br>
<table BORDER="0" CELLSPACING="5" CELLPADDING="1" width = 100%>
	<colgroup width="20%"></colgroup>
	<colgroup width="40%"></colgroup>
	<colgroup width="40%"></colgroup>
	
	<tr ><td >Business Segment<font color="red">*</font></td>
		<td colspan =2 ><% call DLBoxHTML_BusSeg%></td>
	</tr>
	
	<tr ><td >Status</td>
		<td colspan =2 ><% call DLBoxHTML%></td>
	</tr>
	
			
	<TR><TD>Discontinue Date (MM/DD/YYYY)<font color="red">*</font></td>
		<td>From&nbsp;<input id="txtDiscDateFrom" name="txtDiscDateFrom" value="" class="filter-dateselection" style="HEIGHT: 22px;" size=10 maxLength="10">
            To&nbsp;<input id="txtDiscDateTo" name="txtDiscDateTo" value="" class="filter-dateselection" style="HEIGHT: 22px;" size=10 maxLength="10"></td> 
		<td><FONT size=1><EM>Enter Discontinue date range, <font color="red">*</font> required in order to limit the amount of data that will display</EM></FONT></td>	
	</tr>
	
	
	
	
</table>
<p align = left> <INPUT id=btnSearch language=javascript name=btnSearch type=button value="Continue"  onclick="return btnSearch_onclick()"> &nbsp; 

</p>
<input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</form>

<% 'InsertGlobalFooter

set oRsAMOstatus = nothing
set rsSuggestions = nothing
set rsAllDevelopers = nothing

 %>
</BODY>

</html>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
