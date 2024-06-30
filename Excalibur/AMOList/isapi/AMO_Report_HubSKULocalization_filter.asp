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
<% SessionValidation2 %>
<%
	dim sH1, err_sHTML, nDivisionID
	dim objErr, rsDivision, rsSelOwner, nMode, i, strHubEOLDate
	dim oServer, oErr, sHTMLBussiness, rsDivisionSelected


	set objErr = nothing
	set rsDivision = nothing
	set rsDivisionSelected = Nothing
	
	' Define headers
	if (objErr is nothing) then
		sH1 = "Factory/Hub SKU Localization Report"
	end if	
	
	
	if Request.Form("txtHubEOLDate") = "" then
	'get the cookie. If we didn't get it default it
		strHubEOLDate = GetDBCookie( "AMO txtHubEOLDate")
		if trim(strHubEOLDate) = "" then
			'set default day to 1 month prior
			strHubEOLDate = dateAdd("m", -1, date)
		end if
	else
		strHubEOLDate = Request.Form("txtHubEOLDate")
	end if
	'store the cookie
	Call SaveDBCookie( "AMO txtHubEOLDate", strHubEOLDate )
	

	set objErr = nothing
	' Get business segments the login user is in
	'set objErr = GetCategory(rsDivision, 44, session("USERID"))
	set rsDivision = GetCategory(44, Session("AMOUserID"))	
	if rsDivision is Nothing then
		Response.Write("Recordset error: rsDivision")
		Response.End()
    end if

	if not rsDivision is nothing then
		nDivisionID = ""
		if rsDivision.RecordCount > 0 then
			rsDivision.MoveFirst
			for i = 0 To rsDivision.recordCount-1
				if nDivisionID <> "" then
					nDivisionID = nDivisionID & "," & rsDivision.Fields("DivisionID").Value
				else
					nDivisionID = rsDivision.Fields("DivisionID").Value
				end if
				rsDivision.MoveNext
			Next
		end if
	end if
	
	' Get all business segments
	'set objErr = GetCategory(rsDivision, 44, 0)
	set rsDivision = GetCategory(44, 0)	
	if rsDivision is Nothing then
		Response.Write("Recordset error: rsDivision")
		Response.End()
    end if
	
	if Request.QueryString("nDivisionID") <> "" then
		nDivisionID = Request.QueryString("nDivisionID")
	end if
	
	
	sHTMLBussiness = DualListboxRs_GetHTML7(rsDivision, "Division", "DivisionID", rsDivisionSelected, "Division", "DivisionID", True, False, nDivisionID, _
							"Available Business Segment", "Selected Business Segment", "Division", True, 7, 300, False, False, False, 1, 1, "", "" )	
%>
<!DOCTYPE html>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sH1%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sH1%> - AMO Properties</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script language="JavaScript" src="../library/scripts/DateValidation.js"></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
function btnSearch_onclick() {
	var sText, sDivisionIds, sEOLDate;
	
	
	sEOLDate = frmHW_ModuleFilters.txtHubEOLDate.value;
	if (!checkDate(frmHW_ModuleFilters.txtHubEOLDate, "Show options with End of Manufacturing (EM)", true))
		return false;
		
	
	SelectAll(frmHW_ModuleFilters.lbxSelectedDivision);
	if (frmHW_ModuleFilters.lbxSelectedDivision.value == "") {
		alert("Please select at least one business segment before proceeding");
		return false;
	}
	

	sDivisionIds = "";
	for (i=0; i < frmHW_ModuleFilters.lbxSelectedDivision.options.length; i++)
		sDivisionIds = sDivisionIds + "," + frmHW_ModuleFilters.lbxSelectedDivision.options[i].value;	
	if (sDivisionIds != "")
		sDivisionIds=sDivisionIds.slice(1);
				
	
	if (frmHW_ModuleFilters.chkExcel.checked) 
		  sText = "AMO_Hub_Excel_Report.asp?BusSeg=" + sDivisionIds + "&nHubEOLDate=" + sEOLDate;
	else 
		  sText = "AMO_Hub_Html_Report.asp?BusSeg=" + sDivisionIds + "&nHubEOLDate=" + sEOLDate;
		  
	if (frmHW_ModuleFilters.chkNewWindow.checked) {
		frmHW_ModuleFilters.target = "_blank";		
		frmHW_ModuleFilters.action = sText;
		return true;
	}
	else {
		frmHW_ModuleFilters.target = "_self";
		frmHW_ModuleFilters.action = sText;
		return true;
	}

}


//-->
</script>

<script language="javascript" src="../library/scripts/cancel.js"></script>
<script language="JavaScript" src="../library/scripts/calendar.js"></script>

</head>

<body bgcolor="#FFFFFF">       
<%='BuildHelpCentered("Factory/Hub SKU Localization Report", "../../Help/ComponentMgmt/Report/HELP_Component_Management_Reports.asp#hardware_features","640") %>
<form NAME="frmHW_ModuleFilters" METHOD="post" action="">
<table border="0" cellPadding="0" cellSpacing="0" WIDTH="100%" align="Center">
    <%	if (objErr is nothing) then %>	
    <tr>
	<td>
	<table WIDTH="65%" BORDER="0" CELLSPACING="2" CELLPADDING="2" align="Center">	
			<TR bgcolor="lightsteelblue" STYLE="FONT-WEIGHT: bold" height="21"><TD>Business Segment<font color="red" face="">*</font>:</TD></tr>
			<tr>
					<td><%=sHTMLBussiness%></td>
			</tr> 
			<TR bgcolor="lightsteelblue" STYLE="FONT-WEIGHT: bold" height="21"><TD>Publish options with End of Manufacturing (EM)(<i>base unit</i>) on or after<font color="red" face="">*</font>:</TD></tr>
			<tr>
				<td><input type=text maxlength=10 size=10 id=txtHubEOLDate NAME=txtHubEOLDate value=<%= strHubEOLDate %> class="filter-dateselection">(MM/DD/YYYY)</td>
			</tr>
			 <TR bgcolor="lightsteelblue" STYLE="FONT-WEIGHT: bold" height="21"><TD>Report Options:<font color=maroon size=1>&nbsp;&nbsp;<em>Report is default to html format</em></font></TD></tr>
				<tr>
					<td><INPUT id=chkNewWindow name=chkNewWindow type=checkbox checked>New window
						<INPUT id=chkExcel name=chkExcel type=checkbox >Excel format
					</td>
				</tr>
				<tr>
				<td><i><font color="red">If Reports are created in Excel format, Microsoft Office XP is needed to run the reports.</font></i></td></tr>
			
				<tr>
					<td><hr></td>
				</tr>
				<tr>
					<td align="left"><font color="red" face="">*</font>&nbsp;&nbsp;Required field.</td>
				</tr>
				<tr>
					<td align="center"><br>
						<input id="btnSearch" name="btnSearch" type="submit" value="Display Report" LANGUAGE="javascript" onclick="return btnSearch_onclick()">
					</td>
				</tr>
			</table>
		</td>
	</tr>
    <%	else	%>			
	<tr>
		<td><%=err_sHTML%> 				
		</td>
	</tr>
    <%	end if	%>
</TABLE>
	<input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />	
</form>
<% 'InsertGlobalFooter %>
</body>
</html>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->