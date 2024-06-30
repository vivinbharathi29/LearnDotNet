<%@ Language=VBScript %>
<% OPTION EXPLICIT 
Server.ScriptTimeout = 6000 %>
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
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<%
Call ValidateSession
'printrequest

dim sHeader, sHelpFile, sErr
dim sModuleCategoryHTML, sModuleCategoryHTML2
dim nCategoryID, nMode, nEdit
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, bRAS
dim bRASCreate, bRASView, bRASUpdate, bRASDelete
dim oSvr, oErr, sCtrlStyle, nNumTotalRules, sRuleId
dim oRsAllCategory, oRsOptionRules, sCategoryID
dim intCount, oRsCreateGroups, oRsBusSeg, oRsBusSegSelected
dim sRuleDescription, sMin, sMax, sDivisionIds, sCategory
const HIGHLIGHT = "#FFFF99"

dim nMsgMode
nMsgMode=0

if len(Request.QueryString("MsgMode")) > 0 then
	nMsgMode = clng(Request.QueryString("MsgMode"))

end if 
sHelpFile = "../help/HELP_AMO_Platforms.asp"
sHeader = "After Market Option List"

'set rsRoles and IRSUserID Session: ----
Call SetPermission()

'get permissions
GetRights2 Application("AMOList"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMORAS"), bRASCreate, bRASView, bRASUpdate, bRASDelete


if bAMOUpdate or bAMOCreate then
	sCtrlStyle = ""
else
	sCtrlStyle = " disabled "				
end if


set oRsOptionRules = nothing
set oRsBusSeg = Nothing
set oRsBusSegSelected = Nothing

sRuleDescription = ""
sMin = ""
sMax = ""
sDivisionIds = ""

nMode = clng(Request.QueryString("Mode"))

'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO

if nMode = 1 then
	sCategoryID = Request.QueryString("CategoryID")
	sRuleId = Request.QueryString("RuleID")
else
	sCategoryID = Request.Form("lbxCategory2")
end if

if sCategoryID = "" then
	sCategoryID = 0
end if 

if sRuleID = "" then
	sRuleID = 0
end if 


set oRsOptionRules = oSvr.AMO_RulesSearch(Application("REPOSITORY"), clng(sCategoryId), clng(sRuleId))

if (oRsOptionRules is nothing) then
    sErr = "Missing required parameters.  Unable to complete your request., AMO_AutoloadFiles.asp"
	Response.Write(sError)
    Response.End()
else
    nNumTotalRules = oRsOptionRules.RecordCount
    if oRsOptionRules.RecordCount > 0 and sRuleID <> 0 then
	    sRuleDescription = oRsOptionRules("Rule_Desc").Value
	    sMin = oRsOptionRules("Rule_Min").Value
	    sMax = oRsOptionRules("Rule_Max").Value
	    sDivisionIds = oRsOptionRules("DivisionIds").Value
	    sCategory = oRsOptionRules("Category_Desc").Value
    end if
end if

'set oErr = GetMOLCategory(oRsBusSeg, 28)
set oRsBusSeg = GetMOLCategory(34)	
if (oRsBusSeg is Nothing) then
	Response.Write("Recordset error: oRsBusSeg")
	Response.End()
end if

'get the list of module/option category groups
sModuleCategoryHTML = ""
set oRsAllCategory = nothing

'set oErr = GetMOLCategory(oRsAllCategory, 13)
set oRsAllCategory = GetMOLCategory(33)	
if oRsAllCategory is Nothing then
	Response.Write("Recordset error: oRsAllCategory")
	Response.End()
else
	if not oRsAllCategory is nothing then
		if not oRsAllCategory.EOF and not oRsAllCategory.BOF then
			oRsAllCategory.Filter = "State = 'Active' "
			oRsAllCategory.Sort = "Name ASC"
			if nCategoryID = -1 and oRsAllCategory.RecordCount > 1 then
				'move to the second item in the list past All to get the ID
				'so the default is the second item instead of All
				oRsAllCategory.MoveFirst
				oRsAllCategory.MoveNext
				nCategoryID = clng(oRsAllCategory("FeatureCategoryID"))
				oRsAllCategory.MoveFirst
			end if
		end if
		sModuleCategoryHTML = Lbx_GetHTML5("lbxCategory", false, 1, 0, _
					oRsAllCategory, "Name", "FeatureCategoryID",  clng(sCategoryID), false, "", false)	
					
		oRsAllCategory.Filter = ""
		oRsAllCategory.Filter = "State = 'Active'"			
		sModuleCategoryHTML2 = Lbx_GetHTML5("lbxCategory2", false, 1, 0, _
					oRsAllCategory, "Name", "FeatureCategoryID", clng(sCategoryID), false, "", false)	
	end if 
end if


%>
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
<title><%=sHeader%> - Reports</title>
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
<SCRIPT ID=clientEventHandlersJS type="text/javascript" LANGUAGE=javascript>
<!--

function btnApply_onclick() {
	//steven: only prompt when new Rule
	if ( ValidateInput() ){
	<%if nMsgMode<>1 then %>
			if (!confirm("Are you sure you want to create a new rule for the selected Feature Categories?"))
				return false;
	<% end if %>	 	
		SelectAll(document.getElementById("lbxSelectedDivision"));
		thisform.action = "AMO_CategoryRulesSave.asp?Mode=1&RuleID=<%=sRuleID%>&nCategory=<%=sCategory%>"
		return thisform.submit();
	}
}

function btnRemove_onclick(sCategoryID, sRuleId) {
	if (!confirm("Are you sure you want to remove the category rule ?"))
		return;
			
	thisform.action = "AMO_CategoryRulesSave.asp?Mode=2&CategoryID=" +  "<%=sCategoryID%>" + "&RuleID=" + sRuleId
	return thisform.submit();
}

function btnFilter_onclick() {
	thisform.action = "AMO_OptionCategoryRule.asp?Mode=3&CategoryID=" +  document.getElementById("lbxCategory2")
	return thisform.submit();
}

function btnClear_onclick() {
	document.getElementById("txtRulesDescription").value = "";
	document.getElementById("txtRulesMin").value = "";
	document.getElementById("txtRulesMax").value = "";
}

function ValidateInput() {
	if (isWhitespace(thisform.txtRulesDescription.value) ) {
		warnInvalid(thisform.txtRulesDescription, "Please enter Rules Description before proceed.");
		return false;
	}
	
	if (isWhitespace(thisform.txtRulesMin.value) ) {
		warnInvalid(thisform.txtRulesMin, "Please enter Rules Min before proceed.");
		return false;
	}
	
	if (isWhitespace(thisform.txtRulesMax.value) ) {
		warnInvalid(thisform.txtRulesMax, "Please enter Rules Max before proceed.");
		return false;
	}
	

	SelectAll(thisform.lbxSelectedDivision);
	if (thisform.lbxSelectedDivision.value == "") {
		alert("Please select at least one business segment before proceed....");
		return false;
	} 	
		
	return true;
}

function checkNumeric(e) {
	// Get ASCII value of key that user pressed
	if (!e) e = window.event;
	var key = e.keyCode ? e.keyCode : e.which;

	// Was key that was pressed a numeric character (0-9) or backspace?
	if (( key > 47 && key < 58 ) || key == 8 ) //|| key == 46 decimal point
		return; // if so, do nothing
	else // otherwise, discard character 
		if (window.event)
			e.returnValue = null; // IE
		else
			e.preventDefault(); // Firefox
}

function checkNumeric01(e) {
	// Get ASCII value of key that user pressed
	if (!e) e = window.event;
	var key = e.keyCode ? e.keyCode : e.which;

	// Was key that was pressed a numeric character (0-1) or backspace?
	if (key == 48 || key == 49 || key == 8 )
		return; // if so, do nothing
	else // otherwise, discard character 
		if (window.event)
			e.returnValue = null; // IE
		else
			e.preventDefault(); // Firefox
}
//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="#FFFFFF">
<!-- #include file="../library/includes/popup.inc" -->
<h1 class="page-title"><%=sHeader%></h1>
<FORM NAME=thisform method=post>
<%

Response.Write ""
WriteTabs "Rules"

if sErr <> "" then
	Response.Write sErr
else
	%>
	<div ID=erroutputArea></div>
	<TABLE border=0 cellPadding=2 cellSpacing=2 width=100%>
		<tr>
			<td colspan=3>
			<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
				<tr>
					<td width=90%><strong>Update Feature Category Rules</strong></td>
					<td><INPUT id=btnClear name=btnClear type=button value="Clear Top" LANGUAGE=javascript onclick="return btnClear_onclick()"></td>
				</tr>
			</TABLE>
			</td>
		</tr>
		
		<tr><td colspan=3><hr></td></tr>
		<tr>
			<td width=20%>Feature Categories</td>
			<td>&nbsp;<%=sModuleCategoryHTML%></td>
		</tr>
		<br>
		<TR>
			<TD width=20%>Rules Description</TD>
			<TD colspan=2 valign="middle">
				<table border=0>
					<tr>
						<td><textarea cols="50" rows="3" id="txtRulesDescription" name="txtRulesDescription" maxlength=1024><%=sRuleDescription%></textarea></td>
						<td>
							<table border=0>
								<tr>
									<td>Min</td>
									<td><input onKeyPress="return checkNumeric01(event)" type=text id="txtRulesMin" name="txtRulesMin" size=1 maxlength=1 value='<%=sMin%>'></td>
									<td><font size=1><i>0 or 1</i></font></td>
								</tr>
								<tr>
								<td>Max</td>
								<td><input onKeyPress="return checkNumeric(event)" id="txtRulesMax" name="txtRulesMax" size=3 maxlength=3 value='<%=sMax%>'></td>
								<td><i><font size=1>0 - 999</i></font></td>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
			</TD>
		</TR>  
		
		<tr>
				<td width=20%>Applies to SCMs Published By</td>
				<td width=80%>
				<%	
					DualListboxRs_GetHTML6_Write oRsBusSeg, "SegmentName", "BusinessSegmentID", oRsBusSegSelected, _
					"Segmentname", "BusinessSegmentID", true, true, sDivisionIds, "Available", "Selected", _
					"Division", true, 130, 250, false, true, false, 350, 13		
				%>
				</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td colspan=3 align="center"><INPUT id=btnApply name=btnApply type=button value="Apply" LANGUAGE=javascript onclick="return btnApply_onclick()" <%=sCtrlStyle%>></td>
		</tr>
		<tr><td colspan=3><hr></td></tr>
	</table>
	
	<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
		<tr>
			<td>
				<TABLE border=0 cellPadding=0 cellSpacing=0 width=70%>
					<tr>
						<td width=20%>Feature Categories</td>
						<td><%=sModuleCategoryHTML2%></td>		
						<td><INPUT id=btnFilter name=btnFilter type=button value="Filter" LANGUAGE=javascript onclick="return btnFilter_onclick()"></td>	
					</tr>
				</TABLE>
			</td>
		</tr>
	</TABLE>
	
	<TABLE border=0 cellPadding=1 cellSpacing=1 width="100%">
		<colgroup></colgroup>
		<%
		if nNumTotalRules <= 0 then
			Response.Write "<tr><td>"
			Response.Write "<br>No Feature Categories Rules have been found for the above filter.<br><br><br>" 
			Response.Write "</td></tr>" & vbCrLf
		else
			%>
			<tr><td align=left>	
			<%			
				WriteOptionCategoryRulesGridHTML sCtrlStyle, oRsOptionRules
				
			%>
			</td></tr>
			<%
		end if 
		%>
	</TABLE>
    <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</FORM>
<%
end if

%>
</BODY>
</HTML>
<%
    '---Close DB Connection: ---
    Set oSvr = New DBConnection 
    oSvr.CloseDBConnection(True)
%>