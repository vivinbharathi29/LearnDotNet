<%@ Language=VBScript %>
<% OPTION EXPLICIT 
	
    Response.Buffer = False
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"

    Server.ScriptTimeout = 6000 
%>
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
Call ValidateSession
'printrequest

dim sHeader, sHelpFile, sErr, sStatusIDs, strEOLDate
dim sModuleCategoryHTML, sStatusHTML, sBusSegIDs, sBusSegHTML
dim nCategoryID, nNumTotalModules, sGpsyCom, sRasCom, nMode, sHideColumns
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, bRAS, bGpsyCom, bRasCom
dim bRASCreate, bRASView, bRASUpdate, bRASDelete
dim oSvr, oErr
dim oRsAMOModules
dim sDivisionIDs, intCount, oRsCreateGroups, sOwnerIDs, sOwnersHTML
dim bDisplayHideromMOL, bDisplayHideromSCM, sDisplayHideromMOL, sDisplayHideromSCM, sDisplayHideromSCL, bDisplayHideromSCL, bDisplayIncludeBlankGSD, sDisplayIncludeBlankGSD
dim strMRFromDate, strMRToDate, strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, strRasObsoleteToDate, strRasObsoleteFromDate
dim strchkMRBlank, strchkCPLBlank, strchkDisBlank, sGroupIDs, nCategoriesSelected, oRsCategories, strKeyWord

const HIGHLIGHT = "#FFFF99"

sHelpFile = "../help/HELP_AMO_Platforms.asp"
sHeader = "After Market Option List"

 'set rsRoles and IRSUserID Session: ----
Call SetPermission()

'get permissions
GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMORAS_Permission"), bRASCreate, bRASView, bRASUpdate, bRASDelete

'get user permission to Business Segement and Group: ---
'Validate User Role has Permission to either CREATE AMO List, AMO RAS, or AMO Cost; if so, creates a string of Business Segment and Group IDs the User has additional access to: ----
set oRsCreateGroups = GetGroupsForRole2(cstr(Session("AMOUserRoleNames")), cstr(Application("AMO_Permission")), true, false, false, false, true, false)
if (oRsCreateGroups is nothing) then
	Response.Write("Empty Recordset: oRsCreateGroups")
	Response.End()
else
	sDivisionIDs = ""
	sGroupIDs = ""
	For intCount = 0 To oRsCreateGroups.RecordCount-1
		sDivisionIDs = "," & sDivisionIDs & replace(oRsCreateGroups("DivisionIDs"), "|", ",")
		sGroupIDs = "," & sGroupIDs & oRsCreateGroups("GroupID") & ","
		oRsCreateGroups.MoveNext		
	Next
end if	

bRAS = False
if bRASUpdate then
	bRAS = True
end if


sHideColumns = ""

if Request.Form("lbxCategory") = "" then
	'get the cookie. If we didn't get it default it
	nCategoryID = GetDBCookie( "AMO lbxCategory")
	if trim(nCategoryID) = "" then
		nCategoryID = -1
	else
		nCategoryID = clng(nCategoryID)
	end if
else
	nCategoryID = clng(Request.Form("lbxCategory"))
end if
'store the cookie
Call SaveDBCookie( "AMO lbxCategory", cstr(nCategoryID) )

if Request.Form("chkStatus") = "" then
	'get the cookie. If we didn't get it default it
	sStatusIDs = GetDBCookie( "AMO chkStatus")
	if trim(sStatusIDs) = "" then
		sStatusIDs = cstr(Application("AMO_NEW")) & "," & cstr(Application("AMO_COMPLETE")) & "," & cstr(Application("AMO_RASREVIEW")) & "," & cstr(Application("AMO_RASUPDATE"))
	end if
else
	sStatusIDs = Request.Form("chkStatus")
end if
'store the cookie
Call SaveDBCookie( "AMO chkStatus", sStatusIDs )

if Request.Form("chkBusSeg") = "" then
	'get the cookie. If we didn't get it default it
	sBusSegIDs = GetDBCookie( "AMO chkFeatureBusSeg")
else
	sBusSegIDs = Request.Form("chkBusSeg")
end if
'store the cookie
Call SaveDBCookie( "AMO chkFeatureBusSeg", sBusSegIDs)

if Request.Form("chkGroupOwner") = "" then
	'get the cookie. If we didn't get it default it
	sOwnerIDs = GetDBCookie( "AMO chkGroupOwner")
else
	sOwnerIDs = Request.Form("chkGroupOwner")
end if
'store the cookie
Call SaveDBCookie( "AMO chkGroupOwner", sOwnerIDs )

nMode = Request.QueryString("nMode")
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strMRFromDate = GetDBCookie( "AMO txtMRFromDate")
else
	strMRFromDate = Request.Form("txtMRFromDate")
	Call SaveDBCookie( "AMO txtMRFromDate", strMRFromDate)
end if
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strMRToDate = GetDBCookie( "AMO txtMRToDate")
else  
	strMRToDate = Request.Form("txtMRToDate")
	'store the cookie
Call SaveDBCookie( "AMO txtMRToDate",strMRToDate)
end if
	
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strCPLFromDate = GetDBCookie( "AMO txtCPLFromDate")
else
	strCPLFromDate = Request.Form("txtCPLFromDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtCPLFromDate",strCPLFromDate)
end if
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strCPLToDate = GetDBCookie( "AMO txtCPLToDate")
else
	strCPLToDate = Request.Form("txtCPLToDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtCPLToDate",strCPLToDate)
end if
		
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strDisFromDate = GetDBCookie( "AMO txtDisFromDate")
else
	strDisFromDate = Request.Form("txtDisFromDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtDisFromDate",strDisFromDate)
end if
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strDisToDate = GetDBCookie( "AMO txtDisToDate")
else
	strDisToDate = Request.Form("txtDisToDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtDisToDate",strDisToDate)
end if


if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strRasObsoleteToDate = GetDBCookie( "AMO txtRasObsoleteToDate")
else
	strRasObsoleteToDate = Request.Form("txtRasObsoleteToDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtRasObsoleteToDate",strRasObsoleteToDate)
end if
	
	
if nMode <> 1 Then 
	'get the cookie. If we didn't get it default it
	strRasObsoleteFromDate = GetDBCookie( "AMO txtRasObsoleteFromDate")
else
	strRasObsoleteFromDate = Request.Form("txtRasObsoleteFromDate")
	'store the cookie
	Call SaveDBCookie( "AMO txtRasObsoleteFromDate",strRasObsoleteFromDate)
end if
	
	
if nMode <> 1 Then
	'get the cookie. If we didn't get it default it
	strchkCPLBlank = GetDBCookie( "AMO chkCPLBlankDate")
else
	strchkCPLBlank = Request.Form("chkCPLBlankDate")
	Call SaveDBCookie( "AMO chkCPLBlankDate", Request.Form("chkCPLBlankDate"))	
end if


if strchkCPLBlank <> "on" then
	strchkCPLBlank = ""
else
	strchkCPLBlank = "checked"
end if
	
	
if nMode <> 1 Then
	'get the cookie. If we didn't get it default it
	strchkMRBlank = GetDBCookie( "AMO chkMRBlankDate")
else
	strchkMRBlank = Request.Form("chkMRBlankDate")
	Call SaveDBCookie( "AMO chkMRBlankDate", Request.Form("chkMRBlankDate"))	
end if


if strchkMRBlank <> "on" then
	strchkMRBlank = ""
else
	strchkMRBlank = "checked"
end if
	
	
if nMode <> 1 Then
	'get the cookie. If we didn't get it default it
	strchkDisBlank = GetDBCookie( "AMO chkDisBlankDate")
else
	strchkDisBlank = Request.Form("chkDisBlankDate")
	Call SaveDBCookie( "AMO chkDisBlankDate", Request.Form("chkDisBlankDate"))	
end if

if strchkDisBlank <> "on" then
	strchkDisBlank = ""
else
	strchkDisBlank = "checked"
end if



'save cookies for Ras and GPSy incompleted filter
			
if nMode <> 1 Then
	bGpsyCom = GetDBCookie( "AMO chkGpsyCom")
else
	bGpsyCom = Request.Form("chkGpsyCom")
	Call SaveDBCookie( "AMO chkGpsyCom", Request.Form("chkGpsyCom"))
end if

			
		
if nMode <> 1 Then
	bRasCom = GetDBCookie( "AMO chkRasCom")
else
	bRasCom = Request.Form("chkRasCom")
	Call SaveDBCookie( "AMO chkRasCom", Request.Form("chkRasCom"))
end if
		

				
if bRasCom = "on" then
	bRasCom = true
	sRasCom = "checked"
else
	bRasCom = false
	sRasCom = ""
end if
		
if bGpsyCom = "on" then
	bGpsyCom = true
	sGpsyCom = "checked"
else
	bGpsyCom = false
	sGpsyCom = ""
end if

	if nMode <> 1 Then
		bDisplayHideromMOL = GetDBCookie( "AMO chkIncludeHidefromMOL")
		if bDisplayHideromMOL ="" then
			'default this check box to be checked
			bDisplayHideromMOL = "on"
		end if 
		
	else
		bDisplayHideromMOL = Request.Form("chkIncludeHidefromMOL")
		if bDisplayHideromMOL = "on" then
			Call SaveDBCookie( "AMO chkIncludeHidefromMOL", Request.Form("chkIncludeHidefromMOL"))
		else
			Call SaveDBCookie( "AMO chkIncludeHidefromMOL", "off")
		end if 
	end if
	
	if bDisplayHideromMOL = "on" then
		bDisplayHideromMOL = true
		sDisplayHideromMOL = "checked"
	else
		bDisplayHideromMOL = false
		sDisplayHideromMOL = ""
	end if
	
	if nMode <> 1 Then
		bDisplayHideromSCM = GetDBCookie( "AMO chkIncludeHidefromSCM")
	else
		bDisplayHideromSCM = Request.Form("chkIncludeHidefromSCM")
		Call SaveDBCookie( "AMO chkIncludeHidefromSCM", Request.Form("chkIncludeHidefromSCM"))

	end if
	
	if bDisplayHideromSCM = "on" then
		bDisplayHideromSCm = true
		sDisplayHideromSCM = "checked"
	else
		bDisplayHideromSCM = false
		sDisplayHideromSCM = ""
	end if
	
	if nMode <> 1 Then
		bDisplayHideromSCL = GetDBCookie( "AMO chkIncludeHidefromSCL")
	else
		bDisplayHideromSCL = Request.Form("chkIncludeHidefromSCL")
		Call SaveDBCookie( "AMO chkIncludeHidefromSCL", Request.Form("chkIncludeHidefromSCL"))

	end if
	
	if bDisplayHideromSCL = "on" then
		bDisplayHideromSCL = true
		sDisplayHideromSCL = "checked"
	else
		bDisplayHideromSCL = false
		sDisplayHideromSCL = ""
	end if
	

	sModuleCategoryHTML = ""
	sStatusHTML = ""
	sBusSegHTML = ""
	sOwnersHTML = ""
	
	 '6/29/16 - VHarris - PBI 17487/Task 21005 - Create AMOFeatureCategory cookie for pages using Feature Category
    If GetDBCookie("AMO lbxAMOFeatureCategoryStored") = "" Then
        'If User never used Feature Category Stored, create cookie
        Call SaveDBCookie( "AMO lbxAMOFeatureCategoryStored", "1")
    End If

	if Request.Form("lbxSelectedCategory") = "" and GetDBCookie("AMO lbxAMOFeatureCategoryStored") = "1"  and Request.Form("emptycategory") <> "1" then
		'get the cookie. If we didn't get it default it
		set oRsCategories = GetDBCookieSet( "AMO lbxAMOFeatureCategoryIDs")
        
        'store the cookie
	    Call SaveDBCookieSet( "AMO lbxAMOFeatureCategoryIDs", "", oRsCategories)
	else
        'get the selected categories: ---
		nCategoriesSelected = Request.Form("lbxSelectedCategory")
        
        'store the selected category in cookie table: ---
	    Call SaveDBCookieSet( "AMO lbxAMOFeatureCategoryIDs", nCategoriesSelected, "")

       'get the cookie. If we didn't get it default it
		set oRsCategories = GetDBCookieSet( "AMO lbxAMOFeatureCategoryIDs")
	end if	
	Call SaveDBCookie( "AMO lbxAMOFeatureCategoryStored", "1")
	
	if not oRsCategories is nothing then
		if oRsCategories.RecordCount > 0 then
			nCategoriesSelected = replace(RecordsetToDelimitedString(oRsCategories, ",", "Value"), " ", "")
		else
			nCategoriesSelected = ""
		end if
	end if
	strKeyWord = ""
	
    '9/21/2016 - Harris, Valerie - Added 'Global Series Date' checkbox back and set default value to off for faster page loading
    if nMode <> 1 Then
	    bDisplayIncludeBlankGSD = GetDBCookie( "AMO chkIncludeBlankGSD")
	    if bDisplayIncludeBlankGSD ="" then
		    'default this check box to be unchecked
		    bDisplayIncludeBlankGSD = "off"
	    end if 		
    else
	    bDisplayIncludeBlankGSD = Request.Form("chkIncludeBlankGSD")
	    if bDisplayIncludeBlankGSD = "on" then
		    Call SaveDBCookie( "AMO chkIncludeBlankGSD", Request.Form("chkIncludeBlankGSD"))
	    else
		    Call SaveDBCookie( "AMO chkIncludeBlankGSD", "off")
	    end if 
    end if
	
    if bDisplayIncludeBlankGSD = "on" then
	    bDisplayIncludeBlankGSD = true
	    sDisplayIncludeBlankGSD = "checked"
    else
	    bDisplayIncludeBlankGSD = false
	    sDisplayIncludeBlankGSD = ""
    end if

	sErr = GetFilterInfo( nCategoriesSelected, sStatusIDs, sBusSegIDs, strEOLDate,_
			 sModuleCategoryHTML, sStatusHTML, oRsAMOModules, _
			  nNumTotalModules, sBusSegHTML, bGpsyCom, bRasCom, bRAS ,_
			  bDisplayHideromMOL, bDisplayHideromSCM, bDisplayHideromSCL, strMRFromDate, strMRToDate, _
			  strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, _
			  strchkMRBlank, strchkMRBlank, strchkMRBlank, strRasObsoleteFromDate, strRasObsoleteToDate, sOwnersHTML, sOwnerIDs, bDisplayIncludeBlankGSD, strKeyWord, "NewBus")
				'4/18/2005 since strchkMRBlank value is used for all these three checkboxes, 
		' but teh 'complete' tab has it's own situation, so pass strchkMRBlank here to all three


'if oRsAMOModules.RecordCount > 0 then
	'user has to be in owner group in order to do AMO updates
	'if not IsUserInGroup(oRsAMOModules("GroupID").value) then
		'bAMOCreate = False
		'bAMOUpdate = False
		'bAMODelete = False
		'bAMOView = False
	'end if
'end if

function isIdBelong(byval dIds, byval sIds)
	dim i, arrIds, bFlag
	
	if Trim(dIds) <> "" then
		arrIds = Split(Trim(dIds), ",")
		bFlag = False
		for i = 0 to UBound(arrIds)
			if Trim(arrIds(i)) <> "" and instr(Trim(sIds), "," & Trim(arrIds(i)) & ",") > 0 then
				bFlag = True
				Exit For
			end if
		Next
	end if
	isIdBelong = bFlag	 
end function	
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
<title><%=sHeader%> - Platforms</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<!--[if lte IE 8]>
    <link rel="stylesheet" type="text/css" href="../style/amo_ie8.css" />
<![endif]-->
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../library/scripts/general.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<script type="text/javascript" src="../library/scripts/popitmenu.js"></script> 
<SCRIPT type="text/javascript">
<!--
function ClickEvent(evt) {
	var objUnknown;
	var ModuleID, CategoryID, PlatformID, DivisionID;
	if (!evt) evt = window.event;
	objUnknown = evt.srcElement? evt.srcElement : evt.target;	
	
	if ((objUnknown.tagName.toUpperCase() != "TD") && (objUnknown.tagName.toUpperCase() != "IMG"))
		return;

	if (objUnknown.tagName.toUpperCase() == "IMG") {
		// make object the parent TD
		objUnknown = objUnknown.parentNode; //parentElement
	}

	//Extract ModuleID and CategoryID from 'm123c123'
	ModuleID = objUnknown.parentNode.getAttribute("mid");
	CategoryID = objUnknown.parentNode.getAttribute("cid");

	switch (objUnknown.getAttribute("name")){
		case 'prop': //Properties column
			ShowModuleProperties(ModuleID)
			break;
		case 'o': //Option column
			break;
	    case 'plat': //Platform column
			//Extract PlatformID
		    PlatformID = objUnknown.getAttribute("pid");
		  
			if (objUnknown.innerHTML == "&nbsp;") {	
				// add checkmark
				cM_ChangePlat(evt, ModuleID, PlatformID, 1, objUnknown);
			} else {
				// remove checkmark
				cM_ChangePlat(evt, ModuleID, PlatformID, 0, objUnknown);
			}
			break;
		case 'comparability': //comparability column
			//Extract DivisionID
			DivisionID = objUnknown.getAttribute("pid");
			if (objUnknown.innerHTML == "&nbsp;") {	
				// add checkmark
				cM_ChangeComparability(evt, ModuleID, DivisionID, 1, objUnknown);
			} else {
				// remove checkmark
				cM_ChangeComparability(evt, ModuleID, DivisionID, 0, objUnknown);
			}
			break;
	}
}

function cM_ChangePlat(evtobj, ModuleID, PlatformID, SetStatus, objUnknown) {
	var menu1=new Array()
	var wide = 50;
	var myid = objUnknown.id;

	if (SetStatus == 1) {
		menu1.push("<a href='../nj.asp' onClick=\"ChangePlatStatus(" + ModuleID + "," + PlatformID + ", 165 ,'" + myid + "'); return false;\">A</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangePlatStatus(" + ModuleID + "," + PlatformID + ", 164 ,'" + myid + "'); return false;\">N</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangePlatStatus(" + ModuleID + "," + PlatformID + ", 167 ,'" + myid + "'); return false;\">N/A</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangePlatStatus(" + ModuleID + "," + PlatformID + ", 166 ,'" + myid + "'); return false;\">O</a>");
	} else {
		menu1.push("<a href='../nj.asp' onClick=\"ChangePlatStatus(" + ModuleID + "," + PlatformID + "," + SetStatus + ",'" + myid + "'); return false;\">Remove from Platform</a>");
		wide = 150;
	}

	showmenu(evtobj, menu1.join(""), wide+'px');
}

function cM_ChangeComparability(evtobj, ModuleID, DivisionID, SetStatus, objUnknown) {
	var menu1=new Array()
	var wide = 50;
	var myid = objUnknown.id;

	if (SetStatus == 1) {
		menu1.push("<a href='../nj.asp' onClick=\"ChangeComparabilityStatus(" + ModuleID + "," + DivisionID + ", 165 ,'" + myid + "'); return false;\">A</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangeComparabilityStatus(" + ModuleID + "," + DivisionID + ", 164 ,'" + myid + "'); return false;\">N</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangeComparabilityStatus(" + ModuleID + "," + DivisionID + ", 167 ,'" + myid + "'); return false;\">N/A</a>");

		menu1.push("<a href='../nj.asp' onClick=\"ChangeComparabilityStatus(" + ModuleID + "," + DivisionID + ", 166 ,'" + myid + "'); return false;\">O</a>");
	} else {
		menu1.push("<a href='../nj.asp' onClick=\"ChangeComparabilityStatus(" + ModuleID + "," + DivisionID + "," + SetStatus + ",'" + myid + "'); return false;\">Remove from Compatibility</a>");
		wide = 180;
	}

	showmenu(evtobj, menu1.join(""), wide+'px');
}

//*****************************************************************
//Function:     ChangePlatStatus();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangePlatStatus(ModuleID, PlatformID, SetStatus, myid) {
	var objUnknown = document.getElementById(myid);
	var strCheckmark;
    var ajaxurl = "";
    var strValue = ""; //leave empty

	if (SetStatus == 164 || SetStatus == 165 || SetStatus == 166 || SetStatus == 167) {	
		strCheckmark = "<input onKeyPress='return checkEnter(this, event);' " +
						"onBlur='javascript:getDatePlat(" + ModuleID + "," + PlatformID + 
						"," + SetStatus + ",event)' name='txtADate' id='txtADate' maxlength=20 size=10 value=''>"
				
		objUnknown.innerHTML = strCheckmark 
		document.getElementById("txtADate").focus()
		// highlight the field
		objUnknown.className = "clsAMO_ChangedCell";
	} else {
		//var objRS = RSGetASPObject("AMO_RS.asp");
	    var fullname = "<%= session("FullName") %>";

	    ajaxurl = "AMO_SetPlatform.asp?RGS=1&ModuleID=" + ModuleID + "&PlatformID=" + PlatformID + "&SetStatus=" + SetStatus + "&Value=" + strValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
		//var objResult = objRS.setPlatform("<%=Application("REPOSITORY")%>", ModuleID, PlatformID, SetStatus, " ", "<%= session("FullName") %>");
		
		$.ajax({
		    url: ajaxurl,
		    type: "GET",
		    async: false,
		    success: function (data) {
		        errormsg = data;                     
		    },
		    error: function (xhr, status, error) {
		        errormsg = error; 
		        erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";                  
		    }
		})

		if (errormsg == "success") {
			// clear the checkmark
			objUnknown.innerHTML = "&nbsp;"
			// highlight the field
			objUnknown.className = "clsAMO_ChangedBlankCell";
		}
	}	
	hidemenu();
}

//*****************************************************************
//Function:     ChangeComparabilityStatus();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeComparabilityStatus(ModuleID, DivisionID, SetStatus, myid) {
	var objUnknown = document.getElementById(myid);
	var strCheckmark
	var ajaxurl = "";

	//var objRS = RSGetASPObject("AMO_RS.asp");
	var fullname = "<%= session("FullName") %>";
	ajaxurl = "AMO_SetComparability.asp?RGS=1&ModuleID=" + ModuleID + "&DivisionID=" + DivisionID + "&SetStatus=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
	//var objResult = objRS.setComparability("<%=Application("REPOSITORY")%>", ModuleID, DivisionID, SetStatus, " ", "<%= session("FullName") %>");
	
    $.ajax({
		url: ajaxurl,
		type: "GET",
		async: false,
		success: function (data) {
		    errormsg = data;                     
		},
		error: function (xhr, status, error) {
		    errormsg = error; 
		    erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";                  
		}
	})

	if (errormsg == "success") {
		if (SetStatus == 164 || SetStatus == 165 || SetStatus == 166 || SetStatus == 167) {	
			switch (SetStatus){
					case 164: 
						objUnknown.innerHTML = "<font color=blue>N<\/font>";	
						break;
					case 165: 
						objUnknown.innerHTML = "<font color=blue>A<\/font>";	
						break;	
					case 166: 
						objUnknown.innerHTML = "<font color=blue>O<\/font>";
						break;
					case 167: 
						objUnknown.innerHTML = "<font color=blue>N/A<\/font>";
						break;
				}
			// highlight the field
			objUnknown.className = "clsAMO_ChangedCell";
		} else {
			// clear the checkmark
			objUnknown.innerHTML = "&nbsp;"
			// highlight the field
			objUnknown.className = "clsAMO_ChangedBlankCell";
		}
	}
	hidemenu();
}

//*****************************************************************
//Function:     getDatePlat();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function getDatePlat(ModuleID, PlatformID, SetStatus, evt) {
	if (!evt){ 
        evt = window.event;
    }
	objUnknown = evt.srcElement? evt.srcElement : evt.target;	
	var theParent = objUnknown.parentNode; //parentElement
	var asObject;
	var ajaxurl = "";
	var strValue = ""; //leave empty
		
	if (objUnknown.id == "txtADate") {
	    var NewValue = objUnknown.value;
		
		if ( NewValue == "" ){ 
			NewValue = " ";
		} else {
			if (!checkDate (objUnknown, "", true)) {
				return false;		
			}
		}
        strValue = NewValue;
		
		//var objRS = RSGetASPObject("AMO_RS.asp");

        var fullname = "<%= session("FullName") %>";
        ajaxurl = "AMO_SetPlatform.asp?RGS=1&ModuleID=" + ModuleID + "&PlatformID=" + PlatformID + "&SetStatus=" + SetStatus + "&Value=" + strValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
		//var objResult = objRS.setPlatform("<%=Application("REPOSITORY")%>", ModuleID, PlatformID, SetStatus, NewValue, "<%= session("FullName") %>");
		
        $.ajax({
            url: ajaxurl,
            type: "GET",
            async: false,
            success: function (data) {
                errormsg = data;
            },
            error: function (xhr, status, error) {
                errormsg = error;
                erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
            }
        })

		if (errormsg == "success") {
				switch (SetStatus){
					case 164: 
						theParent.innerHTML = "<font color=blue>" + NewValue + "(N)<\/font>";	
						break;
					case 165: 
						theParent.innerHTML = "<font color=blue>" + NewValue + "(A)<\/font>";	
						break;	
					case 166: 
						theParent.innerHTML = "<font color=blue>" + NewValue + "(O)<\/font>";
						break;
					case 167: 
						theParent.innerHTML = "<font color=blue>" + NewValue + "(N/A)<\/font>";
						break;
				}
		}			
	}
}

function lbxGo_onchange() {
	// if no status check boxes checked, then check all
	var bNoneChecked = 0;
	var oObject = document.getElementsByName("chkStatus")
	
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoneChecked = 1;
				break;
			}
		}
		if (bNoneChecked == 0) {
			// none were checked, go through and check them all
			for (i = 0; i < oObject.length; i++) {
				oObject[i].checked = true;
			}
		}
	}
	// if no business segment check boxes checked, then check all
	bNoneChecked = 0;
	oObject = document.getElementsByName("chkBusSeg")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoneChecked = 1;
				break;
			}
		}
		if (bNoneChecked == 0) {
			// none were checked, go through and check them all
			for (i = 0; i < oObject.length; i++) {
				oObject[i].checked = true;
			}
		}
	}
	if (!checkDate(thisform.txtDisFromDate, "Show options with End of Manufacturing (EM)", true)) {
	    return false;
	}
	if (!checkDate(thisform.txtDisToDate, "Show options with End of Manufacturing (EM)", true)) {
	    return false;
	}
	if (!checkDate(thisform.txtMRFromDate, "Show options with PHweb (General) Availability (GA)", true)) {
	    return false;
	}
	if (!checkDate(thisform.txtMRToDate, "Show options with PHweb (General) Availability (GA)", true)) {
	    return false;
	}
	if (!checkDate(thisform.txtCPLFromDate, "Show options with CPL Date", true)) {
	    return false;
	}
	if (!checkDate(thisform.txtCPLToDate, "Show options with CPL Date", true)) {
	    return false;
	}
		
	SelectAll(thisform.lbxSelectedCategory);
	if (thisform.lbxSelectedCategory.value == "") {
		alert("Please select at least one AMO Category before proceed....");
		return false;
	} 	

	if (thisform.lbxSelectedCategory.selectedIndex == -1) {
	    thisform.emptycategory.value = "1"
	}

	thisform.action = "AMO_Platforms.asp?nMode=1";	
	thisform.submit();
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
WriteTabs "Products"

if sErr <> "" then
	Response.Write sErr
else
	%>
	<div ID=erroutputArea></div>
	<TABLE border=0 cellPadding=1 cellSpacing=1 width=100%>
		<% Call WriteCategoryFilter("Platforms", sModuleCategoryHTML, sStatusHTML, sBusSegHTML, strEOLDate,_
		 bRAS, sGpsyCom, sRasCom, sDisplayHideromMOL, sDisplayHideromSCM, sDisplayHideromSCL, strMRFromDate, strMRToDate, _
		 strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, strchkMRBlank, strchkCPLBlank, strchkDisBlank, _
		 strRasObsoleteFromDate, strRasObsoleteToDate, nCategoriesSelected, sOwnersHTML, sDisplayIncludeBlankGSD) %>
	</table>

	<TABLE id="tblAMOList" border=0 cellPadding="1" cellSpacing="1">
		<colgroup></colgroup>
		<%
		if nNumTotalModules <= 0 then
			Response.Write "<tr><td>"
			Response.Write "No AMO options have been found for the above filter.<br><br><br>" 
			Response.Write "</td></tr>" & vbCrLf
			Response.Write "<tr><td>"
		else
			%>
			<tr><td align=left>	
			<%
			WritePlatformUpdateableGridHTML oRsAMOModules, sBusSegIDs, sDivisionIDs, sGroupIDs
			%>
			</td></tr>
			<tr><td>
			<%
		end if
		%>
	<div id="errorreport" NAME="errorreport"></div>

	</td></tr>
</table>

<table>
	<tr>
		<td><%
			if oRsAMOModules.RecordCount > 0 then
				if oRsAMOModules.RecordCount = 1 then
					response.write "1 Option displayed"
				else
					response.write oRsAMOModules.RecordCount & " Options displayed"
				end if
			end if
			%></td>
	</tr>
	<tr>
		<td bgcolor='#FFFF99'>Highlighted Cell = Value changed. A Complete status resets.</td>
	</tr>
</table>
	<%
	set oSvr = nothing
	oRsAMOModules.Close
	set oRsAMOModules = nothing
end if 'no error 
%>
<input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</FORM>
<%

%>
<input type="hidden" name="emptycategory" Value="">
</BODY>
</HTML>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
    //hide show menu if no selection is made: ---
    $(document).on('click', function(e) {
        clearhidemenu();
    })
    //resize the height of vertical table headers to match height:
    $(".rotate-text").each(function() {
        var $this = $(this),
            child = $this.children(":first");
        $this.css("height", function() {
            return (child[0].getBoundingClientRect().height + 35);
        });
    });
</script>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->