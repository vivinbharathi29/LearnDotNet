<%@ Language=VBScript %>
<% OPTION EXPLICIT 
Server.ScriptTimeout = 6000 %>
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->
<%
Call ValidateSession

dim sHeader, sHelpFile, sErr, sStatusIDs, strEOLDate
dim sModuleCategoryHTML, sStatusHTML, sBusSegIDs, sBusSegHTML
dim nCategoryID, nNumTotalModules, sGpsyCom, sRasCom, nMode
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, bRAS, bGpsyCom, bRasCom
dim bAMOGEOCreate, bAMOGEOView, bAMOGEOUpdate, bAMOGEODelete
dim bRASCreate, bRASView, bRASUpdate, bRASDelete
dim oSvr, oErr, strError
dim oRsAMOModules
dim sDivisionIDs, intCount, oRsCreateGroups
dim bDisplayHideromMOL, bDisplayHideromSCM, sDisplayHideromMOL, sDisplayHideromSCM
dim strMRFromDate, strMRToDate, strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate
dim strchkMRBlank, strchkCPLBlank, strchkDisBlank, sGroupIDs


sHelpFile = "../help/HELP_AMO_Localization.asp"
sHeader = "After Market Option List"

'set rsRoles and IRSUserID Session: ----
Call SetPermission()


'get permissions
GetRights2 Application("AMOList"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMOGEO"), bAMOGEOCreate, bAMOGEOView, bAMOGEOUpdate, bAMOGEODelete
GetRights2 Application("AMORAS"), bRASCreate, bRASView, bRASUpdate, bRASDelete


set oRsCreateGroups = GetGroupsForRole2(cstr(Application("AMOLIST")), true, false, false, false, true, false)
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

if Request.Form("chkStatus") = "" and Request.Form("lbxCategory") = "" and Request.Form("txtEOLDate") = "" then
	'Have to handle the situation where no status checkboxes were checked. If filter done, there would still be both other values.
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
	sBusSegIDs = GetDBCookie( "AMO chkBusSeg")
else
	sBusSegIDs = Request.Form("chkBusSeg")
end if
'store the cookie
Call SaveDBCookie( "AMO chkBusSeg", sBusSegIDs )

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
	strchkCPLBlank = GetDBCookie( "AMO chkCPLBlankDate")
else
	strchkCPLBlank = Request.Form("chkCPLBlankDate")
	Call SaveDBCookie( "AMO chkCPLBlankDate", Request.Form("chkCPLBlankDate"))	
end if
Call SaveDBCookie( "AMO chkCPLBlankDate", Request.Form("chkCPLBlankDate"))

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
	else
		bDisplayHideromMOL = Request.Form("chkIncludeHidefromMOL")
		Call SaveDBCookie( "AMO chkIncludeHidefromMOL", Request.Form("chkIncludeHidefromMOL"))
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
sModuleCategoryHTML = ""
sStatusHTML = ""
sBusSegHTML = ""
sErr = GetFilterInfo( nCategoryID, sStatusIDs, sBusSegIDs, strEOLDate,_
			 sModuleCategoryHTML, sStatusHTML, oRsAMOModules, _
			  nNumTotalModules, sBusSegHTML, bGpsyCom, bRasCom, bRAS ,_
			  bDisplayHideromMOL, bDisplayHideromSCM, strMRFromDate, strMRToDate, _
			  strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, _
			  strchkMRBlank, strchkCPLBlank, strchkDisBlank, "", "", "OldBus" )

if oRsAMOModules.RecordCount > 0 then
	'user has to be in owner group in order to do AMO updates
	if not IsUserInGroup(oRsAMOModules("GroupID").value) then
		bAMOCreate = False
		bAMOUpdate = False
		bAMODelete = False
		bAMOView = False
	end if
end if
%>
<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">

<title><%=sHeader%></title>
<script language="JavaScript" type="text/javascript" src="../library/scripts/formChek.js"></script>
<script language="JavaScript" type="text/javascript" src="../library/scripts/calendar.js"></script>
<SCRIPT ID=clientEventHandlersJS type="text/javascript" LANGUAGE=javascript>
<!--
var oPopup = window.createPopup();
var xmldoc = new ActiveXObject("Microsoft.XMLDOM");
var xsldoc = new ActiveXObject("Microsoft.XMLDOM");
var textdecoration = '';

function row_onmouseover(e) {
	var id;
	if (!e) e = window.event;
	var objUnknown = e.srcElement? e.srcElement : e.target;	
	id = objUnknown.id;

	if ((objUnknown.tagName.toUpperCase() == "TD") && (id != 'sd')) {
		//objUnknown.className = 'clsMouseOver';
		textdecoration = objUnknown.style.textDecoration;
		objUnknown.style.cursor = 'pointer';
		objUnknown.style.textDecoration = objUnknown.style.textDecoration + ' underline';
	}
	return true;
}

function row_onmouseout(e) {
	var id;
	if (!e) e = window.event;
	var objUnknown = e.srcElement? e.srcElement : e.target;	
	id = objUnknown.id;

	if ((objUnknown.tagName.toUpperCase() == "TD") && (id != 'sd') ) {
		//objUnknown.className = '';
		objUnknown.style.cursor = 'auto';
		objUnknown.style.textDecoration = textdecoration;
	}
	return true;
}

function ClickEvent(e) {
	var id;
    var sIRSLink = "<%=Application("IRSWebServer")%>";
	var ModuleID, CategoryID, PlatformID, RegionID, GEOID;
	if (!e) e = window.event;
	var objUnknown = e.srcElement? e.srcElement : e.target;	
	id = objUnknown.id;

	if ((objUnknown.tagName.toUpperCase() != "TD") && (objUnknown.tagName.toUpperCase() != "IMG"))
		return;

	if (objUnknown.tagName.toUpperCase() == "IMG") {
		// make object the parent TD
		objUnknown = objUnknown.parentNode; //parentElement
	}

	//if (id == "misc")
	//	id = objUnknown.id1;

	//Extract ModuleID and CategoryID from 'm123c123'
	ModuleID = objUnknown.parentNode.getAttribute("mid");
	CategoryID = objUnknown.parentNode.getAttribute("cid");		
     

	switch (id){
		case 'prop': //Properties column
		    ShowModuleProperties(ModuleID);
			break;
		case 'o': //Short Description column
			//cM_Mod(ModuleID, CategoryID, objUnknown.locked);
			break;
		case 'rc': //Region Comment column
			enterComment(ModuleID, 'rc')
			break;
		case 'gc': // GEO column
			GEOID = objUnknown.getAttribute("gid");
			editGEODate(e, ModuleID, GEOID);
			break;
		case 'reg': // Region column
			//Extract RegionID
			RegionID = objUnknown.getAttribute("rid");
			if (objUnknown.innerHTML == "&nbsp;") {
				// add checkmark
				cM_ChangeRegion(ModuleID, RegionID, 1, objUnknown);
			} else {
				// remove checkmark
				cM_ChangeRegion(ModuleID, RegionID, 0, objUnknown);
			}
			break;
	}
}

function checkEnter(theitem, e) {
	var charCode = e.keyCode
	if (charCode == 13) {
		theitem.blur()
		return false;
	}
}

function editGEODate(e, ModuleID, GEOID) {
	var sHTML
	var objUnknown = e;	// assuming passing from ClickEvent(e)
	if (objUnknown.innerHTML.indexOf("editcell" + ModuleID + GEOID)<0) {
		sHTML = "<input onKeyPress='return checkEnter(this, event)' "
		sHTML += "onBlur='javascript:getGEODate(" + e + "," + ModuleID + "," + GEOID + ",\"" + objUnknown.innerHTML.replace(/&nbsp;/g, "") + "\")' "
		sHTML += "type=text maxlength=10 size=10 value=\"" + objUnknown.innerHTML.replace(/&nbsp;/g, "") + "\"' "
		sHTML += "id='editcell" + ModuleID + GEOID + "' NAME='editcell" + ModuleID + GEOID + "'>";
		objUnknown.innerHTML = sHTML
		document.getElementById("editcell" + ModuleID + GEOID).focus()
	}
}
//*****************************************************************
//Function:     getGEODate();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function getGEODate(e, ModuleID, GEOID, OldValue) {
	var objUnknown = e;	
	var Parent = objUnknown.parentNode; //parentElement
	var ajaxurl = "";
	var strValue = "";

	if (objUnknown.id == "editcell" + ModuleID + GEOID) {
	    if (!checkDate (objUnknown, "GEO Date", true)){
	        return false;
	    }

	    var NewValue = objUnknown.value;

		if (OldValue == NewValue) { // nothing changed
		    if (NewValue == '') {
		        Parent.innerHTML = '&nbsp;';
		    } else {
		        Parent.innerHTML = objUnknown.value;
		    }
		} else { // something changed, save the data
		    //var objRS = RSGetASPObject("AMO_RS.asp");
		    strValue = NewValue;
            
            var fullname = "<%= session("FullName") %>";
            ajaxurl = "AMO_SetGEODate.asp?RGS=1&ModuleID=" + ModuleID + "&GEOID=" + GEOID + "&Value=" + strValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
		    //var objResult = objRS.setGEODate("<%=Application("REPOSITORY")%>", ModuleID, GEOID, NewValue, "<%= session("FullName") %>");
	
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
		        // highlight the field
		        Parent.className = "clsAMO_ChangedCell";
		        if (NewValue == '') {
		            Parent.innerHTML = '&nbsp;'
		        } else {
		            Parent.innerHTML = objUnknown.value;
		        }
		        // update status to 'In Process'
		        var asObject = document.all.item("as" + ModuleID);
		        if (asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled"){ 
		            asObject.innerHTML = "In Process";
		        } else {
		            erroutputArea.innerHTML = "<p><font color=red>" + errormsg + "<\/font><\/p>";
		        }
		    }
		}
	}
}

function enterComment(ModuleID, Field) {
	thisform.ID.value = ModuleID
	thisform.Field.value = Field
	thisform.action = "AMO_AddComment.asp";
	thisform.submit ();
}

function cM_Mod(ModuleID, CategoryID, bLocked) {
	// The variables "lefter" and "topper" store the X and Y coordinates
	// to use as parameter values for the following show method. In this
	// way, the popup displays near the location the user clicks. 
	var lefter = event.clientX;
	var topper = event.clientY;
	var popupBody;
    var sIRSLink = "<%=Application("IRSWebServer")%>";
	
	popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative;TOP:0px\">"; 
	popupBody= popupBody+"<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'\"onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	<% if bAMOUpdate then %>
	if (bLocked == 1) {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ShowModuleProperties(" + ModuleID + ")'\" >&nbsp;&nbsp;View Option Properties<\/SPAN><\/FONT><\/DIV>";
	} else {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ShowModuleProperties(" + ModuleID + ")'\" >&nbsp;&nbsp;View/Modify Option Properties<\/SPAN><\/FONT><\/DIV>";
	}
	<% else %>
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ShowModuleProperties(" + ModuleID + ")'\" >&nbsp;&nbsp;View Option Properties<\/SPAN><\/FONT><\/DIV>";
	<% end if %>

	popupBody = popupBody + "</DIV>";
	oPopup.document.body.innerHTML = popupBody; 

	<% if bAMOUpdate then %>
	if (bLocked == 1) {
		oPopup.show(lefter, topper, 180, 18, document.body);
	} else {
		oPopup.show(lefter, topper, 200, 18, document.body);
	}
	<% else %>
	oPopup.show(lefter, topper, 180, 18, document.body);
	<% end if %>
}

function cM_ChangeRegion(ModuleID, RegionID, SetStatus, objUnknown) {
	// The variables "lefter" and "topper" store the X and Y coordinates
	// to use as parameter values for the following show method. In this
	// way, the popup displays near the location the user clicks. 
	var lefter = event.clientX;
	var topper = event.clientY;
	var popupBody;
	
	popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative;TOP:0px\">"; 

	popupBody= popupBody+"<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'\"onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	if (SetStatus == 1) {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeRegionStatus(" + ModuleID + "," + RegionID + "," + SetStatus + "," + objUnknown.sourceIndex + ")'\" >&nbsp;&nbsp;Add to Region <img src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0><\/SPAN><\/FONT><\/DIV>";
	} else {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeRegionStatus(" + ModuleID + "," + RegionID + "," + SetStatus + "," + objUnknown.sourceIndex + ")'\" >&nbsp;&nbsp;Remove from Region<\/SPAN><\/FONT><\/DIV>";
	}

	popupBody = popupBody + "</DIV>";
	oPopup.document.body.innerHTML = popupBody; 

	if (SetStatus == 1) {
		oPopup.show(lefter, topper, 130, 22, document.body);
	} else {
		oPopup.show(lefter, topper, 150, 18, document.body);
	}
}

function ShowModuleProperties(ModuleID, sIRSLink) {
    window.open("/IPulsar/Features/AMOFeatureProperties.aspx?FromModule=1&FeatureID="+ModuleID, "_blank", "resizable=yes,menubar=yes,scrollbars=yes,toolbar=yes");
}

//*****************************************************************
//Function:     ChangeRegionStatus();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeRegionStatus(ModuleID, RegionID, SetStatus, srcIndex) {
    var strCheckmark;
    var ajaxurl = "";

	//var objRS = RSGetASPObject("AMO_RS.asp");
	
    var fullname = "<%= session("FullName") %>";
    ajaxurl = "AMO_SetRegion.asp?RGS=1&ModuleID=" + ModuleID + "&RegionID=" + RegionID + "&SetStatus=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
	//var objResult = objRS.setRegion("<%=Application("REPOSITORY")%>", ModuleID, RegionID, SetStatus, "<%= session("FullName") %>");
	
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
		if (SetStatus == 1) {
			// add the checkmark
		    strCheckmark = "<img onclick='javascript:ClickEvent(event);return true;' ";
			strCheckmark += "id='" + document.all(srcIndex).id + "' title='" + document.all(srcIndex).title + "' ";
			strCheckmark += "rid='" + document.all(srcIndex).rid + "' ";
			strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>";
			document.all(srcIndex).innerHTML = strCheckmark;
			// highlight the field
			document.all(srcIndex).className = "clsAMO_ChangedCell";
		} else {
			// clear the checkmark
		    document.all(srcIndex).innerHTML = "&nbsp;";
			// highlight the field
			document.all(srcIndex).className = "clsAMO_ChangedBlankCell";
		}
	    // update status to 'In Process'
		var asObject = document.all.item("as" + ModuleID);
		if (asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled"){ 
			asObject.innerHTML = "In Process";
	    } else {
		    erroutputArea.innerHTML = "<p><font color=red>" + errormsg + "<\/font><\/p>";
	    }
	oPopup.hide();
}

function lbxGo_onchange() {
	// if no status check boxes checked, then check all
	var bNoneChecked = 0;
	var oObject = document.getElementsByName("chkStatus")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoneChecked = 1;
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
			}
		}
		if (bNoneChecked == 0) {
			// none were checked, go through and check them all
			for (i = 0; i < oObject.length; i++) {
				oObject[i].checked = true;
			}
		}
	}
	if (!checkDate (thisform.txtDisFromDate, "Show options with End of Manufacturing (EM)", true))
		return false;
	if (!checkDate (thisform.txtDisToDate, "Show options with End of Manufacturing (EM)", true))
		return false;
	if (!checkDate (thisform.txtMRFromDate, "Show options with PHweb (General) Availability (GA)", true))
		return false;
	if (!checkDate (thisform.txtMRToDate, "Show options with PHweb (General) Availability (GA)", true))
		return false;
	if (!checkDate (thisform.txtCPLFromDate, "Show options with Select Availability (SA)", true))
		return false;
	if (!checkDate (thisform.txtCPLToDate, "Show options with Select Availability (SA)", true))
		return false;

	thisform.action = "AMO_Localization.asp?nMode=1";	
	thisform.submit();
}
//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="#FFFFFF">
<!-- #include file="../library/includes/popup.inc" -->
<FORM NAME=thisform method=post>
<%

Response.Write "<br>"
WriteTabs "Localization"

if sErr <> "" then
	Response.Write sErr
else
	%>
	<!--//<script type="text/javascript" language="JavaScript">
	  RSEnableRemoteScripting("_SCRIPTLIBRARY");
	</script>//-->
	<div ID=erroutputArea></div>
	<TABLE border=0 cellPadding=1 cellSpacing=1 width=100%>
		<% Call WriteCategoryFilter( "Localization", sModuleCategoryHTML, sStatusHTML, sBusSegHTML, strEOLDate,_
		 bRAS, sGpsyCom, sRasCom, sDisplayHideromMOL, sDisplayHideromSCM, strMRFromDate, strMRToDate, _
		 strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, strchkMRBlank, strchkCPLBlank, strchkDisBlank, "", "") %>

	</table>

	<TABLE border=0 cellPadding=1 cellSpacing=1 width="100%">
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
			sErr = WriteLocalizationUpdateableGridHTML( oRsAMOModules, sBusSegIDS, sDivisionIDs, sGroupIDs )
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
			if oRsAMOModules.RecordCount > 0 and sErr = "" then
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
<input type="hidden" name="ID" value="">
<input type="hidden" name="Field" value="">
<input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</FORM>
<%


Server.ScriptTimeout = 60
%>
</BODY>
</HTML>
