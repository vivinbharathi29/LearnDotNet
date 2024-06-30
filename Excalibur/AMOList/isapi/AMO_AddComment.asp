<%@ Language="VBScript" %>
<% Option Explicit %>
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

dim strComment, strBluePN, strTab, strRedPN
dim sHeader, sHelpFile, strField, strCommentField
dim strShortDesc, strError
dim intID, intStatusID, oRsCreateGroups, sGroupIDs, intCount
dim oSvr, oRs, oErr, strLabel, strLabel2, sDivisionIDs
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, bUpdate
dim bRASCreate, bRASView, bRASUpdate, bRASDelete
dim stab, strKeyWord

sHelpFile = "" '"/Help/MOL/MOL_Quick_Help.asp"
sHeader = "After Market Options List - Comment Error"

'set rsRoles and IRSUserID Session: ----
Call SetPermission()

'get permission
GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMORAS_Permission"), bRASCreate, bRASView, bRASUpdate, bRASDelete
bUpdate = True



set oRsCreateGroups = GetGroupsForRole2(cstr(Session("AMOUserRoleNames")), cstr(Application("AMO_Permission")), true, false, false, false, true, false)
if (oRsCreateGroups is nothing) then
	Response.Write("Empty Recordset: oRsCreateGroups")
	Response.End()
else	
	sGroupIDs = ""
	sDivisionIDs = ""
	For intCount = 0 To oRsCreateGroups.RecordCount-1
		sGroupIDs = "," & sGroupIDs & oRsCreateGroups("GroupID") & ","
		sDivisionIDs = "," & sDivisionIDs & replace(oRsCreateGroups("DivisionIDs"), "|", ",")
		oRsCreateGroups.MoveNext		
	Next
			
end if

if Request.Form("ID") = "" then
	strError = "Error Processing Module Search. AMO_AddComment.asp"
	Response.Write(strError)
    Response.End()
else
	intID = clng(Request.Form("ID"))
end if

if Request.Form("Field") = "" then
	strError = "Error Processing Module Search. AMO_AddComment.asp"
	Response.Write(strError)
    Response.End()
else
	strField = Request.Form("Field")
end if

strTab =Request.QueryString("stab")
stab = strTab
if not len(strTab)> 0 then

	strTab = "All" & "&nbsp;" & "Options"
else
	stab = "_"	& stab
end if 

if strTab="RAS" then
	strTab = "RAS" & "&nbsp;" & "Review"

end if 

strLabel = "Comment"
strLabel2 = "255 maximum characters"

if strError = "" then
	select case strField
		case "rc" 'regional comment
			sHeader = "After Market Options List - Regional Comment"
			strCommentField = "RegionComment"
			'strTab = "Localization"
		case "ctr" 'comment to ras
			sHeader = "After Market Options List - Comment to RAS"
			strCommentField = "InfoComment"
			'strTab = "Options"
		case "dsc" 'Long Description
			sHeader = "After Market Options List - Long Description"
			strCommentField = "LongDescription"
			'strTab = "Options"
			strLabel = "Long Description"
			strLabel2 = "160 maximum characters"
		case "rdsc" 'Rules Description
			sHeader = "After Market Options List - Rules Description"
			strCommentField = "RuleDescription"
			'strTab = "Options"
			strLabel = "Rules Description"
			strLabel2 = "1024 maximum characters"
		case "ord" 'Order Instructions
			sHeader = "After Market Options List - Order Instructions"
			strCommentField = "OrderInstruction"
			'strTab = "Options"
			strLabel = "Order Instructions"
			strLabel2 = "600 maximum characters"
		case "reldes" 'Replacement AV Description
			sHeader = "After Market Options List - Replacement AV Description"
			strCommentField = "ReplacementAVDescription"
			'strTab = "Options"
			strLabel = "Replacement AV Description"
			strLabel2 = "80 maximum characters"
		case "cfr" 'comment from ras
			sHeader = "After Market Options List - Comment from RAS"
			strCommentField = "RASComment"
			'strTab = "Options"
		case else
			strError = "Error Processing Module Search. AMO_AddComment.asp"
			Response.Write(strError)
            Response.End()
	end select
end if

if strError = "" then
    strKeyWord = ""
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	set oRs = oSvr.AMOModule_Search(Application("REPOSITORY"), " and O.FeatureID=" & cstr(intID) & " and O.SCMID = 1 ", strKeyWord, null, null)
	if (oRs is nothing) then
		strError = "Error Processing Module Search. AMO_AddComment.asp"
		Response.Write(strError)
        Response.End()
	else
		if oRs.RecordCount = 0 then
			strError = "Error Processing Module Search. AMO_AddComment.asp"
			Response.Write(strError)
            Response.End()
		else
			strShortDesc = oRs.Fields("ShortDescription").Value
			strBluePN = oRs.Fields("BluePN").value
			strRedPN = oRs.Fields("RedPN").value
			strComment = oRs.Fields(strCommentField).value
			intStatusID = clng(oRs.Fields("AMOStatusID").value)
			
			'No longer using Owned By - if instr(sGroupIDs,"," & Trim(oRs.Fields("GroupID").value) & ",") > 0 and bAMOUpdate then
            if bAMOUpdate then
				bUpdate = True
			else
				bUpdate = False
			end if
			
			
			select case strField
				case "rc" 'regional comment
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "dsc" 'Long Description
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "rdsc" 'Rules Description
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "ord" 'Order Instructions
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "reldes" 'Replacement AV Description
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "ctr" 'comment to ras
				
					bUpdate = isIdBelong(Trim(oRs.Fields("FeatureDivisionIDs").value), sDivisionIDs)
				
					if intStatusID = clng(Application("AMO_RASREVIEW")) or intStatusID = clng(Application("AMO_RASUPDATE")) or intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
				case "cfr" 'comment from ras
					if intStatusID = clng(Application("AMO_DISABLED")) or intStatusID = clng(Application("AMO_OBSOLETE")) then
						bUpdate = False
					end if
			end select

		end if
		oRs.Close
		set oRs = nothing
	end if
	set oSvr = nothing
end if


function isIdBelong(byval dIds, byval sIds)
	dim i, arrIds, bFlag
	
	if Trim(dIds) <> "" then
		arrIds = Split(Trim(dIds), ",")
		bFlag = False
		for i = 0 to UBound(arrIds)
			if Trim(arrIds(i)) <> "" and instr("," & Trim(sIds) & ",", "," & Trim(arrIds(i)) & ",") > 0 then
				bFlag = True
				Exit For
			end if
		Next
	end if
	isIdBelong = bFlag	 
end function

%>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="AMO - Add Comment" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title>AMO - Add Comment</title>
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
<SCRIPT type="text/javascript" ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<% if bUpdate then %>
function Validate() {
	if (thisform.txtComment.value.length > 255) {
		warnInvalid(thisform.txtComment, "Maximum length for the Comment is 255 characters.");
		return false;
	}
	return true;
}

function ValidateDesc() {
	if (thisform.txtComment.value.length > 160) {
		warnInvalid(thisform.txtComment, "Maximum length for long description is 160 characters.");
		return false;
	}
	return true;
}


function ValidateReplacement() {
	if (thisform.txtComment.value.length > 80) {
		warnInvalid(thisform.txtComment, "Maximum length for Replacement AV Descriptiuon is 80 characters.");
		return false;
	}
	return true;
}


function ValidateOrder() {
	if (thisform.txtComment.value.length > 600) {
		warnInvalid(thisform.txtComment, "Maximum length for Order Instruction is 600 characters.");
		return false;
	}
	return true;
}

function ValidateRule() {
	if (thisform.txtComment.value.length > 1024) {
		warnInvalid(thisform.txtComment, "Maximum length for Rules Description is 1024 characters.");
		return false;
	}
	return true;
}

function btnSave_onclick() {

	<%if strField = "dsc" then %>

		if (ValidateDesc()) {
			thisform.action = "AMO_SaveComment.asp?stab=<%=stab%>";
			return true;
		}
		return false;
		
	<%elseif strField = "reldes" then %>
	
		if (ValidateReplacement()) {
			thisform.action = "AMO_SaveComment.asp?stab=<%=stab%>";
			return true;
		}
		return false;
	
	<%elseif strField = "ord" then %>
	
		if (ValidateOrder()) {
			thisform.action = "AMO_SaveComment.asp?stab=<%=stab%>";
			return true;
		}
		return false;
		
	<%elseif strField = "rdsc" then %>
	
		if (ValidateRule()) {
			thisform.action = "AMO_SaveComment.asp?stab=<%=stab%>";
			return true;
		}
		return false;
	
		
   <% else %>
		
		if (Validate()) {
			thisform.action = "AMO_SaveComment.asp?stab=<%=stab%>";
			return true;
		}
		return false;
		
   <%end if%>
}



<% end if %>

function btnCancel_onclick() {
	<% if strField = "rc" then %>
  thisform.action = "AMO_ModuleList<%=stab%>.asp";
	<% else %>
  thisform.action = "AMO_ModuleList<%=stab%>.asp";
	<% end if %>
  return true;
}
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor="#FFFFFF">
<h1 class="page-title"><%=sHeader%></h1>
<FORM name=thisform method=post>
<%
Response.Write ""
WriteTabs strTab
%>
<DIV id=divReturnStatus name=divReturnStatus></DIV>
<%
if strError <> "" then
	Response.Write strError
else
	%>
	<TABLE border=0 width="100%">
	  <TR>
	    <TD width="20%">Option (Short Description)</TD>
	    <TD width="50%"><b><%=strShortDesc%>&nbsp;</b></TD>
			<td width="30%">&nbsp;</td></TR>
	  <TR>
	    <TD width="20%">HP Part Number</TD>
	    <TD width="50%"><b><%=strBluePN%>&nbsp;</b></TD>
			<td width="30%">&nbsp;</td></TR>
		<tr>
			<td width="20%" valign="middle"><%=strLabel%></td>
			<td width="50%"><textarea rows=5 style="width: 100%;" name="txtComment" id="txtComment" <%
			if not bUpdate then response.write "readonly style=""color:gray"""
			%>><%= strComment %></textarea></td>
			<td width="30%"><font size=1><i><%=strLabel2%></i></font></td>
			</tr>

	  <TR>
	    <TD colspan=3 align=left><br>
			<% if bUpdate then %>
				<INPUT id=btnSave name=btnSave type=submit value="Save" LANGUAGE=javascript onClick="return btnSave_onclick()">
			<% end if %>
				<INPUT id=btnCancel name=btnCancel type=submit value="Cancel" LANGUAGE=javascript onClick="return btnCancel_onclick()">
	  	</TD></TR>
	</TABLE>
	
	<INPUT type="hidden" id="ID" name="ID" value=<%= intID %>>
	<INPUT type="hidden" id="Field" name="Field" value=<%= strField %>>
<%end if%>
 
</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->
