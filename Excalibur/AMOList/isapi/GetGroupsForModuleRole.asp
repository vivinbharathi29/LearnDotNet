<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/ListBoxRs.inc" -->
<!-- #include file="../library/includes/Overview.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<%
Call ValidateSession

dim sLink, sH1, sDomainUserName, sRoleNumber, sGroupName, sAMORoleNumber, sDelTypeHTML, sHelpfile
dim sErr, strSRPModuleHTML, strAMOModuleHTML
dim nGroupID, nFlag, nMode
dim oSvr, oErr
dim adoRs, adoAMORs
dim bSRPCreate, bSRPUpdate, bSRPView, bSRPDelete, bExcludeIRS, bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim bAMOExcludeIRS
dim SRPgroup_sHTML, AMOgroup_sHTML, nModuleDeliveryType, rdDelType
dim rRoleNumber, rCreate, rView, rUpdate, rDelete, rExcludeIRS
dim bFromAMO, nSelectedGroupID, adoRs2

sHelpFile = "" '/Help/MOL/MOL_Quick_Help.htm"

if Request.QueryString("Mode") <> "" then
	nMode = clng(Request.QueryString("Mode"))
else
	nMode = 0
end if

if Request.QueryString("FromAMO") = "1" then
	bFromAMO = true
else
	bFromAMO =false
end if 

if nMode = 3 then
	sH1 = "Clone an AMO Option"
else
	sH1 = "Create a New Module"
end if


GetRights2 Application("AMOList"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete

sErr = ""
strSRPModuleHTML = ""
strAMOModuleHTML = ""
if nMode <> 3 then	'Clone an AMO Option
	GetRights2 Application("MODULE"), bSRPCreate, bSRPView, bSRPUpdate, bSRPDelete
	'get Module rights and if Create then make list box
	if bSRPCreate then
		set adoRs = GetGroupsForRole(Application("MODULE"), true, false, false, false, true)
		if adoRs is nothing then
			Response.Write("Empty Recordset: adoRs")
		    Response.End()
		else
			'if the user is in Admin role, we return all the groups with create access.  But Suggestion 407 requires that
			'we default to the actual Marketing role the admin user is in.
			set adoRs2 = GetGroupsForRole2(Application("MODULE"), true, false, false, false, true, false)
			if adoRs2 is nothing then
				Response.Write("Empty Recordset: adoRs2")
		        Response.End()
			else
				if (not adoRs2 is nothing and adoRs2.RecordCount > 0) then
					nSelectedGroupID = adoRs2.Fields("GroupID").Value
				else
					nSelectedGroupID = 0 
				end if
				if (not adoRs is nothing and adoRs.RecordCount > 0) then
					if adoRs.RecordCount > 1 then
						strSRPModuleHTML = Lbx_GetHTML2("lbxTempGroupID", false, 1, 300, adoRs, "GroupName", "GroupID", 0)
					else
						strSRPModuleHTML = "<b>" & adoRs("GroupName") & "</b>" & vbCrLF
						strSRPModuleHTML = strSRPModuleHTML & "<input type=hidden name=lbxTempGroupID gn=""" & adoRs("GroupName") & """ value=" & adoRs("GroupID") & ">"
					end if
				end if
				adoRs.Close
				set adoRs = nothing
			end if
		end if
	end if
end if

if sErr = "" then
	'get AMO List rights and if Create then make list box
	if bAMOCreate then
		set adoRs = GetGroupsForRole2(Application("AMOList"), true, false, false, false, true, false)
		if adoRs is nothing then
			Response.Write("Empty Recordset: adoRs")
		    Response.End()
		else
			if (not adoRs is nothing and adoRs.RecordCount > 0) then
				if adoRs.RecordCount > 1 then
					strAMOModuleHTML = Lbx_GetHTML2("lbxTempGroupID", false, 1, 300, adoRs, "GroupName", "GroupID", 0)
				else
					strAMOModuleHTML = "<b>" & adoRs("GroupName") & "</b>" & vbCrLF
					strAMOModuleHTML = strAMOModuleHTML & "<input type=hidden name=lbxTempGroupID gn=""" & adoRs("GroupName") & """ value=" & adoRs("GroupID") & ">"
				end if
			end if
		end if
		adoRs.Close
		set adoRs = nothing
	end if
end if

if sErr = "" then
	sDelTypeHTML = ""
	if nMode <> 3 then
		if bSRPCreate and bAMOCreate and strAMOModuleHTML <> "" and strSRPModuleHTML <> "" then
			if bFromAMO then
				sDelTypeHTML = sDelTypeHTML & "<div style='display:none;'><INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("MODULE") & "> CTO (SRP) </div>" 
				sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("AMOList") & " checked > AMO"
				sDelTypeHTML = sDelTypeHTML & "<div style='display:none;'><INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("MODULE_TECHAV") & "> Technical AV </div>" 
	
			
			else
				sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("MODULE") & " checked > CTO (SRP)" 
				sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("AMOList") & " > AMO"
				sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("MODULE_TECHAV") & " > Technical AV" 
	
			
			end if 
		
		elseif (bSRPCreate or strAMOModuleHTML = "") and strSRPModuleHTML <> "" then
			sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' checked value=" & Application("MODULE") & " > CTO (SRP)" 
			sDelTypeHTML = sDelTypeHTML	& "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' disabled value=" & Application("AMOList") & " > AMO"
			sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' value=" & Application("MODULE_TECHAV") & " > Technical AV" 
		elseif (bAMOCreate or strSRPModuleHTML = "") and strAMOModuleHTML <> "" then
			sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' disabled value=" & Application("MODULE") & " > CTO (SRP)"
			sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' checked value=" & Application("AMOList") & " > AMO"
			sDelTypeHTML = sDelTypeHTML & "<INPUT type='radio' id=rdDelType name=rdDelType  onclick='javascript:rdDelType_onclick()' disabled value=" & Application("MODULE_TECHAV") & " > Technical AV"
		else
			sDelTypeHTML = "You do not have any Create rights"
		end if
	end if
end if
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<title>Validate Groups</title>
<script language="JavaScript" src="../library/scripts/formChek.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function rdDelType_onclick() {
	var bSRP=true, bAMO=true;
  
	<% if nMode <> 3 then %>
	if (frmGetGroups.rdDelType != null) {
		for (var i=0; i<frmGetGroups.rdDelType.length; i++) {
			if (frmGetGroups.rdDelType[i].checked) {
				if (frmGetGroups.rdDelType[i].value == <%= Application("MODULE") %> || frmGetGroups.rdDelType[i].value == <%= Application("MODULE_TECHAV") %>) 
					bAMO=false;
				else if (frmGetGroups.rdDelType[i].value == <%= Application("AMOList") %>) 
					bSRP=false;
			}
		}	
		if (bSRP)
			strInnerHTML = document.getElementById("divSRPgroup").innerHTML;
	<% end if %>
		if (bAMO)
			strInnerHTML = document.getElementById("divAMOgroup").innerHTML;

		var re = /lbxTempGroupID/g;             
		document.getElementById("divgroup").innerHTML = strInnerHTML.replace(re, "lbxGroupID");
	<% if nMode <> 3 then %>
	}
	<% end if %>
	return true;
}

function btnContinue_onclick() {
	var sText, lbx
	
	<% if nMode = 3 then %>
		lbx = frmGetGroups.lbxGroupID;
		sText = "AMO_Properties.asp?nEditLocalization=1&"
		if (lbx.type == 'hidden')
			sText += "nGroupID=" + lbx.value + "&sGroupName=" + lbx.gn;
		else
			sText += "nGroupID=" + lbx.value + "&sGroupName=" + lbx.options[lbx.selectedIndex].text;
		sText += "&<%= Request.Querystring %>"
		frmGetGroups.action = sText
	<% else %>
	if (frmGetGroups.rdDelType != null) {
		lbx = frmGetGroups.lbxGroupID;
		for (var i=0; i<frmGetGroups.rdDelType.length; i++) {
			if (frmGetGroups.rdDelType[i].checked) {
				if (frmGetGroups.rdDelType[i].value == <%= Application("MODULE") %>) {
					sText = "Module_Properties.asp?Mode=1"
				} else if (frmGetGroups.rdDelType[i].value == <%= Application("MODULE_TECHAV") %>) {
					sText = "Module_Properties.asp?Mode=1&TechAV=1"		
				} else if (frmGetGroups.rdDelType[i].value == <%= Application("AMOList") %>) {
					sText = "AMO_Properties.asp?nEditLocalization=1&Mode=1"
				}
				if (lbx.type == 'hidden')
					sText += "&nGroupID=" + lbx.value + "&sGroupName=" + lbx.gn;
				else
					sText += "&nGroupID=" + lbx.value + "&sGroupName=" + lbx.options[lbx.selectedIndex].text;
				frmGetGroups.action = sText
			}
		}	
	}
	<% end if %>
	return true;
}

function window_onload() {
	<% if sErr = "" then %>
	rdDelType_onclick();
	<% end if %>
}
//-->
</SCRIPT>
</head>

<body bgcolor="#FFFFFF" onload="return window_onload()">
<form NAME="frmGetGroups" METHOD="post" action="">
<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<%
if sErr = "" then
	%>
	<tr>
		<td>
			<table border="0" cellPadding="2" cellSpacing="0" width="100%">
				<colgroup width="25%"></colgroup>
				<colgroup width="75%"></colgroup>
				<% if nMode <> 3 then %>
				<tr>
					<td>Select Delivery Type</td>
					<td><%= sDelTypeHTML %></td>
				</tr>
				<% end if %>
				
				<tr style="height: 50px;">
					<td>Select User Group</td>
					<td>
						<DIV id=divSRPgroup name=divSRPgroup style="display:none"><%=strSRPModuleHTML%></DIV>
						<DIV id=divAMOgroup name=divAMOgroup style="display:none"><%=strAMOModuleHTML%></DIV> 
						<DIV id=divgroup name=divgroup></DIV></TD>
				</tr>
				
				<tr>
					<td colSpan=2 align="left">
						<input id="btnContinue" name="btnContinue" type="submit" value="Continue" LANGUAGE=javascript onclick="return btnContinue_onclick()">
					</td>
				</tr>
			</table>			
		</td>
	</tr>
	<%
else
	%>
	<tr>
		<td>
			<table WIDTH="100%" BORDER="0" CELLSPACING="1" CELLPADDING="1">
				<tr>
					<td>
						<%= sErr %>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<%
end if
%>
</table>

</form>
</body>
</html>
