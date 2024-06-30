<%@ Language=VBScript %>
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

dim strError, strComment, sHeader, strField, strCommentField, strReturnURL, sErr
dim intID
dim oSvr, oErr
dim stab
stab= Request.QueryString("stab")
strError = ""
sHeader = "Error Saving Comment"

if Request.Form("ID") = "" then
	sErr = "Missing required parameters.  Unable to complete your request. AMO_SaveComment.asp"
	Response.Write(sErr)
	Response.End()
else
	intID = clng(Request.Form("ID"))
end if

if Request.Form("Field") = "" then
	sErr = "Missing required parameters.  Unable to complete your request. AMO_SaveComment.asp"
	Response.Write(sErr)
	Response.End()
else
	strField = Request.Form("Field")
end if

if strError = "" then
	strComment = Request.form("txtComment")
	
	select case strField
		case "rc" 'regional comment
			strCommentField = "regioncomment"
			'strReturnURL = "AMO_Localization.asp"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "ctr" 'comment to ras
			strCommentField = "infocomment"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "cfr" 'comment from ras
			strCommentField = "rascomment"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "dsc" 'long description
			strCommentField = "longdescription"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "ord" 'Order Instructions
			strCommentField = "orderinstructions"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "reldes" 'Replacement AV Description
			strCommentField = "replacementavdescription"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case "rdsc" 'Rules Description
			strCommentField = "ruledescription"
			strReturnURL = "AMO_ModuleList" & stab &".asp" 
		case else
			sErr = "Missing required parameters.  Unable to complete your request."
		    Response.Write(sErr)
		    Response.End()
	end select
end if

if strError = "" then
	'save the comment
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	sErr = oSvr.SaveFieldValue(Application("REPOSITORY"), clng(intID), strCommentField, strComment, session("FullName"), session("AMOUserID"))
	if sErr <> "True" then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
	end if
	set oSvr = nothing
end if

if strError = "" then
	Response.Redirect strReturnURL
end if
%>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>
</HEAD>
<BODY LANGUAGE=javascript>
<FORM name=thisform method=post>
<%
'Response.Write BuildHelp(sHeader, "")

Response.Write strError 


%>
</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->