<%@ Language=VBScript %>
<% Option Explicit %>

<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/Groups.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<%

dim oSvr, oErr, sErrHTML, sCategoryId, sRuleDescription, sMin, sMax
dim nMode, sDivisionIds, sRuleId, sReturnValue, sCategory, strError		
		
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
		set oSvr = New ISAMO
		
		nMode = Request.QueryString("Mode")
		sCategory = Request.QueryString("nCategory")
		
		if nMode = 1 then
			
			sCategoryId = Request.Form("lbxCategory")
			sRuleId = Request.QueryString("RuleId")
			sRuleDescription = Request.Form("txtRulesDescription")
			sMin = Request.Form("txtRulesMin")
			sMax = Request.Form("txtRulesMax")
			sDivisionIds = Request.Form("lbxSelectedDivision")

			sReturnValue = oSvr.UpdateCategoryRules(Application("REPOSITORY"), clng(sCategoryId), clng(trim(sRuleId)) , sRuleDescription, clng(sMin), clng(sMax), sDivisionIds)
		
			if sReturnValue <> "True" then
				strError = "Error updating Category Rules. AMO_CategoryRulesSave.asp"
	            Response.Write(strError)
                Response.End()
			elseif clng(sReturnValue) = 0 then
				sErrHTML = "Cannot create/update  rule, there are already a rule exists with the same category and business segment."
			else
				Response.Redirect "AMO_OptionCategoryRule.asp?Mode=1&CategoryID=" & sCategoryId 
			end if
			
		else
		
			sCategoryId = Request.QueryString("CategoryId")
			sRuleId = Request.QueryString("RuleId")
			
			strError = oSvr.RemoveCategoryRules(Application("REPOSITORY"), clng(trim(sRuleId)) )
		
			if strError <> "True" then
				strError = "Error updating Category Rules. AMO_CategoryRulesSave.asp"
	            Response.Write(strError)
                Response.End()
			else
				Response.Redirect "AMO_OptionCategoryRule.asp?Mode=1&CategoryID=" & sCategoryId 
			end if
		
		
		end if
        
%>

<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
</HEAD>
<BODY bgcolor="#FFFFFF">
<%

%>


<%	if sErrHTML <> "" then %>
	<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td><strong><%=sErrHTML%></strong></td> 
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td><a href='AMO_OptionCategoryRule.asp?Mode=1&CategoryID=<%=sCategoryId%>&RuleID=<%=sRuleId%>'>Back To Rules Search Page</a></td> 
		</tr>
		<tr><td><hr></td></tr>
	</TABLE>
<%
end if

%>
</FORM>
</BODY>
</HTML>
		
