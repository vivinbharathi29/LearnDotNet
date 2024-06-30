<%@  Language=VBScript %>
<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/Overview.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/Cookies.inc" -->

<%
dim oSvr, oErr, oRsProductfamily, sErr
SessionValidation2

'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO

set oRsProductfamily = oSvr.AVAMO_AllProductFamily(Application("REPOSITORY"))

if oRsProductfamily is nothing then
    sErr = "Missing required parameters.  Unable to complete your request."
	Response.Write(sErr)
	Response.End()
end if
%>

<html>
<head>
<title>AVAMO Matrix Report - Filter </title>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<script language="JavaScript" src="../library/scripts/formChek.js"></script>
<script language="JavaScript" src="../library/scripts/calendar.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

-->
</script>
</head>

<body bgcolor="#ffffff">
<% %>
<%= 'BuildHelpCentered( "AVAMO Matrix Report - Filter", "","640") %>

<form NAME="form1" target = "_blank" action="AVAMO_Matrix_ExcelReport.asp" method="post" >

<table cellSpacing="0" cellPadding="0" width="640" align="center" border="0">
	<tr height="21" bgcolor="lightsteelblue">
		<td><b>1.&nbsp;&nbsp;Select Platform Product Family:</b><br></td>
	</tr>
	<tr>
		<td><br><SELECT id=cboPdfamily name=cboPdfamily>
			<OPTION value="">All</OPTION>
			<%if  oRsProductfamily.recordcount > 0 then 
				do while not oRsProductfamily.eof%>
				
				<OPTION value="<%=oRsProductfamily("ProductFamily")%>"><%=oRsProductfamily("ProductFamily")%></OPTION>
				
				
			<%	oRsProductfamily.movenext
				loop
			
				end if				%>
				
</SELECT></td>
	</tr>
	
	

	<tr height="21">
		<td><hr></td>
	</tr>
	<tr height="21">
		<td align="center"><input id="display" type="submit" value="Display Report" name="display"></td>
	</tr>
</table>

</form>
<%  %>
</body>
</html>

