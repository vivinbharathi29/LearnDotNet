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


dim irowcount, oSvr, oErr, sErr, sDivisionIds, sRASDiscontinueDate, oRsInvalidData
dim arrColors, sHeader
arrColors = Array("#CCCCCC", "white")

sHeader = "AMO Validate Data"

sRASDiscontinueDate = Request.QueryString("RASDiscontinueDate")
sDivisionIds = Request.QueryString("BusSeg")

set oRsInvalidData = Nothing


'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO
	
set oRsInvalidData = oSvr.AMO_ValidateData(Application("REPOSITORY"), sDivisionIds, sRASDiscontinueDate)
if oRsInvalidData is nothing then
	sErr = "Missing required parameters.  Unable to complete your request."
	Response.Write(sErr)
	Response.End()
end if
%>
<!DOCTYPE html>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
<script type="text/javascript">
    function goBack() {
        window.history.back();
    }
</script>
</HEAD>
<body bgcolor="#FFFFFF">
<FORM name=thisform method=post>
<%
'InsertGlobalNavigationBar_HomeParm(true)
%>
<h2 valign='top' align='left'>AMO Validate Data</h2>
<%
if sErr <> "" then
	Response.Write sErr
else
	%>
	<h4 valign='top' align='left'>AMO Options that have Global Series Config EOL not in between PHweb (General) Availability (GA) and End of Manufacturing (EM):</h4>
	<TABLE border=1 width="100%">
    <TR bgcolor="lightsteelblue" >
        <Th >Short Description</Th>
        <Th >HP PartNo</Th>
        <Th >PHweb (General) Availability (GA)</Th>
        <th >End of Manufacturing (EM)</th>
        <th >Global Series Config EOL</th>
    </TR>
	
	    <%if oRsInvalidData.recordcount> 0 then
		    irowcount=0 
	    do while not oRsInvalidData.eof
		    irowcount = irowcount + 1
		    %>    
		    <TR  bgcolor="<%= arrColors(irowcount mod 2) %>" >
			    <td><%=oRsInvalidData("ShortDescription")%>&nbsp;</td> 
			    <td  align=center><%=oRsInvalidData("BluePartNo")%>&nbsp;</td> 
			    <td  align=center><%=oRsInvalidData("BOMRevADate")%>&nbsp;</td>
			    <td  align=center><%=oRsInvalidData("RASDiscontinueDate")%>&nbsp;</td> 
			    <td  align=center><%=oRsInvalidData("GlobalSeriesDate")%> </td> 
			
		    </tr>
			
		    <%	
			    oRsInvalidData.movenext		
	    loop
				
	    else
		    Response.Write "<tr> <td colspan =3>No invalid AMO Option that has Global Series Config EOL less than PHweb (General) Availability (GA) or greater than End of Manufacturing (EM) was found</td></tr>"
		
	    end if %>
	</TABLE>
    <br/><br/>
    <button type="button" onclick="goBack();">Return to Report</button> 
	<%
	set oRsInvalidData = Nothing
end if


%>

</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->

