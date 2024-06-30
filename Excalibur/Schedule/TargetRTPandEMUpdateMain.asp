<%@ Language=VBScript %>
<%Option Explicit%>

<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  
%>
	
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Update Current Commitment Dates</title>
<SCRIPT id=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdDate_onclick(FieldID, FieldID2) {
	var strID;
	var oldValue = window.frmUpdate.elements(FieldID).value;
    // remove FCS from all areas - task 20243
	strID = window.showModalDialog("../mobilese/today/caldraw1.asp",FieldID,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof (strID) == "undefined")
	    return strID;
	window.frmUpdate.elements(FieldID2).value = strID;
}
//-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY bgcolor=ivory>
<%
	dim cn
	dim rs 
	dim strSQL
    	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    Response.Write "<h3>Update Current Commitment Dates</h3>"
    Response.Write "<h5>These are required when the product Phase is changed to Development, Production, or Post-Production.</h5>"
	%>
    
	<form id="frmUpdate" method="post" action="TargetRTPandEMUpdateSave.asp">
	<table id="tabUpdate" width="100%" bgcolor="cornsilk" border="1" cellspacing="0" cellpadding="2" bordercolor="tan">
    <%

	strSQL = "usp_SelectScheduleTargetRTPandEOMDates " & clng(request("ID"))
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
    %>
	<tr>
	    <td style="white-space:nowrap; vertical-align:top; font-weight:bold;"><%=rs("item_description")%>:</td>
		<td>
            <input type="hidden" id="scheduledataid<%=rs("schedule_definition_data_id")%>" name="scheduledataid<%=rs("schedule_definition_data_id")%>" value="<%=rs("schedule_data_id")%>" />
		    <input type="hidden" id="hdn<%=rs("schedule_definition_data_id")%>" name="hdn<%=rs("schedule_definition_data_id")%>" value="<%=rs("projected_start_dt")%>" />
		    <input type="text" id="txt<%=rs("schedule_definition_data_id")%>" name="txt<%=rs("schedule_definition_data_id")%>" value="<%=rs("projected_start_dt")%>" onkeydown="return false;" autocomplete="off" />
		    <a id="hrefProjectStartDtPicker" href="javascript: cmdDate_onclick('hdn<%=rs("schedule_definition_data_id")%>', 'txt<%=rs("schedule_definition_data_id")%>')"><img id="imgProjectStartDtPicker" src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21" /></a>
	    </td>
	</tr>
    <%
    rs.MoveNext
	loop
	rs.Close    
    %>
	</table>    
    <input type="hidden" id="PVID" name="PVID" value="<%=Request("ID")%>" />
    </form>
<%
	
set rs = nothing
set cn = nothing

%>
</BODY>
</HTML>