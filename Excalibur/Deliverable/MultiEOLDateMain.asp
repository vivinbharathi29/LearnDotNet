<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdEOLDate_onclick() {
	var strID;
	strID = window.showModalDialog("../mobilese/today/calDraw1.asp",frmMain.txtEOLDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			frmMain.txtEOLDate.value = strID;
		}
}

function cboDateChange_onclick() {
	if (frmMain.cboDateChange.selectedIndex==1)
		{
		DateField.style.display="";
		frmMain.txtEOLDate.focus();
		}
	else
		DateField.style.display="none";


	if (frmMain.cboDateChange.selectedIndex==2)
		DeactivateWarning.style.display="";
	else
		DeactivateWarning.style.display="none";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
.VersionTable TD
{
	FONT-SIZE: xx-small;
    COLOR: black;
    FONT-FAMILY: Verdana
}
</STYLE>
<BODY bgcolor=ivory>
<LINK href="../style/wizard style.css" type=text/css rel=stylesheet >
<%if trim(request("TypeID"))="2" then%>
	<font size=3 face=verdana><b>Update End of Service Availability Date</b><BR><BR></font>
<%else%>
	<font size=3 face=verdana><b>Update End of Availability Date</b><BR><BR></font>
<%end if%>
<%
	dim cn 
	dim rs
	dim strSQL
	dim strName
	dim strVersion
	dim strRevision
	dim strPass
	dim strTypeID
	dim strModelNumber
	dim strPartNumber
	dim strEOLDate
	dim strVendor
	

'	if strName = "" then
'		Response.write "<font size=2 face=verdana>Deliverable not found. (" & request("IDList") & ")</font>"	
'	else

%>
<form ID=frmMain action="MultiEOLDateSave.asp" method=post>
<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>End&nbsp;of&nbsp;Availability:&nbsp;&nbsp;&nbsp;</b></TD>
	<TD width="100%" nowrap>
		<SELECT id=cboDateChange name=cboDateChange style="width: 190" LANGUAGE=javascript onchange="return cboDateChange_onclick()">
			<OPTION selected value="">No Change</OPTION>
			<OPTION value="1">Change Date</OPTION>
			<OPTION value="2">Deactivate</OPTION>
            <%if trim(request("TypeID"))="2" then%>
    			<OPTION value="3">Sync to Factory EOA Date</OPTION>
			<%end if %>
		</SELECT>
			<span ID=DateField style="display:none">
				<INPUT type="text" id=txtEOLDate name=txtEOLDate value="<%=strEOLDate%>">
				<a href="javascript: cmdEOLDate_onclick()"><img ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
			</span>
			<span style="display:none" ID=DeactivateWarning><font size=1 face =verdana color=green>The selected deliverables will set to "Unavailable".</font></span>
	</TD>
</TR>
</table>


<%
  function ScrubSQL(strWords) 

    dim badChars 
    dim newChars 
    dim i

    
    strWords=replace(strWords,"'","''")
    
    badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
    newChars = strWords 
    
    for i = 0 to uBound(badChars) 
      newChars = replace(newChars, badChars(i), "") 
    next 
    
    ScrubSQL = newChars 
  
  end function 
  	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	strSQL = "SELECT v.id, v.deliverablename, v.partnumber, v.modelnumber, v.version, v.revision, v.pass, vd.name as Vendor, endoflifedate as eoadate, serviceeoadate " & _
			 "FROM DeliverableVersion v with (NOLOCK), vendor vd with (NOLOCK) " & _
			 "WHERE v.vendorid = vd.id " & _
			 "and v.id in (" & scrubsql(request("IDList")) & ") " & _
			 "ORDER BY v.deliverablename, v.id desc;"
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open strSQL,cn,adOpenForwardOnly

	if rs.EOF and rs.BOF then
		Response.Write "<BR><BR><font size=2 face=verdana>No Deliverables Selected.</font>"
	else
		Response.Write "<BR><BR><font size=2 face=verdana><b>Deliverables to Update.</b><BR></font>"
		Response.Write "<table class=VersionTable  ID=VersionTable border=1 width=""100%"" bordercolor=tan cellspacing=0 cellpadding=2>"
		if trim(request("TypeID"))="2" then	
			Response.Write "<TR bgcolor=cornsilk><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>Available&nbsp;Until</b></TD><TD><b>Factory&nbsp;EOA</b></TD><TD><b>Version</b></TD><TD><b>Vendor</b></TD><TD><b>Model</b></TD><TD><b>Part&nbsp;Number</b></TD></TR>"
		else
			Response.Write "<TR bgcolor=cornsilk><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>Available&nbsp;Until</b></TD><TD><b>Version</b></TD><TD><b>Vendor</b></TD><TD><b>Model</b></TD><TD><b>Part&nbsp;Number</b></TD></TR>"
		end if
		do while not rs.EOF
		
			strVersion = rs("Version") & ""
			if trim(rs("Revision") & "") <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if trim(rs("Pass") & "") <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if

			if trim(request("TypeID"))="2" then
				strEOL = rs("ServiceEOADate") & ""
			else
				strEOL = rs("eoadate") & ""
			end if
			
			Response.Write "<TR>"
			Response.Write "<TD><INPUT type=""checkbox"" checked id=lstID name=lstID style=""WIDTH: 14px; HEIGHT: 14px"" size=""14"" value=""" & rs("ID") & """>&nbsp;" & rs("ID") & "</TD>"
			Response.Write "<TD>" & rs("DeliverableName") & "</TD>"
			Response.Write "<TD>" & strEOL & "&nbsp;</TD>"
			if trim(request("TypeID"))="2" then	
				Response.Write "<TD>" & rs("eoadate") & "&nbsp;</TD>"
			end if
			Response.Write "<TD>" & strVersion & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Vendor") & "</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "</TD>"
			Response.Write "<TD>" & rs("PartNumber") & "</TD>"
			rs.MoveNext
		loop
		Response.Write "</TR>"
		Response.Write "</table>"	
	end if

	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing
%>

<INPUT type="hidden" id=txtTypeID name=txtTypeID value="<%=request("TypeID")%>">

</form>

	<%'end if%>
</BODY>
</HTML>
