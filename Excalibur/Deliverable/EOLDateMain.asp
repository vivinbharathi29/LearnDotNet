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

function window_onload() {
	frmMain.txtEOLDate.focus();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<LINK href="../style/wizard style.css" type=text/css rel=stylesheet >
<%if trim(request("TypeID"))="2" then%>
	<font size=3 face=verdana><b>Update End of Service Availability Date</b><BR><BR></font>
<%else%>
	<font size=3 face=verdana><b>Update End of Availability Date</b><BR><BR></font>
<%end if%>
<%
	dim cn 
	dim rs
	dim strName
	dim strVersion
	dim strRevision
	dim strPass
	dim strTypeID
	dim strModelNumber
	dim strPartNumber
	dim strEOLDate
    dim strSEOLDate
	dim strVendor
	dim strFactoryEOA
	dim strEOAText
	dim strChkEOL

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spGetVersionProperties4Web " & clng(request("ID")),cn,adOpenForwardOnly
	if not(rs.EOF and rs.BOF) then
		strName = trim(rs("DeliverableName") & "")
		strVersion = trim(rs("Version") & "")
		strRevision = trim(rs("Revision") & "")
		strPass = trim(rs("Pass") & "")
		strTypeID = trim(rs("TypeID") & "")
		strModelNumber = trim(rs("ModelNumber") & "")
		strPartNumber = trim(rs("PartNumber") & "")
		
        strChkEOL = ""

        strSEOLDate = rs("ServiceEOADate") & ""
        strEOLDate = rs("EOLDate") & ""
		strFactoryEOA = rs("EOLDate") & "&nbsp;"

        if trim(request("TypeID"))="2" then
            if rs("ServiceActive") = true then
                strChkEOL = "checked"
            end if
        else
            if rs("ActiveVersion") = true then
                strChkEOL = "checked"
            end if
        end if

		if trim(rs("VersionVendor") & "") <> "" then
			strVendor = rs("VersionVendor") & ""
		else
			strVendor = rs("Vendor") & ""
		end if 


        
	end if
	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing

	if strName = "" then
		Response.write "<font size=2 face=verdana>Deliverable not found. (" & request("ID") & ")</font>"	
	else

%>
<form ID=frmMain action="EOLDateSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Name:</b></TD>
	<TD width="100%"><%=strName%></TD>
</TR>
<% if strTypeID = "1" then%>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>HW Version:</b></TD>
	<TD width="100%"><%=strVersion%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>FW Version:</b></TD>
	<TD width="100%"><%=strRevision%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Revision:</b></TD>
	<TD width="100%"><%=strPass%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Vendor:</b></TD>
	<TD width="100%"><%=strVendor%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Model&nbsp;Number:</b></TD>
	<TD width="100%"><%=strModelNumber%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Part&nbsp;Number:</b></TD>
	<TD width="100%"><%=strPartNumber%>&nbsp;</TD>
</TR>
<%else%>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Version:</b></TD>
	<TD width="100%"><%=strVersion%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Revision:</b></TD>
	<TD width="100%"><%=strRevision%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Pass:</b></TD>
	<TD width="100%"><%=strPass%></TD>
</TR>
<%end if%>

<%if trim(request("TypeID"))="2" then %>
    <TR bgcolor=cornsilk>
	    <TD nowrap valign=top><b>Factory EOA:</b></TD>
	    <TD width="100%"><%=strFactoryEOA%></TD>
    </TR>
    <TR bgcolor=cornsilk>
	    <TD nowrap valign=top><b>Service&nbsp;EOA:</b></TD>
	    <TD width="100%"><INPUT type="text" id=txtEOLDate name=txtEOLDate value="<%=strSEOLDate%>">
	    <a href="javascript: cmdEOLDate_onclick()"><img ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
	    </TD>
    </TR>
    <TR bgcolor=cornsilk>
	    <TD nowrap valign=top><b>Service Active Status:</b></TD>
	    <TD width="100%"><INPUT type="checkbox" <%=strChkEOL %>  id=chkEOL name=chkEOL></TD>
    </TR>
<%else %>

    <TR bgcolor=cornsilk>
	    <TD nowrap valign=top><b>End&nbsp;of&nbsp;Availability: </b></TD>
	    <TD width="100%"><INPUT type="text" id=txtEOLDate name=txtEOLDate value="<%=strEOLDate%>">
	    <a href="javascript: cmdEOLDate_onclick()"><img ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
	    </TD>
    </TR>
    <TR bgcolor=cornsilk>
	    <TD nowrap valign=top><b>Availability:</b></TD>
	    <TD width="100%"><INPUT type="checkbox" <%=strChkEOL %>  id=chkEOL name=chkEOL>This Deliverable is Active.</TD>
    </TR>
<%end if %>


</table>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtTypeID name=txtTypeID value="<%=request("TypeID")%>">

</form>

	<%end if%>
</BODY>
</HTML>
