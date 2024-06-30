<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
 <script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdOK_onclick() {
	if (frmMain.cboTTS.selectedIndex == 0 )
		{
		alert("You must select the WWAN TTS to continue");
		frmMain.cboTTS.focus();
		}
	else
		frmMain.submit();
}

function cmdCancel_onclick() {
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.close();
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<LINK href="../style/wizard style.css" type=text/css rel=stylesheet >
<font size=3 face=verdana><b>Update WWAN TTS Status</b><BR><BR></font>
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
	dim strVendor
	dim strWWANTTS
	dim strCategory
	
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
		strWWANTTS = rs("TTS") & ""
		strCategory =  trim(rs("Category") & "")
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
<form ID=frmMain action="WWANPanelStatusSave.asp" method=post>
<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Name:</b></TD>
	<TD colspan=3 width="100%"><%=strVendor & " " & strName%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Category:</b></TD>
	<TD colspan=3 width="100%"><%=strCategory%></TD>
</TR>

<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>HW Version:</b></TD>
	<TD width="50%"><%=strVersion%>&nbsp;</TD>
	<TD nowrap valign=top><b>FW Version:</b></TD>
	<TD width="50%"><%=strRevision%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Revision:</b></TD>
	<TD width="50%"><%=strPass%>&nbsp;</TD>
	<TD nowrap valign=top><b>Model&nbsp;Number:</b></TD>
	<TD width="50%"><%=strModelNumber%>&nbsp;&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Part&nbsp;Number:&nbsp;&nbsp;</b></TD>
	<TD width="50%"><%=strPartNumber%>&nbsp;</TD>
	<TD nowrap valign=top><b>WWAN TTS:</b>&nbsp;<font color="#ff0000" size="1">*</font></TD>
	<TD width="50%"><SELECT style="width:100%" id=cboTTS name=cboTTS>
						<OPTION selected></OPTION>						<OPTION>Waived</OPTION>						<OPTION>Failed</OPTION>
					</SELECT>
	</TD>
</TR>
</table>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">

<TABLE width=100%>
	<TR><TD align=right>
		<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
		<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
		</TD>
	</TR>
</Table>

</form>
	<%end if%>
</BODY>
</HTML>
