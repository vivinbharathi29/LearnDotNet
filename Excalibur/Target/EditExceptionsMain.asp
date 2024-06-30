<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {
	frmChange.submit();
}

function cmdCancel_onclick() {
    if (parent.window.parent.loadDatatodiv != undefined) {
        parent.window.parent.closeExternalPopup();
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}


function window_onload() {
	frmChange.txtExceptions.focus();
}


function ChangeThis_onclick() {
	frmChange.optThis.checked=true;
	frmChange.optFuture.checked=false;
}

function ChangeDefault_onclick() {
	frmChange.optThis.checked=false
	frmChange.optFuture.checked=true
}

function ChangeThis_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ChangeDefault_onmouseover() {
	window.event.srcElement.style.cursor = "hand";

}

//-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY  bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%

	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strExceptions
	dim strOOC

	strExceptions = ""
	strOOC = ""
	
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")


	if not blnLoadFailed then
		rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProdName = ""
			blnLoadFailed = true
		else
			strprodName = rs("name") & ""
		end if
		
		rs.Close
	end if

	if not blnLoadFailed then
		rs.Open "spGetDeliverableVersionProperties " & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strDeliverable = ""
			blnLoadFailed = true
		else
			strDeliverable = rs("name") & "&nbsp;-&nbsp;" & rs("version")
			if rs("Revision") & "" <> "" then
				strDeliverable = strDeliverable & "," & rs("Revision") & ""
			end if
			if rs("Pass") & "" <> "" then
				strDeliverable = strDeliverable & "," & rs("Pass") & ""
			end if
			strDeliverable = strDeliverable & "&nbsp;"
			
		end if
		
		rs.Close
	end if

	if not blnLoadFailed then
		rs.Open "spGetTargetNotes " & clng(request("ProductID")) & "," & clng(request("VersionID")) ,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			blnLoadFailed = true
		else
			strExceptions = rs("TargetNotes") & ""
			strOOC = rs("OOCRelease")
			if isnull(rs("OOCRelease")) then
				strOOC = ""
			elseif rs("OOCRelease") then
				strOOC = " checked "
			end if
			
		end if
		
		rs.Close
	end if


%>



<h3>Edit Target Notes<h3>
<h4><%=strDeliverable & " (" & strProdName & ")"%></h4>

<form ID=frmChange action="EditExceptionsSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">

<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Target&nbsp;Notes:</b></font></TD>
		<TD valign=top><INPUT type="text" style="Width:100%" id=txtExceptions name=txtExceptions value="<%=strExceptions%>" maxlength=255>
		</TD>
	</TR>
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Release:</b></font></TD>
		<TD valign=top><INPUT <%=strOOC%> type="checkbox" id=chkOOC name=chkOOC>&nbsp;Out&nbsp;of&nbsp;Cycle&nbsp;Release
		</TD>
	</TR>
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Scope:</b></font></TD>
		<TD valign=top>
		<INPUT type="radio" id=optThis name=optScope value="1">&nbsp;<font size=2 face=verdana ID=ChangeThis LANGUAGE=javascript onclick="return ChangeThis_onclick()" onmouseover="return ChangeThis_onmouseover()">Change this version only</font><BR>
		<INPUT type="radio" id=optFuture name=optScope value="2" checked>&nbsp;<font size=2 face=verdana ID=ChangeDefault LANGUAGE=javascript onclick="return ChangeDefault_onclick()" onmouseover="return ChangeDefault_onmouseover()">Change this version and all future versions</font><BR>
		<!--<INPUT type="radio" id=optAll name=optScope value="3" >&nbsp;<font size=2 face=verdana>Change All Existing and Future Versions</font><BR>-->
		</TD>
	</TR>
</table>
</form>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript  onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</TABLE>


<%
	set rs= nothing
	set cn=nothing
%>



</BODY>
</HTML>
