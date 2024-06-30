<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>WHQL Signature Request</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdTargetDate_onclick() {
	var strID;
	strID = window.showModalDialog("../mobilese/today/calDraw1.asp",frmUpdate.txtDateNeeded.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			frmUpdate.txtDateNeeded.value = strID;
		}
}


//-->

</SCRIPT>
<%
	dim strDriverCategory
	dim strVendorName
	dim strOperatingSystem
	dim strVersionPass
	dim strLinktoDriver
	dim strFileName
	dim strPlatform
	dim strDateNeeded
%>

</HEAD>
<link href="../../wizard%20style.css" type="text/css" rel="stylesheet">
<BODY bgcolor=Ivory>

<font face=verdana size=><b>
<label ID="lblTitle">
	<br>WHQL Test Signature Request Email Form
</label></b></font>
		
<form id="frmUpdate" method="post" action="WHQLSignatureSave.asp">

<table WIDTH=900 BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr><td valign=top width=200 nowrap><b>&nbsp;Vendor Name:</b></td>
	<td><INPUT type="text" style="width:400" id=txtVendorName name=txtVendorName value="<%=strVendorName%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: ATI, Synaptics, Broadcom... </td></tr>	

	<tr><td valign=top width=200 nowrap><b>&nbsp;Driver Category:</b></td>
	<td><INPUT type="text" style="width:400" id=txtDriverCategory name=txtDriverCategory value="<%=strDriverCategory%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: Display, WLAN, LAN, Input/Ouput... </td></tr>	

	<tr><td valign=top width=200 nowrap><b>&nbsp;Operating System:</b></td>
	<td><INPUT type="text" style="width:400" id=txtOperatingSystem name=txtOperatingSystem value="<%=strOperatingSystem%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: Windows XP, Windows 2000</td></tr>	

	<tr><td valign=top width=200 nowrap><b>&nbsp;Version/Pass:</b></td>
	<td><INPUT type="text" style="width:400" id=txtVersionPass name=txtVersionPass value="<%=strVersionPass%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: 1.00 A5</td></tr>	
	<tr><td valign=top width=200 nowrap><b>&nbsp;Link to Driver:</b></td>
	<td><INPUT type="text" style="width:400" id=txtLinktoDriver name=txtLinktoDriver value="<%=strLinktoDriver%>" maxlength=255>
	</td><td valign=top width=300 nowrap>example: \\TLHOME\SHare</td></tr>	
	<tr><td valign=top width=200 nowrap><b>&nbsp;File Name:</b></td>
	<td><INPUT type="text" style="width:400" id=txtFileName name=txtFileName value="<%=strFileName%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: NX51V_A5.100.ZIP</td></tr>	
	<tr><td valign=top width=200 nowrap><b>&nbsp;Platform(s) Supported:</b></td>
	<td><INPUT type="text" style="width:400" id=txtPlatform name=txtPlatform value="<%=strPlatform%>" maxlength=80>
	</td><td valign=top width=300 nowrap>example: Davos 1.0</td></tr>	
	<tr><td valign=top width=200 nowrap><b>&nbsp;Date Needed:</b></td>
	<td><INPUT type="text" style="width:370" id=txtDateNeeded name=txtDateNeeded value="<%=strDateNeeded%>" maxlength=50>
	<a href="javascript: cmdTargetDate_onclick()"><img ID="picTarget" SRC="../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
	</td><td valign=top width=300 nowrap>example: 04/28/2004</td></tr>			<tr><td valign=top width=200 nowrap><input type="checkbox" checked id="chkCCEmail" name="chkCCEmail">Send Confirmation</td></tr>
</table>

</form>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
frmUpdate.txtVendorName.focus()
//-->
</script>

</BODY>
</HTML>
