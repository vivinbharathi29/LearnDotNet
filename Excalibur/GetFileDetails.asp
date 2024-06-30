<%@ Language=VBScript %>

<%Response.Expires = 0%>

<HTML>
<HEAD><TITLE>Add Device</TITLE>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--




function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
FindFile.submit();	
}




function window_onload() {
	//FindFile.file1.click();
}



function file1_onchange() {
	//FindFile.submit();	
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=gainsboro LANGUAGE=javascript onload="return window_onload()">
<form ENCTYPE="multipart/form-data" ID=FindFile ACTION="Upload/Upload.asp" METHOD="post" NAME="FindFile">
<TABLE WIDTH="100%" bgcolor=gainsboro BORDER=0 CELLSPACING=0 CELLPADDING=2 bordercolor=tan style="LEFT: 9px; BORDER-BOTTOM: lightgrey thin solid; TOP: 1px" align=center>
	<TR>
		<TD nowrap width="10%" vAlign=top><FONT size=1 face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT>Select&nbsp;File:</FONT></TD>
		<TD  vAlign=top>&nbsp;<INPUT style="WIDTH:340" id=file1 type=file name=file1 LANGUAGE=javascript onchange="return file1_onchange()"></TD>
    </TR>
	</TABLE>
	<table width="100%" style="BORDER-TOP: white thin solid">
	<TR>
		<TD colspan=2 align=right bgColor=gainsboro>
      <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()" style="FONT-FAMILY: "> <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>
</form>
</BODY>
</HTML>
