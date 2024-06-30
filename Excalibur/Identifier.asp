<%@ Language=VBScript %>
<HTML>
<HEAD><TITLE>Add System Identifier</TITLE>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var s = window.dialogArguments;
	txtID.value = s.substr(0,4); 
	txtName.value = s.substr(7); 
}

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
	window.returnValue = txtID.value + ' - ' + txtName.value;
	window.close();
}

function txtID_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
		
}

function txtName_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=cornsilk LANGUAGE=javascript onload="return window_onload()">
<TABLE WIDTH="100%" bgcolor=cornsilk BORDER=0 CELLSPACING=0 CELLPADDING=2 bordercolor=tan style="LEFT: 9px; BORDER-BOTTOM: tan thin solid; TOP: 1px" align=center>
	<TR>
		<TD nowrap width="10%"  
    vAlign=top><FONT size=1 
      face=Verdana>System 
      ID:</FONT> <FONT color=#ff0000 size=1>*</FONT> 
       </TD>
		<TD  vAlign=top><INPUT id=txtID name=txtID style="WIDTH: 83px; HEIGHT: 22px" size=11 LANGUAGE=javascript onkeypress="return txtID_onkeypress()">&nbsp;<FONT color=#0000ff size=1 
      face="MS Sans Serif">(4-digit hex number)</FONT></TD>
    </TR>
	<TR>
		<TD width="10%" nowrap  vAlign=top><FONT size=1 
     face=Verdana>System 
      Name(s):</FONT> <FONT color=#ff0000 size=1>*</FONT></TD>
		<TD  vAlign=top><INPUT id=txtName 
      style="WIDTH: 100%; HEIGHT: 22px" size=72 name=txtName LANGUAGE=javascript onkeypress="return txtName_onkeypress()"></TD>
	</TR>
	</TABLE>
	<table width="100%">
	<TR>
		<TD colspan=2 align=right bgColor=cornsilk>
      <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()" style="FONT-FAMILY: "> <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>

</BODY>
</HTML>
