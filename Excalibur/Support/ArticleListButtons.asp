<%@ Language=VBScript %>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function cmdOK_onclick(){
    window.parent.frames["MainWindow"].frmMain.submit();
}

function cmdCancel_onclick() {
    //window.parent.close();
    parent.window.parent.modalDialog.cancel(false);
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<TABLE  BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</TABLE>

</BODY>
</HTML>