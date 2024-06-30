<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdOK_onclick() {


	window.parent.frames["UpperWindow"].cmdTo_onclick();
	window.returnValue = window.parent.frames["UpperWindow"].hidReturn.value;
	window.parent.close();
	
}



function cmdCancel_onclick() {
	//if (window.confirm ("Are you sure you want to exit this screen without releasing this deliverable?") == true)
		window.parent.close();
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