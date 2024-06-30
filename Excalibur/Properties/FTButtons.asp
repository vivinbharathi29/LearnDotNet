<%@ Language=VBScript %>

<%
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

%>


<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdOK_onclick() {

	window.parent.frames["MainWindow"].PrintScreen();
	
	
}



function cmdCancel_onclick() {
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.close();
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<TABLE  BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="Print" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Close" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</TABLE>

</BODY>
</HTML>