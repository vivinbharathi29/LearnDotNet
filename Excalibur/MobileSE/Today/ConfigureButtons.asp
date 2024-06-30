<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--



function cmdCancel_onclick() {
		CloseIframeDialog();
}

function cmdOK_onclick() {
	window.parent.frames["UpperWindow"].Configure.submit();
	parent.window.parent.location.reload(true);
}

function CloseIframeDialog() {
    var iframeName = window.name;
    if (iframeName != '') {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            parent.window.parent.ClosePropertiesDialog();
        }
        
    } else {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {

            window.close();
        }
        //window.close();
    }
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>