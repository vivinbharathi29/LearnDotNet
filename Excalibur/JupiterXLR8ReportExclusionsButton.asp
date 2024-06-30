<%@ Language=VBScript %>
<html>
<head>
    <link href="style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="includes/client/jquery-ui.min.js" type="text/javascript"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
    window.parent.close();
}

function cmdOK_onclick() {
	cmdCancel.disabled =true;
	cmdOK.disabled = true;
	window.parent.frames["UpperWindow"].GetSelectedProducts();
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right" >
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>