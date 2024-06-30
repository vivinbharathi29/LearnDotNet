<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--



function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() {
    
    var strCycle="";
    
    for (i=0;i< window.parent.frames["UpperWindow"].frmMain.chkCycle.length;i++)
        if (window.parent.frames["UpperWindow"].frmMain.chkCycle[i].checked)
            strCycle = strCycle + ", " + window.parent.frames["UpperWindow"].frmMain.chkCycle[i].getAttribute('CycleName');
    
    if (strCycle!="")
        strCycle = strCycle.substring(2);
    
    window.parent.frames["UpperWindow"].frmMain.txtCycleList.value= strCycle;
	window.parent.frames["UpperWindow"].frmMain.submit();
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