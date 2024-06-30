<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
		window.parent.close();
}

function cmdOK_onclick() {
    var rc;
    if(window.parent.frames("UpperWindow").optType1.checked)
        {
        rc=1;
        }
    else if(window.parent.frames("UpperWindow").optType2.checked)
        {
        rc=2;
        }
    else if(window.parent.frames("UpperWindow").optType3.checked)
        {
        rc=3;
        alert("We are currently working on automating this process.  In the meantime, you can following the manual process outlined below:\r\r1. Contact John Roche to receive approval for updating the CVA file.\r2. Make any necessary updates to the deliverable property screen in Excalibur.\r3. Create the new CVA file.\r4. Attach the new CVA file to John's approval email and forward it to HouHPQBNBContacts@hp.com\r\rThe Release Team will archive the old CVA File, copy the new one into the deliverable folder, and send an email notification of the change to the normal preinstall deliverable release distribution list.");
        }
    else if(window.parent.frames("UpperWindow").optType4.checked)
        {
        rc=4;
        }
    
	window.returnValue = rc;
	window.parent.opener='X';
	window.parent.open('','_parent','')
	window.parent.close();	
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