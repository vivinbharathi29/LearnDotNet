<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
    <title></title>
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>

<script type="text/javascript" language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS language=JavaScript >
<!--

function frmMilestoneVerify()
{
	with (window.parent.frames["UpperWindow"].frmUpdate)
	{
	    if (!validateDateInput(txt60, 'Target RTP/MR Date')) { return false; }
	    if (!validateDateInput(txt112, 'End of Manufacturing (EM) Date')) { return false; }
	    if ((txt60.value == "") || (txt112.value == "")) { alert("Please populate both dates."); return false; }
	}
	return true;
}

function cmdCancel_onclick() 
{
    var iframeName = parent.window.name;
    if (iframeName != '') {
        parent.window.parent.ClosePropertiesDialog();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() 
{
	cmdCancel.disabled = true;
	cmdOK.disabled = true;
	if (window.parent.frames["UpperWindow"].frmUpdate)
	{
		if (frmMilestoneVerify())
		{
			window.returnValue=1;
			window.parent.frames["UpperWindow"].frmUpdate.submit();
		}
		else
		{
			cmdCancel.disabled = false;
			cmdOK.disabled = false;
		}
	}	
	else
	{
	    window.parent.frames["UpperWindow"].frmUpdate.submit();
	}
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
    </tr>
</table>
</body>
</html>