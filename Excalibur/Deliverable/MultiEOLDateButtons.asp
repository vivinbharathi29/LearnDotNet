<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../includes/Date.asp" -->


function VerifySave(){
	var blnSuccess = true;	
	var blnFound = false;
	
	
	if (typeof(window.parent.frames["UpperWindow"].frmMain.lstID.length)!="undefined")
		{
			for (i=0;i<window.parent.frames["UpperWindow"].frmMain.lstID.length;i++)
				if (window.parent.frames["UpperWindow"].frmMain.lstID[i].checked)
					blnFound=true;
		}
	else
		{
			if (window.parent.frames["UpperWindow"].frmMain.lstID.checked)
				blnFound=true;
		}	
	
	if (! blnFound )
		{
			blnSuccess=false;
			alert("You must select at least one deliverable to continue.");
		}
	else if(window.parent.frames["UpperWindow"].frmMain.txtEOLDate.value != "" && window.parent.frames["UpperWindow"].frmMain.cboDateChange.selectedIndex==1)
		{
		if(! isDate(window.parent.frames["UpperWindow"].frmMain.txtEOLDate.value))
			{
			blnSuccess=false;
			alert("Date is not a valid date.");
			window.parent.frames["UpperWindow"].frmMain.txtEOLDate.focus();
			}
		}
	else if(window.parent.frames["UpperWindow"].frmMain.cboDateChange.selectedIndex==0)
		{
			blnSuccess=false;
			alert("No change selected.");
			window.parent.frames["UpperWindow"].frmMain.cboDateChange.focus();
		}
	return blnSuccess;
}

function cmdCancel_onclick() {
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else{
        window.parent.close();
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmMain.submit();
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