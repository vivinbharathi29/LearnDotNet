<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--




function ltrim ( s ) 
{ 
	return s.replace( /^\s*/, "" ) 
} 

function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

/*function VerifyNext(){
	var blnSuccess = true;	
	var i;
	var strProducts = "";
	
	for (i=0;i<window.parent.frames["UpperWindow"].frmSoftpaq.chkSelected.length;i++)
		{
			if (window.parent.frames["UpperWindow"].frmSoftpaq.chkSelected(i).checked)
				if (strProducts == "")
					strProducts = strProducts + window.parent.frames["UpperWindow"].frmSoftpaq.chkSelected(i).value;
				else
					strProducts = strProducts + "\r" + window.parent.frames["UpperWindow"].frmSoftpaq.chkSelected(i).value;
		}
		
	window.parent.frames["UpperWindow"].txtSelectedProducts.value = strProducts;
	
	if (strProducts == "")
		{
		window.alert("You must select at least one product to continue.");
		blnSuccess = false;
		}
	return blnSuccess;
}*/

function cmdCancel_onclick() {
		window.parent.close();
}

function cmdFinish_onclick(){
	window.parent.frames["UpperWindow"].frmSoftpaq.submit();
}

function cmdPrevious_onclick() {
	switch (window.parent.frames["UpperWindow"].CurrentState)
	{
		case "Preview":
			window.parent.frames["UpperWindow"].CurrentState = "CVA";
		break;
		case "CVA":
			window.parent.frames["UpperWindow"].CurrentState = "Softpaq";
		break;
	}
	window.parent.frames["UpperWindow"].ProcessState();
}

function cmdNext_onclick() {
	var blnOK = true;
	
	switch (window.parent.frames["UpperWindow"].CurrentState)
	{
		case "CVA":
			window.parent.frames["UpperWindow"].CurrentState = "Preview";
			//blnOK = VerifyNext();
		
			//if (blnOK)
		break;
		case "Softpaq":
			window.parent.frames["UpperWindow"].CurrentState = "CVA";
		break;
	}
	window.parent.frames["UpperWindow"].ProcessState();
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>

			<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>
</TR></table>
</body>
</html>