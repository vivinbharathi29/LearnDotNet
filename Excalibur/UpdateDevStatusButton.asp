<%@ Language=VBScript %>
<html>
<head>
    <link href="style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        $(function () {
            $("input:button").button();
        });
    </script>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	if (window.parent.frames["UpperWindow"].frmUpdateDevStatus.cboDevStatus.value == "2")
		if (window.parent.frames["UpperWindow"].frmUpdateDevStatus.txtComments.value == "")
		{
			alert("You must explain in the comments field why you set the status to rejected");
			window.parent.frames["UpperWindow"].frmUpdateDevStatus.txtComments.focus();
			blnSuccess = false;
		}
	
	return blnSuccess;
}

function ValidDate(pDate){
	if (isNaN(Date.parse(pDate)))
		return false;
	else
		return true;

}

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function cmdCancel_onclick() {
    window.parent.Cancel();
}

function cmdOK_onclick() {
	var blnAll = true;
	var i;
	
	if (VerifySave()) {		
		cmdCancel.disabled =true;
		cmdOK.disabled =true;
		window.parent.frames["UpperWindow"].frmUpdateDevStatus.txtStatusName.value = window.parent.frames["UpperWindow"].frmUpdateDevStatus.cboDevStatus.options[window.parent.frames["UpperWindow"].frmUpdateDevStatus.cboDevStatus.selectedIndex].text;
		window.parent.frames["UpperWindow"].frmUpdateDevStatus.submit();
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