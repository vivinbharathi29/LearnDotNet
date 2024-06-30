<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script type="text/javascript" src="../Scripts/PulsarPlus.js"></script>
<script src="../includes/client/jquery-1.11.0.min.js"></script>
<script src="../includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifySave(){
	var blnSuccess = true;	
	if (window.parent.frames["UpperWindow"].frmMain.txtComments.value == "")
		{
			alert("Comments are required when rejecting delverable requests.");
			window.parent.frames["UpperWindow"].frmMain.txtComments.focus();
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
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.cancel();
    }
}

function cmdOK_onclick() {
	var blnAll = true;
	var i;
	
	if (VerifySave()) {	    
	    var ajaxurl = 'SaveDevNotificationStatus.asp?Type=1&txtMultiID=' + window.parent.frames["UpperWindow"].frmMain.txtID.value + '&NewValue=' + window.parent.frames["UpperWindow"].frmMain.txtNewValue.value + '&txtComments=' + window.parent.frames["UpperWindow"].frmMain.txtComments.value;	    
	    $.ajax({
	        url: ajaxurl,
	        type: "GET",
	        async: false,
	        success: function (data) {
	            window.parent.closewindow(2, window.parent.frames["UpperWindow"].frmMain.txtID.value);
	        },
	        error: function (xhr, status, error) {
	            alert(error);
	        }
	    });
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