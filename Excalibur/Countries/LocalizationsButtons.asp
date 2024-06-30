<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
    <title></title>
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(function () {
            $("input:button").button();
        });
    </script>

<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
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

function VerifySave(){
	var blnSuccess = true;	
	if (window.parent.frames["UpperWindow"].frmCountries.cboDcr)
	{	
		if (window.parent.frames["UpperWindow"].frmCountries.cboDcr.value == 0)
		{
			alert('Please Select the appropriate DCR');
			return false
		}
	}
	return blnSuccess;
}

function cmdCancel_onclick() {
    var pulsarplusDivId = document.getElementById('hdnTabName');
    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    }
    else {

        var iframeName = parent.window.name;
        if (iframeName != '') {
            parent.window.parent.ClosePropertiesDialog();
        } else {
            window.parent.close();
        }
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].frmCountries.submit();
		}

}


//-->
</script>
</head>
<body bgcolor="ivory">
<div style="text-align:right;">
		<input type="button" value="OK" id="cmdOK" name="cmdOK" onclick="return cmdOK_onclick()" />
		<input type="button" value="Cancel" id="cmdCancel" name="cmdCancel"  onclick="return cmdCancel_onclick()" />
      <input type="hidden" id="hdnTabName" name="hdnTabName" value="<%= Request("pulsarplusDivId")%>" />
</div>
</body>
</html>