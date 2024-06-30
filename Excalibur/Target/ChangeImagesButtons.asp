<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<script src="../Scripts/Pulsar2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


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

function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

function VerifySave(){
	var blnSuccess = true;	
	if (window.parent.frames["UpperWindow"].frmChange.txtSummary.value == "")
		{
			alert("You must specify a summary of your image strategy");
			window.parent.frames["UpperWindow"].frmChange.txtSummary.focus();
			blnSuccess = false;
		}
	
	return blnSuccess;
}

function cmdCancel_onclick() {
    if (isFromPulsar2()) {
        closePulsar2Popup(false);
    }
    else if (parent.window.parent.loadDatatodiv != undefined) {
        parent.window.parent.closeExternalPopup();
    }
    else if (parent.parent.window.parent.loadDatatodiv != undefined) {
        parent.window.parent.closeExternalPopup();
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}


function TransferRestoreimageInfo(){
    var strOutput="";
    window.parent.frames["UpperWindow"].GetRestoreLanguageChanges();
    if (typeof(window.parent.frames["UpperWindow"].chkImage) != "undefined"){
    	if(typeof(window.parent.frames["UpperWindow"].chkImage.length)=="undefined")
            {
            if (window.parent.frames["UpperWindow"].chkImage.checked)
                window.parent.frames["UpperWindow"].frmChange.chkRestoreImage.value= window.parent.frames["UpperWindow"].chkImage.value + ",";
            }
        else
            {
            for (i=0;i<window.parent.frames["UpperWindow"].chkImage.length;i++)
			    {
                if (window.parent.frames["UpperWindow"].chkImage[i].checked)
                    strOutput= strOutput + ", " + window.parent.frames["UpperWindow"].chkImage[i].value;

                }
                if (strOutput != "")
                    strOutput = strOutput.substr(2) + ",";

                window.parent.frames["UpperWindow"].frmChange.chkRestoreImage.value= strOutput;
            }
    }

}


function cmdOK_onclick(pulsarplusDivId) {
	var blnAll = true;
	var blnNone = true;
	if (VerifySave())
		{
            TransferRestoreimageInfo();
		    window.parent.frames["UpperWindow"].GetLanguageChanges();
		    window.parent.frames["UpperWindow"].GetRestoreLanguageChanges();
		    var i;
		    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
		        for (i = 0; i < window.parent.frames["UpperWindow"].document.forms.frmChange.elements.length; i++) {
		            if (window.parent.frames["UpperWindow"].document.forms.frmChange.elements[i].className == "chkBase") {
		                if ((!window.parent.frames["UpperWindow"].document.forms.frmChange.elements[i].checked) || window.parent.frames["UpperWindow"].document.forms.frmChange.elements[i].indeterminate)
		                    blnAll = false;
		            }
		        }
		    }
		    else {
		        for (i = 0; i < window.parent.frames["UpperWindow"].document.forms("frmChange").all.length; i++) {
		            if (window.parent.frames["UpperWindow"].document.forms("frmChange").all(i).className == "chkBase") {
		                if ((!window.parent.frames["UpperWindow"].document.forms("frmChange").all(i).checked) || window.parent.frames["UpperWindow"].document.forms("frmChange").all(i).indeterminate)
		                    blnAll = false;
		            }
		        }
		    }
			window.parent.frames["UpperWindow"].frmChange.chkAllChecked.checked = blnAll;	

			/*for (i=0;i < window.parent.frames["UpperWindow"].document.all.length;i++)
				{
					if (window.parent.frames["UpperWindow"].document.all(i).className == "chkBase")
						{
						if (window.parent.frames["UpperWindow"].document.all(i).checked || window.parent.frames["UpperWindow"].document.all(i).indeterminate)
							blnNone = false;
						}
				}

		
			if(blnNone)
				{
					alert("Please select at least one image.");
					window.parent.frames["UpperWindow"].frmChange.txtSummary.focus;
				}
			else */
			if((! blnAll) && trim(String(window.parent.frames["UpperWindow"].frmChange.txtSummary.value).toUpperCase()) == "ALL")
				{
					alert("Please update the Image Strategy to reflect the images you have selected.");
					window.parent.frames["UpperWindow"].frmChange.txtSummary.focus;
				}
			else
				{
				cmdCancel.disabled =true;
				cmdOK.disabled =true;
				window.parent.frames["UpperWindow"].frmChange.submit();
				}
		}

}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory">


<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick('<%=Request("pulsarplusDivId")%>')"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>