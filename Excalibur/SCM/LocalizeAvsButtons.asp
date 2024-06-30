<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>


//function trim( varText)
//    {
//    var i = 0;
//    var j = varText.length - 1;
//    
//	for( i = 0; i < varText.length; i++ )
//		{
//		if( varText.substr( i, 1 ) != " " &&
//			varText.substr( i, 1 ) != "\t")
//		break;
//		}
		
   
//	for( j = varText.length - 1; j >= 0; j-- )
//		{
//		if( varText.substr( j, 1 ) != " " &&
//			varText.substr( j, 1 ) != "\t")
//		break;
//		}
//
//    if( i <= j )
//		return( varText.substr( i, (j+1)-i ) );
//	else
//		return("");
//    }

//function VerifyEmail(src) {
//     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
//     var regex = new RegExp(emailReg);
//     return regex.test(src);
//  }


    function cmdCancel_onclick() {
        if (window.parent.frames["UpperWindow"]) {

            var pulsarplusDivId = document.getElementById("pulsarplusDivId").value;
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.modalDialog.cancel(false);
            }
            else {
                parent.window.parent.modalDialog.cancel(true);
            }
        } else {
            window.parent.close();
        }
    }

function cmdOK_onclick() {
	//var blnAll = true;
	//var blnNone = true;
	
			//window.parent.frames["UpperWindow"].GetLanguageChanges();
			//var i;
			//for (i=0;i < window.parent.frames["UpperWindow"].document.all.length;i++)
			//	{
			//		if (window.parent.frames["UpperWindow"].document.all(i).className == "chkBase")
			//			{
			//			if ((! window.parent.frames["UpperWindow"].document.all(i).checked) || window.parent.frames["UpperWindow"].document.all(i).indeterminate)
			//				blnAll = false;
			//			}
			//	}

			//window.parent.frames["UpperWindow"].frmChange.chkAllChecked.checked = blnAll;	

			//for (i=0;i < window.parent.frames["UpperWindow"].document.all.length;i++)
			//	{
			//		if (window.parent.frames["UpperWindow"].document.all(i).className == "chkBase")
			//			{
			//			if (window.parent.frames["UpperWindow"].document.all(i).checked || window.parent.frames["UpperWindow"].document.all(i).indeterminate)
			//				blnNone = false;
			//			}
			//	}

			
			/*if(blnNone)
				{
					alert("Please select at least one image.");
					window.parent.frames["UpperWindow"].frmChange.txtSummary.focus;
				}
			else 
			if((! blnAll) && trim(String(window.parent.frames["UpperWindow"].frmChange.txtSummary.value).toUpperCase()) == "ALL")
				{
					alert("Please update the Image Strategy to reflect the images you have selected.");
					window.parent.frames["UpperWindow"].frmChange.txtSummary.focus;
				}
			else*/
            window.parent.frames["UpperWindow"].inlineProgressBarDiv.style.display = '';
			cmdCancel.disabled =true;
			cmdOK.disabled = true;
			window.parent.frames["UpperWindow"].frmChange.hidFunction.value = "save";
			window.parent.frames["UpperWindow"].frmChange.submit();

			setTimeout(function () {
			    if (window.parent.frames["UpperWindow"]) {
			        var pulsarplusDivId = document.getElementById("pulsarplusDivId").value;
			        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
			            parent.window.parent.modalDialog.cancel(true);
			            document.getElementById('SupplyChainLegacy').contentDocument.location.reload(true);
			        }
			        else {
			            parent.window.parent.ReloadLocalizeAVs("YES");
			            parent.window.parent.modalDialog.cancel(true);
			        }
			    } else {
			        window.parent.close();
			    }
			}, 1000);
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
    <input id="pulsarplusDivId" type="hidden" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>