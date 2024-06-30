<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<!-- #include file="../includes/bundleConfig.inc" -->
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function VerifyEmail(src) 
{
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
}

function VerifyStatus()
{
	with (window.parent.frames["UpperWindow"].frmStatus)
	{
        //WILL DEFINE AFTER FEEDBACK FROM USERS
		//switch (cboStatus.value)
		//{
		//	case 'L' :
		//	    return true;
		//		break;
		//	case 'NS' :
		//		return true;
		//		break;
		//	case 'O' :
		//	    if (!validateDateInput(txtProjectedDate, 'Availability Date')) {
        //            alert("Please enter an Availability Date.")
		//	        return false;
		//	    }
		//		break;
		//	case 'C' :
		//		break;
		//	case 'P' :
		//		if (!validateTextInput(txtNote, 'Notes')){	return false; }
		//		break;
		//	default :
		//		break;
		//}
		if (!validateTextInput(cboPorDcr, 'Added By')){	return false; }
	}
	return true;
}

function VerifyLeverage()
{
	var blnSuccess = true;
	return blnSuccess;
}

function cmdCancel_onclick() {
    var iframeName = parent.window.name;
    if (iframeName != '') {
        parent.window.parent.CloseIframeDialog();
    } else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}

function cmdOK_onclick() 
{
	var blnAll = true;
	var i;
	var sReturnValue;
	
	if (window.parent.frames["UpperWindow"].frmStatus)
	{
		if (VerifyStatus())
		{
			sReturnValue = window.parent.frames["UpperWindow"].frmStatus.cboStatus.value;
			if (sReturnValue == 'O' && window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value != '')
			{
				sReturnValue = window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value;
			}
			if (sReturnValue == 'SU' && window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value != '') {
			    sReturnValue = window.parent.frames["UpperWindow"].frmStatus.txtProjectedDate.value;
			}
			//window.frmButtons.cmdCancel.disabled =true;
			//window.frmButtons.cmdOK.disabled =true;
			window.returnValue = sReturnValue;
			window.parent.frames["UpperWindow"].frmStatus.hidSave.value=true;
		    //window.parent.frames["UpperWindow"].frmStatus.hidClose.value=true;

		    //08/31/2016 Mailichi - Save selected products and countries the batch update popup
			window.parent.frames["UpperWindow"].GetSelectedProducts();
			window.parent.frames["UpperWindow"].GetSelectedCountries();

			if (window.parent.frames["UpperWindow"].frmStatus.hidSelectedProducts.value == "") {
			    alert("Please select applicable products from the list");
			    return false;
			}

			if (window.parent.frames["UpperWindow"].frmStatus.hidSelectedCountries.value == "") {
			    alert("Please select applicable countries from the list");
			    return false;
			}

			if (window.parent.frames["UpperWindow"].frmStatus.cboPorDcr.value == 'DCR') {
                var cboDCRs = window.parent.frames["UpperWindow"].document.getElementById("cboDcr");
                if (cboDCRs.options[cboDCRs.selectedIndex].value == "") {
                    alert("Please select a DCR");
                    return false;
                }
			}
			
			window.parent.frames["UpperWindow"].frmStatus.hidSave.value = true;
			window.parent.frames["UpperWindow"].frmStatus.hidClose.value = true;
			window.parent.frames["UpperWindow"].frmStatus.submit();
		}
	}
	else if (window.parent.frames["UpperWindow"].frmLeverage)
	{
		if (VerifyLeverage())
		{
			window.parent.frames["UpperWindow"].frmLeverage.SaveMode.value=true;
			window.parent.frames["UpperWindow"].frmLeverage.submit();
		}
	}
		
	return;
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">
<FORM id="frmButtons"  action=AgencyButtons.asp method=post>
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Save" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>
</FORM>
</body>
</html>