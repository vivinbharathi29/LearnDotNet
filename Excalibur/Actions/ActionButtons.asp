<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->
    
<script src="../Scripts/PulsarPlus.js"></script>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "../includes/Date.asp" -->

/*function Valid_Date(pDate){
	if (isNaN(Date.parse(pDate)))
		return false;
	else
		{
		var NewDate = new Date(pDate);
		if (NewDate.getFullYear() >= 1900 && NewDate.getFullYear() <= 2100)
			return true;
		else
			return false;
		}
}
*/

function IsNumeric(sText)
{
   var ValidChars = "0123456789.-";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
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


function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

function VerifySave(){
	var blnSuccess = true;	
		if (window.parent.frames["UpperWindow"].frmUpdate.cboProject.selectedIndex < 1)
		{
			alert("You must select a project.");
			window.parent.frames["UpperWindow"].frmUpdate.cboProject.focus();
			blnSuccess = false;
		}
		else if (isNaN(trim(window.parent.frames["UpperWindow"].frmUpdate.txtOrder.value)) )
		{
			alert("Your Order must be a number.");
			window.parent.frames["UpperWindow"].frmUpdate.txtOrder.focus();
			blnSuccess = false;
		}
		else if (parseInt(window.parent.frames["UpperWindow"].frmUpdate.txtOrder.value) < 1 && trim(window.parent.frames["UpperWindow"].frmUpdate.txtOrder.value) != "" )
		{
			alert("Your Order must be greater than 0.");
			window.parent.frames["UpperWindow"].frmUpdate.txtOrder.focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmUpdate.txtSummary.value == "")
		{
			alert("You must specify a summary.");
			window.parent.frames["UpperWindow"].frmUpdate.txtSummary.focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate.value != "" && (! isDate(window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate.value)) )
		{
			alert("Invalid date format specified.");
			window.parent.frames["UpperWindow"].frmUpdate.txtTargetDate.focus();
			blnSuccess = false;
		}
		else if (window.parent.frames["UpperWindow"].frmUpdate.txtDuration.value != "" && (! IsNumeric(window.parent.frames["UpperWindow"].frmUpdate.txtDuration.value)) )
		{
			alert("Hours Required must be a number if it is supplied.");
			window.parent.frames["UpperWindow"].frmUpdate.txtDuration.focus();
			blnSuccess = false;
		}
		
		
	
	return blnSuccess;
}

//*****************************************************************
//Description:  onClick, Cancel button closes child window
//Function:     cmdCancel_onclick();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs      
//*****************************************************************
function cmdCancel_onclick() { 
    if (IsFromPulsarPlus()) {
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

function cmdOK_onclick() {
	var blnAll = true;
	var i;
	if (VerifySave())
	{
	    cmdCancel.disabled =true;
	    cmdOK.disabled =true;
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
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>