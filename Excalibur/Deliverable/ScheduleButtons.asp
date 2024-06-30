<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<script src="../Scripts/PulsarPlus.js"></script>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../includes/Date.asp" -->




function ltrim ( s ) 
{ 
	return s.replace( /^\s*/, "" ) 
} 

function VerifyEmail(src) {
     var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
     var regex = new RegExp(emailReg);
     return regex.test(src);
  }

function y2k(number) {
	if (number < 50)
		return number +2000;
	else if (number <100)
		return number+1900;
	else		
		return number;
 
}

function daysElapsed(date1,date2) {
    var difference =
        Date.UTC(y2k(date1.getYear()),date1.getMonth(),date1.getDate(),0,0,0)
      - Date.UTC(y2k(date2.getYear()),date2.getMonth(),date2.getDate(),0,0,0);
    return difference/1000/60/60/24;
}



function VerifySave(){
	var blnSuccess = true;	
	var i;
	var d;
	var d1;
	for (i=0;i<window.parent.frames["UpperWindow"].document.all.length;i++)
		if(window.parent.frames["UpperWindow"].document.all(i).className=="text")
			{
			if(! isDate(window.parent.frames["UpperWindow"].document.all(i).value))
				{
				blnSuccess=false;
				alert("You must supply valid dates for all milestones.\r\rFormat Required: mm/dd/yyyy");
				window.parent.frames["UpperWindow"].document.all(i).focus();
				}
			if (blnSuccess)
				{
				if (typeof(d) != "undefined")
					{
					d1 = new Date(window.parent.frames["UpperWindow"].document.all(i).value);
					
					if (daysElapsed(d1,d)<0)
						{
						alert("This date can not be before the date in the previous workflow step.");
						window.parent.frames["UpperWindow"].document.all(i).focus();
						blnSuccess=false;
						}
					}
				d = new Date(window.parent.frames["UpperWindow"].document.all(i).value);
				}
			
			}
	
	return blnSuccess;
}

function cmdCancel_onclick() {
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        var iframeName = parent.window.name;
        if (iframeName != '') {
            parent.window.parent.ClosePopUp();
        } else {
            window.parent.close();
        }
    }
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			//if (typeof(window.parent.frames["UpperWindow"].frmRequirement.txtSpecification) != "undefined")
			//	window.parent.frames["UpperWindow"].frmRequirement.txtSpecification.value = window.parent.frames["UpperWindow"].frames["myEditor"].document.body.innerHTML;
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