<%@ Language=VBScript %>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "includes/Date.asp" -->

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
    var MyForm = window.parent.frames["UpperWindow"].frmMain;
    var i;
    var VersionsChecked = false;

    if (MyForm.optChoose.checked)
        {
        if (MyForm.txtAllVersions.value != "")
            VersionsChecked = true;
        else
            {
            if (typeof(MyForm.chkVersion.length)=="undefined")
                {
                if (MyForm.chkVersion.checked)
                    VersionsChecked = true;
                }
            else
                {
                   for (i=0;i<MyForm.chkVersion.length;i++)
                    if (MyForm.chkVersion[i].checked)
                        VersionsChecked = true;
                }
            }
        }


    if (MyForm.txtID.value== "" && MyForm.optID.checked)
        {
        alert("You must enter or select at least one deliverable version to continue.");
        MyForm.txtID.focus();
        return false;
	    }
    else if (!(/^ *[0-9]{1,10} *( *,{1} *[0-9]{1,10} *)*$/.test(MyForm.txtID.value)) && MyForm.optID.checked)
        {
        alert("The list of Version IDs must be a comma-separated list of numbers.");
        MyForm.txtID.focus();
        return false;
	    }
    else if (MyForm.cboRoot.selectedIndex < 1 && MyForm.optChoose.checked)
        {
        alert("You must select a root deliverable to continue.");
        MyForm.cboRoot.focus();
        return false;
	    }
    else if ((!VersionsChecked)  && MyForm.optChoose.checked)
        {
        alert("You must select at least one deliverable version to continue.");
        MyForm.cboRoot.focus();
        return false;
	    }
    else
        return true;
}

function cmdCancel_onclick() {
		window.parent.close();
}

function cmdOK_onclick() {
	if (VerifySave())
		{
			cmdCancel.disabled =true;
			cmdOK.disabled =true;
			window.parent.frames["UpperWindow"].cmdOK_onclick();
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