<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {

	var ValidationFailed = false;
	var count =0;
	var pos;
	var str;
	if (window.parent.frames["UpperWindow"].Release.SelectedMilestone.value == "" )
		{
		ValidationFailed = true;
		window.alert("You must select Milestone to continue");
		}
	else if (window.parent.frames["UpperWindow"].Release.txtFilename.value == ""   && window.parent.frames["UpperWindow"].Release.txtType.value != "1")
		{
		ValidationFailed = true;
		window.parent.frames["UpperWindow"].Release.txtFilename.focus();
		window.alert("Filename is required.");
		}

	else if (window.parent.frames["UpperWindow"].Release.cboReleasePriority.selectedIndex != 1 && window.parent.frames["UpperWindow"].Release.txtReleasePriorityJust.value==""	)
		{
		ValidationFailed = true;
		window.parent.frames["UpperWindow"].Release.txtReleasePriorityJust.focus();
		window.alert("Release justification is required unless Normal priority is selected.");
		}
    else if (window.parent.frames["UpperWindow"].Release.txtReplicator.value != "1" && window.parent.frames["UpperWindow"].Release.txtUserPartner.value != "1" && window.parent.frames["UpperWindow"].Release.txtType.value != "1" && window.parent.frames["UpperWindow"].Release.txtISOFilename.value == "" && (window.parent.frames["UpperWindow"].Release.cboTransferServer.selectedIndex == 0 || window.parent.frames["UpperWindow"].Release.txtTransfer.value == ""))
	    {
	    window.alert("You must upload a deliverable or specify an FTP path to continue.");
		ValidationFailed = true;
	    }
    else if (window.parent.frames["UpperWindow"].Release.txtReplicator.value != "1" && window.parent.frames["UpperWindow"].Release.cboTransferServer.selectedIndex == 0 && window.parent.frames["UpperWindow"].Release.txtUserPartner.value == "1" && window.parent.frames["UpperWindow"].Release.txtType.value != "1" && window.parent.frames["UpperWindow"].Release.txtFilesRequired.value == "true") {
        if (window.parent.frames["UpperWindow"].Release.CloneType.value == "2") 
        {
        window.alert("You must upload a new CVA file.");
        ValidationFailed = true;
    }
    else 
        {
        ValidationFailed = true;
        window.parent.frames["UpperWindow"].Release.cboTransferServer.focus();
        window.alert("Transfer Server is required.");
        }
        }
    else if (window.parent.frames["UpperWindow"].Release.txtReplicator.value != "1" && window.parent.frames["UpperWindow"].Release.txtTransfer.value == "" && window.parent.frames["UpperWindow"].Release.txtUserPartner.value == "1" && window.parent.frames["UpperWindow"].Release.txtType.value != "1" && window.parent.frames["UpperWindow"].Release.txtFilesRequired.value == "true")
		{
		    if (window.parent.frames["UpperWindow"].Release.CloneType.value == "2") {
		        window.alert("You must upload a new CVA file.");
		        ValidationFailed = true;
		    }
		    else {
		        ValidationFailed = true;
		        window.parent.frames["UpperWindow"].Release.txtTransfer.focus();
		        window.alert("Transfer Path is required.");
		    }
        }
	else if (window.parent.frames["UpperWindow"].Release.txtISOFilesRequired.value=="1" && (window.parent.frames["UpperWindow"].Release.chkISOMD5File.checked == false || window.parent.frames["UpperWindow"].Release.chkISOLFSFile.checked == false))
		{
		ValidationFailed = true;
		window.alert("You must provide MD5 and LFS files for each ISO image released.");
		window.parent.frames["UpperWindow"].Release.chkISOMD5File.focus();
		}

    if ((!ValidationFailed) && window.parent.frames["UpperWindow"].Release.txtNotify.value != "") {
        if (!isValidEmailList(window.parent.frames["UpperWindow"].Release.txtNotify.value)) {
            ValidationFailed = true;
            window.parent.frames["UpperWindow"].Release.txtNotify.focus();
        }
    }



	if ((! ValidationFailed) && window.parent.frames["UpperWindow"].txtWarnDeveloper.value=="1")
		{
	    if (! window.confirm("Are you sure you want to release this deliverable to the Release Team?"))
	    		ValidationFailed = true;
		}

		if (!ValidationFailed) {
		    cmdOK.disabled = true;
		    cmdCancel.disabled = true;	
		    window.parent.frames["UpperWindow"].Release.submit();
		}
	
	
}



function isValidEmail(strEmail) {
    var emailReg = "^[\\w-_\.+#]*[\\w-_\.#]\@([\\w]+\\.)+[\\w]+[\\w]$";
    var regex = new RegExp(emailReg);
    return regex.test(strEmail);
}

function isValidEmailList(strList) {
    var AddressArray;

    var i;
    AddressArray = strList.split(";");
    for (i = 0; i < AddressArray.length; i++) {
        if (!isValidEmail(AddressArray[i].replace(/^\s+|\s+$/g, ""))) {
            if (AddressArray[i].replace(/^\s+|\s+$/g, "") != "") {
                alert(AddressArray[i] + " is not a valid email address.");
                return false;
            }
            else if (AddressArray[i].replace(/^\s+|\s+$/g, "") == "" && i != (AddressArray.length - 1)) {
                alert("Missing email address.  Ensure that you have only one semicolon between each address.");
                return false;
            }
        }
    }
    return true;
}

function cmdCancel_onclick() {
		window.parent.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<TABLE  BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</TABLE>

</BODY>
</HTML>