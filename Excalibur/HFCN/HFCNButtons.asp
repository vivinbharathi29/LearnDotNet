<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

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
	 if (window.parent.frames["MainWindow"].AddHFCN.txtTitle.value == "")
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtTitle.focus();
		window.alert("Title is required.");
		}
	 else if (window.parent.frames["MainWindow"].AddHFCN.txtVersion.value == "")
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtVersion.focus();
		window.alert("Version is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.cboDeliverable.selectedIndex < 1)
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.cboDeliverable.focus();
		window.alert("Deliverable is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.cboCategory.selectedIndex < 1)
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.cboCategory.focus();
		window.alert("Category is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.cboVendor.selectedIndex < 1)
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.cboVendor.focus();
		window.alert("Vendor is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.cboDeveloper.selectedIndex < 1)
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.cboDeveloper.focus();
		window.alert("Developer is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.txtEmail.value == "")
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtEmail.focus();
		window.alert("Email List is required.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.txtLocation.value == "")
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtLocation.focus();
		window.alert("The location field should tell the use how to find this deliverable.");
		}
	else if (window.parent.frames["MainWindow"].AddHFCN.txtChanges.value == "" && window.parent.frames["MainWindow"].AddHFCN.lstOTS.length == 0)
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtOTS.focus();
		window.alert("You must describe what is new in this release.  Please fill out the Obvervations Fixed or Other Changes fields.");
		}

	else if (window.parent.frames["MainWindow"].AddHFCN.chkProdInd.checked == false && window.parent.frames["MainWindow"].AddHFCN.lstSelectedProd.length == 0 )
		{
			ValidationFailed = true;
			window.parent.frames["MainWindow"].AddHFCN.lstAvailableProd.focus();
			window.alert("Select Supported Products or check Product Independent.");
		}

	if (! ValidationFailed)
		{
		cmdOK.disabled = true;
		cmdCancel.disabled = true;	
		window.parent.frames["MainWindow"].AddHFCN.txtCategory.value = window.parent.frames["MainWindow"].AddHFCN.cboCategory.options[window.parent.frames["MainWindow"].AddHFCN.cboCategory.selectedIndex].text;
		window.parent.frames["MainWindow"].AddHFCN.txtVendor.value = window.parent.frames["MainWindow"].AddHFCN.cboVendor.options[window.parent.frames["MainWindow"].AddHFCN.cboVendor.selectedIndex].text;
		window.parent.frames["MainWindow"].FindProductChanges();
		window.parent.frames["MainWindow"].FindOTSChanges();
		window.parent.frames["MainWindow"].AddHFCN.submit();
		}	
	
	
}



function cmdCancel_onclick() {
	//if (window.confirm ("Are you sure you want to exit this screen without releasing this deliverable?") == true)
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