<%@ Language=VBScript %>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function cmdOK_onclick(){
    var MyForm=window.parent.frames["MainWindow"].frmMain;

    if (MyForm.txtName.value=="")
        {
        alert("Category name is required");
        MyForm.txtName.focus();
        }
    else if (MyForm.cboProject.selectedIndex==-1)
        {
        alert("Project is required");
        MyForm.cboProject.focus();
        }
    else if (MyForm.cboProject.options[MyForm.cboProject.selectedIndex].text=="")
        {
        alert("Project is required");
        MyForm.cboProject.focus();
        }
    else if (MyForm.cboOwner.selectedIndex==-1)
        {
        alert("Owner is required");
        MyForm.cboOwner.focus();
        }
    else if (MyForm.cboOwner.options[MyForm.cboOwner.selectedIndex].text=="")
        {
        alert("Owner is required");
        MyForm.cboOwner.focus();
        }
    else
        {
        MyForm.txtProjectName.value=MyForm.cboProject.options[MyForm.cboProject.selectedIndex].text;
        MyForm.submit();
        if (window.location.hostname.indexOf("/pulsarplus/") > 0) {
            window.parent.parent.ClosePulsarPlusPopup();
        }
    }
    
}

function cmdCancel_onclick() {
    //window.parent.close();
    if (window.location.hostname.indexOf("/pulsarplus/") > 0) {
        window.parent.parent.ClosePulsarPlusPopup();
    }
    else {
        parent.window.parent.modalDialog.cancel(false);
    }
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