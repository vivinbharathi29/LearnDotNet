<%@ Language=VBScript %>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdOK_onclick(){
    var MyForm=window.parent.frames["MainWindow"].frmMain;

    if (MyForm.txtSubject.value=="")
        {
        alert("Question/Request is required");
        MyForm.txtSubject.focus();
        }
    else if (MyForm.cboCategory.selectedIndex==-1)
        {
        alert("Category is required");
        MyForm.cboCategory.focus();
        }
    else if (MyForm.cboStatus.options[MyForm.cboStatus.selectedIndex].text == "Closed" && MyForm.txtResolution.value=="")
        {
        alert("Response is required");
        MyForm.txtResolution.focus();
        }
    else
        {
        MyForm.txtProjectName.value= MyForm.cboProject.options[MyForm.cboProject.selectedIndex].text;
        MyForm.submit();
        if (window.location.hostname.indexOf("/pulsarplus/") > 0) {
            window.parent.parent.ClosePulsarPlusPopup();
        }
    }
   
}

function cmdCancel_onclick() {
    if (window.location.hostname.indexOf("/pulsarplus/") > 0) {
        window.parent.parent.ClosePulsarPlusPopup();
    }
    else {
        parent.window.parent.modalDialog.cancel(false);
    }
    //window.parent.opener='X';
    //window.parent.open('','_parent','')
    //window.parent.close(); 

   
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