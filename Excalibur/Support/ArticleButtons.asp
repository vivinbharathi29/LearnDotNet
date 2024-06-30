<%@ Language=VBScript %>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    
<script src="../includes/client/jquery-1.10.2.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function cmdOK_onclick(){
    var MyForm=window.parent.frames["MainWindow"].frmMain;

    if (MyForm.txtTitle.value=="")
        {
        alert("Title is required");
        MyForm.txtTitle.focus();
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
    else if (MyForm.cboCategory.selectedIndex==-1)
        {
        alert("Category is required");
        MyForm.cboCategory.focus();
        }
    else if (MyForm.cboCategory.options[MyForm.cboCategory.selectedIndex].text=="")
        {
        alert("Category is required");
        MyForm.cboCategory.focus();
        }
    else if (MyForm.cboOwner.selectedIndex==-1)
        {
        alert("Owner is required");
        MyForm.cboOwner.focus();
        }
    else if (MyForm.cboOwner.options[MyForm.cboOwner.selectedIndex].text == "") {
        alert("Owner is required");
        MyForm.cboOwner.focus();
    }
    else {
        ajaxurl = "ArticleSave.asp?IsValidation=true&Title=" + MyForm.txtTitle.value + "&Id=" + MyForm.txtID.value;
        $.ajax({
            url: ajaxurl,
            type: "POST",
            success: function (data) {
                if (data.includes("TitleDoesNotExist")) {
                    MyForm.submit();
                    if (window.location.hostname.indexOf("/pulsarplus/") > 0) {
                        window.parent.parent.ClosePulsarPlusPopup();
                    }
                }
                else{
                    alert("Article Title already exists.");
                }
            }
        });        
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