<%@ Language=VBScript %>



<HTML>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    //*****************************************************************
    //Description:  Return value from modal dialog to parent page
    //Function:     cmdOK_onclick();
    //Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs      
    //*****************************************************************
    function cmdOK_onclick() {        
        if (window.parent.frames["UpperWindow"].txtTo.value == "") {
            window.parent.frames["UpperWindow"].cmdTo_onclick();
        }

        //original - window.returnValue = window.parent.frames["UpperWindow"].txtTo.value;
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            var sFieldID = parent.window.parent.globalVariable.get('email_field');
            //User modalDialog.returnValue function to update parent page's input field with selection
            parent.window.parent.modalDialog.saveValue(sFieldID, window.parent.frames["UpperWindow"].txtTo.value);
            parent.window.parent.modalDialog.cancel();
        }
        else if (window.dialogArguments=="fromRoot") {            
            // child window.
            window.returnValue = window.parent.frames["UpperWindow"].txtTo.value;
            window.close();
        }
        else {
            window.parent.close();
        }
    }



    function cmdCancel_onclick() {
        //if (window.confirm ("Are you sure you want to exit this screen without releasing this deliverable?") == true)
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel(false);
        } else {
            window.parent.close();
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