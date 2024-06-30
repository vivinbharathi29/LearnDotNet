<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
td{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
}
</STYLE>
    <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        } else {
            window.returnValue = "Cancel";
            window.close();
        }
    
}

function cmdOK_onclick() {
    var txtOverrideComment = document.getElementById("txtOverrideComment");
    if (txtOverrideComment.value.length == 0) {
        alert("Comment is required for Not Actionable action items");
    }
    else {
        if (IsFromPulsarPlus()) {            
            window.parent.parent.GetNotActionableCommentCallback(txtOverrideComment.value);
            ClosePulsarPlusPopup();
        } else {

            window.returnValue = txtOverrideComment.value;
            window.close();
        }

    }
}

function window_onload() {
    frmOverrideComment.txtOverrideComment.focus();
}

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript>

<form ID=frmOverrideComment method=post>
<TABLE width=100%><TR><TD align=center style="font-family:Verdana"><b>Are you sure you want to set the selected Features(s) to Not Actionable?</b></td></tr>
<TR><TD align=center style="font-family:Verdana">Comment Is <b>Required</b> For Non-actionable Action Items</td></tr></table>
<br />
<input type="text" id=txtOverrideComment name=txtOverrideComment value="" maxlength="100" style="position:absolute; width:560px" />
<br />
<br />
<TABLE width=100%><TR><TD align=right><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">&nbsp;<INPUT type="button" value="Save" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></td></tr></table>
</form>
</BODY>
</HTML>
