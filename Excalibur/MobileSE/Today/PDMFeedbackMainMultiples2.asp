<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
    if (IsFromPulsarPlus()) {
        var txtPDMFeedback = document.getElementById("txtPDMFeedback");
        window.parent.parent.parent.PDMFeedbackMainMultiples2Callback(txtPDMFeedback.value);
        ClosePulsarPlusPopup();
    }
    else {
        var txtPDMFeedback = document.getElementById("txtPDMFeedback");
        //alert(txtPDMFeedback.value);
        window.returnValue = txtPDMFeedback.value;
        window.close();
        //Me.body.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", txtPDMFeedback.value));
    }
}

function window_onload() {
    frmPDMFeedback.txtPDMFeedback.focus();
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	FONT-FAMILY=Verdana;
	FONT-SIZE=x-small;
}
</STYLE>
<BODY bgcolor=Ivory LANGUAGE=javascript>

<%
dim Feedback
Feedback=request("Feedback")
%>
<form ID=frmPDMFeedback method=post>
<TABLE width=100%><TR><TD align=center style="font-family:Verdana"><b>Are you sure you want to complete the selected action item(s)?</b></td></tr>
<TR><TD align=center style="font-family:Verdana">PDM Feedback Is <b>Required</b> For Non-actionable Action Items</td></tr></table>
<br />
<input type="text" id=txtPDMFeedback name=txtPDMFeedback value="<%=Feedback%>" maxlength="100" style="position:absolute; width:560px" />
<br />
<br />
<TABLE width=100%><TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr></table>
<INPUT type="hidden" id=txtFeedback name=txtFeedback value="<%=Feedback%>">
</form>
</BODY>
</HTML>
