<%@  language="VBScript" %>
<html>
<head>
    <meta name="VI60_DefaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">

    <script id="clientEventHandlersJS" language="javascript">
        function cmdCancel_onclick() {
            var pulsarplusDivId = document.getElementById('hdnTabName');
            if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.closeExternalPopup();
            }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel();
                } else {
                    window.parent.close();
                }
            }
        }

        function cmdOK_onclick() {
            if (window.parent.frames["UpperWindow"].VerifySave()) {
                cmdCancel.disabled = true;
                cmdOK.disabled = true;
                window.parent.frames["UpperWindow"].frmCountries.submit();
            }

        }
    </script>

</head>
<body bgcolor="ivory">
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <td>
                <input type="button" value="OK" id="cmdOK" name="cmdOK" language="javascript" onclick="return cmdOK_onclick()"></td>
            <td>
                <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" language="javascript" onclick="return cmdCancel_onclick()"></td>
        </tr>
    </table>
    <input type="hidden" id="hdnTabName" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>
