<%@  language="VBScript" %>
<html>
<head>
    <meta name="VI60_DefaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" language="javascript">
<!--
    function cmdCancel_onclick() {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }

    function cmdOK_onclick() {
        var ReleaseNames = [];
        var ReleaseIDs = []; 

        if (typeof (window.parent.frames["UpperWindow"].frmMain.chkRelease) == "undefined") {
            alert("Please add at least one item to product release list.");
            return;
        }

        if (typeof (window.parent.frames["UpperWindow"].frmMain.chkRelease.length) == "undefined") {
            if (window.parent.frames["UpperWindow"].frmMain.chkRelease.checked) {
                ReleaseNames.push(window.parent.frames["UpperWindow"].frmMain.chkRelease.ReleaseName);
                ReleaseIDs.push(window.parent.frames["UpperWindow"].frmMain.chkRelease.value);
            }
        }
        else {
            for (i = 0; i < window.parent.frames["UpperWindow"].frmMain.chkRelease.length; i++) {
                if (window.parent.frames["UpperWindow"].frmMain.chkRelease[i].checked) {
                    ReleaseNames.push(window.parent.frames["UpperWindow"].frmMain.chkRelease[i].getAttribute('releasename'));

                    var tmpReleaseID = window.parent.frames["UpperWindow"].frmMain.chkRelease[i].getAttribute('value') + "~";
                    var row = window.parent.frames["UpperWindow"].frmMain.getElementsByTagName("tr")[i + 2];

                    if (row.cells[2].firstChild != null) {
                        tmpReleaseID += row.cells[2].firstChild.value;
                    }

                    ReleaseIDs.push(tmpReleaseID);
                }
            }
        }

        if (ReleaseNames.length <= 0 || ReleaseIDs.length <= 0) {
            alert("Please select at least one release");
            return;
        }

        parent.window.parent.AddReleaseResult(ReleaseNames.join(","), ReleaseIDs.join(","));
        parent.window.parent.modalDialog.cancel();
    }

//-->
    </script>
</head>
<body bgcolor="ivory">
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <td>
                <input type="button" value="OK" id="cmdOK" name="cmdOK" language="javascript" onclick="return cmdOK_onclick()">
            </td>
            <td>
                <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" language="javascript" onclick="return cmdCancel_onclick()">
            </td>
        </tr>
    </table>
</body>
</html>
