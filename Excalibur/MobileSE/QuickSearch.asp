<%@  language="VBScript" %>
<%Option Explicit%>
<html>
<head>
    <title>QuickSearch</title>
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css" />
    <!-- #include file="../includes/bundleConfig.inc" -->

    <script type="text/javascript">
        function gochange_onmouseover() {
            window.event.srcElement.style.cursor = "hand"
        }

        function goSearch_onclick() {
            var txtSearch = document.getElementById("txtSearch");
            if (txtSearch.value != "") {
                //window.location.href = "..\..\MobileSE\Today\find.asp?Find=" + escape(txtSearch.value) + "&Type=Part";
                //window.location.href = "Today/find.asp?Find=" + escape(txtSearch.value) + "&Type=Part";
                window.open("Today/find.asp?Find=" + escape(txtSearch.value) + "&Type=Part", "", "width=" + GetWindowSize('width') + ",height=600,toolbar=0,resizable=1,scrollbars=1");

                if (window.location != window.parent.location) {
                    // iframe
                    window.parent.modalDialog.cancel();
                }
                else {
                    // no iframe
                    window.close();
                }
            }
            else {
                window.alert("Please enter a part number first.");
                txtSearch.focus();
            }
        }

        function txtSearch_onkeypress() {
            if (window.event.keyCode == 13) {
                goSearch_onclick();
            }
        }

    </script>

</head>
<body bgcolor="cornsilk">
    <br />
    <table cellspacing="1" cellpadding="1" border="0">
        <tr>
            <td valign="top">
                <table>
                    <tr>
                        <td nowrap valign="top">
                            &nbsp;&nbsp;&nbsp;<font face="Verdana" size="2"><strong>Part Number:</strong></font>
                        </td>
                        <td>
                            <input id="txtSearch" name="txtSearch" style="width: 92px; height: 22px" size="13"
                                language="javascript" onkeypress="return txtSearch_onkeypress()" />&nbsp;<a><img
                                    id="gosearch" border="0" src="Today\images\go.gif" width="23" height="20"
                                    language="javascript" onmouseover="return gochange_onmouseover()" onclick="return goSearch_onclick()" /></a>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
