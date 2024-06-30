<%@  language="VBScript" %>
<html>
<head>
    <meta name="VI60_DefaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" type="text/javascript">
<!--

        function ltrim(s) {
            return s.replace(/^\s*/, "")
        }

        function VerifyEmail(src) {
            var emailReg = "^[\\w-_\.]*[\\w-_\.]\@[\\w]\.+[\\w]+[\\w]$";
            var regex = new RegExp(emailReg);
            return regex.test(src);
        }

        String.prototype.trim = function() {
            return this.replace(/^\s+|\s+$/g, "");
        }
        
        String.prototype.ltrim = function() {
            return this.replace(/^\s+/, "");
        }
        
        String.prototype.rtrim = function() {
            return this.replace(/\s+$/, "");
        }


        function VerifySave() {
            var blnSuccess = true;

            with (window.parent.frames["UpperWindow"]) {
                var txtSummary = document.getElementById('txtSummary');
                var txtSpsKitPn = document.getElementById('txtSpareKitNo');
                var txtSaPn = document.getElementById('txtSaNo');
                var txtDescription = document.getElementById('txtDetails');
                var lstProducts = document.getElementById('lstProducts');
            }

            if (txtSummary.value.trim().length == 0) {
                alert("Summary is a required field");
                txtSummary.focus();
                blnSuccess = false;
            }
            else if (txtSpsKitPn.value.trim().length == 0) {
                alert("Spare Kit Pn is a required field");
                txtSpsKitPn.focus();
                blnSuccess = false;
            }
            else if (txtSpsKitPn.value.trim().length > 500) {
                alert("Spare Kit Pn is limited to 500 Characters in length.")
                txtSpsKitPn.focus();
                blnSuccess = false;
            }
            else if (txtSaPn.value.trim().length == 0) {
                alert("Sub Assembly Pn is a required field");
                txtSaPn.focus();
                blnSuccess = false;
            }
            else if (txtSaPn.value.trim().length > 500) {
                alert("Sub Assembly Pn is limited to 500 Characters in length.")
                txtSaPn.focus();
                blnSuccess = false;
            }
            else if (txtDescription.value.trim().length == 0) {
                alert("Description is a required field");
                txtDescription.focus();
                blnSuccess = false;
            }

            strAdding = ""
            Pending = "," + window.parent.frames["UpperWindow"].document.all("txtApproversPending").value;
            if (window.parent.frames["UpperWindow"].ProgramInput.txtType.value != "4")
                ApproverRows = window.parent.frames["UpperWindow"].document.all("ApproverTable").rows.length
            else
                ApproverRows = 0
            for (i = parseInt(window.parent.frames["UpperWindow"].document.all("txtApproversLoaded").value) + 1; i < ApproverRows - 1; i++) {
                if (window.parent.frames["UpperWindow"].document.all("cboApprover" + i).value == 0 && window.parent.frames["UpperWindow"].document.all("chkDelete" + i).checked != true) {
                    blnSuccess = false;
                    window.alert("Approver is required.");
                    window.parent.frames["UpperWindow"].document.all("cboApprover" + i).focus();
                    break;
                }
                else {
                    if (!window.parent.frames["UpperWindow"].document.all("chkDelete" + i).checked) {
                        if (Pending.indexOf("," + window.parent.frames["UpperWindow"].document.all("cboApprover" + i).value + ",") != -1) {
                            blnSuccess = false;
                            window.alert("Can not duplicate approvers.");
                            window.parent.frames["UpperWindow"].document.all("cboApprover" + i).focus();
                            break;
                        }
                        else {
                            strAdding = strAdding + window.parent.frames["UpperWindow"].document.all("cboApprover" + i).value + ",";
                            Pending = Pending + window.parent.frames["UpperWindow"].document.all("cboApprover" + i).value + ",";
                        }
                    }
                }
            }


            window.parent.frames["UpperWindow"].ProgramInput.Approvers2Add.value = strAdding;




            return blnSuccess;
        }

        function cmdEditCancel_onclick() {
            //if (window.confirm ("Are you sure you want to exit this screen without saving your changes?") == true)
            window.parent.close();
        }

        function cmdClear_onclick() {
            window.parent.frames["UpperWindow"].ProgramInput.reset();
            window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.fontStyle = "italic";
            window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.color = "blue";
        }

        function cmdSubmit_onclick() {
            if (VerifySave()) {
                cmdEditCancel.disabled = true;
                cmdSubmit.disabled = true;
                cmdClear.disabled = true;
                window.parent.frames["UpperWindow"].ProgramInput.submit();
            }

        }


//-->
    </script>

    <%
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim TypeID : TypeID = regEx.Replace(Request.QueryString("Type"), "")
    Dim ProdID : ProdID = regEx.Replace(Request.QueryString("ProdID"), "")
    Dim IssueID : IssueID = regEx.Replace(Request.QueryString("ID"), "")
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")%>
</head>
<body bgcolor="ivory">
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <%if IssueID <> "" then%>
            <td>
                <input type="button" value="OK" id="cmdSubmit" name="cmdSubmit" language="javascript"
                    onclick="return cmdSubmit_onclick()">
            </td>
            <td>
                <input style="display: none" type="button" value="Clear Form" id="cmdClear" name="cmdClear"
                    language="javascript" onclick="return cmdClear_onclick()">
            </td>
            <td>
                <input type="button" value="Cancel" id="cmdEditCancel" name="cmdEditCancel" language="javascript"
                    onclick="return cmdEditCancel_onclick()">
            </td>
            <%else%>
            <td>
                <input type="button" value="Submit" id="cmdSubmit" name="cmdSubmit" language="javascript"
                    onclick="return cmdSubmit_onclick()">
            </td>
            <td>
                <input type="button" value="Clear Form" id="cmdClear" name="cmdClear" language="javascript"
                    onclick="return cmdClear_onclick()">
            </td>
            <td>
                <input type="button" value="Cancel" id="cmdEditCancel" name="cmdEditCancel" language="javascript"
                    onclick="return cmdEditCancel_onclick()">
            </td>
            <%end if%>
        </tr>
    </table>
</body>
</html>
