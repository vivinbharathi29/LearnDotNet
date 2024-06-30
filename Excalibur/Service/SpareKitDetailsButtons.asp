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

            var spsPartNo = window.parent.frames['UpperWindow'].document.getElementById('spsPartNo');
            var spsDescription = window.parent.frames['UpperWindow'].document.getElementById('spsDescription');
            spsPartNo.style.backgroundColor = "white";

/*
            with (window.parent.frames['UpperWindow']) {
                var spsPartNo = document.getElementById('spsPartNo');
                var spsDescription = document.getElementById('spsDescription');
            }
*/
            if (spsPartNo.value.trim().length == 0 && spsDescription.value.trim().length == 0) {
                alert("Either a Part Number or Description is required.");
                spsPartNo.focus();
                spsPartNo.style.backgroundColor = "lightsteelblue";
                blnSuccess = false;
            }else if(spsDescription.value.trim().length>40)
            {
                alert("Description can not exceed 40 characters.");
                spsDescription.focus();
                spsDescription.style.backgroundColor = "lightsteelblue";
                blnSuccess = false;
            }

            return blnSuccess;
        }

        function cmdEditCancel_onclick() {
            //if (window.confirm ("Are you sure you want to exit this screen without saving your changes?") == true)   
            window.parent.frames['UpperWindow'].document.getElementById('action').value = "cancel";
            window.parent.returnValue = "cancel";                     
            window.parent.close();
        }

        function cmdSubmit_onclick() {
            if (VerifySave()) {
                cmdEditCancel.disabled = true;
                cmdSubmit.disabled = true;
                window.parent.frames['UpperWindow'].document.getElementById('action').value = "save";
                //window.parent.close();
                window.parent.frames["UpperWindow"].document.getElementById('frmMain').submit();
            }

        }


//-->
    </script>

<%
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim PVID : PVID = regEx.Replace(Request.QueryString("PVID"), "")
    Dim DRID : DRID = regEx.Replace(Request.QueryString("DRID"), "")
    Dim SKID : SKID = regEx.Replace(Request.QueryString("SKID"), "")
    Dim CID : CID = regEx.Replace(Request.QueryString("CID"), "")
    regEx.Pattern = "[^0-9-]"
    Dim SFPN : SFPN = trim(Request.QueryString("SFPN"))

    Dim strMode: strMode=Request.QueryString("M")
%>
</head>
<body>
<% IF strMode="0" THEN %>
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <td>
                <input type="button" value="Submit" id="cmdSubmit" name="cmdSubmit" onclick="cmdSubmit_onclick()" />
            </td>
            <td>
                <input type="button" value="Cancel" id="cmdEditCancel" name="cmdEditCancel" onclick="cmdEditCancel_onclick()" />
            </td>
        </tr>
    </table>
<% ELSE %>
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <td>
                
            </td>
            <td>
                <input type="button" value="Close" id="cmdEditCancel" name="cmdEditCancel" onclick="cmdEditCancel_onclick()" />
            </td>
        </tr>
    </table>
<% END IF %>
</body>
</html>
