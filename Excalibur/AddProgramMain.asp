<%@  language="VBScript" %>

<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script src="Scripts/PulsarPlus.js"></script>
    <script id="clientEventHandlersJS" language="javascript">
<!--

    var KeyString = "";

    function combo_onfocus() {
        KeyString = "";
    }

    function combo_onclick() {
        KeyString = "";
    }

    function combo_onkeydown() {
        if (event.keyCode == 8) {
            if (String(KeyString).length > 0)
                KeyString = Left(KeyString, String(KeyString).length - 1);
            return false;
        }
    }

    function Left(str, n) {
        if (n <= 0)     // Invalid bound, return blank string
            return "";
        else if (n > String(str).length)   // Invalid bound, return
            return str;                // entire string
        else // Valid bound, return appropriate substring
            return String(str).substring(0, n);
    }


    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        } else {
            if (CheckOpener() === false) {
                parent.window.parent.ClosePropertiesDialog();
            } else {
                window.close();
            }
        }
    }

    function cmdOK_onclick() {
        var strID;
        var strPulsar;
        if (frmProductType.optRequirmentPRL.checked) {
            strPulsar = 1;
        } else {
            strPulsar = 0;
        }
        window.location = "mobilese/today/programs.asp?Pulsar=" + strPulsar
    }

   function window_onload() {
       if (document.getElementById("preferredLayout").value == "" && document.getElementById("preferredLayout").value != 'pulsar2') {
           document.getElementById("cmdCancel").style.display = "none";
       }
   }

    function CheckOpener() {
        //If True, page opened with showModalDialog
        //if False, page opened with JQuery Modal Dialog
        var oWindow = window.dialogArguments;
        return (oWindow == null) ? false : true;
    }
    //-->
    </script>
</head>
<style>
    td
    {
        FONT-FAMILY =Verdana;
        FONT-SIZE =x-small;
    }
</style>
<body bgcolor="Ivory" language="javascript" onload="return window_onload()">

    <%
		dim CurrentDomain
		dim CurrentUser
		dim strEmployees
		dim CurrentUserID
'		CurrentUser = lcase(Session("LoggedInUser"))
	
'		if instr(currentuser,"\") > 0 then
'			CurrentDomain = lcase(left(currentuser, instr(currentuser,"\") - 1))
'			Currentuser = lcase(mid(currentuser,instr(currentuser,"\") + 1))
'		end if
	
	
    %>
    <form id="frmProductType">
        <table width="100%" bgcolor="cornsilk" border="1" bordercolor="tan" cellspacing="0" cellpadding="1"></table>
        <table width="100%">
            <tr>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input id="optRequirmentPRL" name="optRequirmentType" type="radio" value="1" checked />
                                <font face="verdana" size="2">Use Pulsar requirements (PRL)</font>
                                <br />
                                <input id="optRequirmentPDD" name="optRequirmentType" type="radio" value="0" />
                                <font face="verdana" size="2">Use legacy Excalibur requirements (Excel based PDD general requirement)</font>
                                <br />
                                &nbsp;<font face="verdana" size="1.8">Legacy IRS SCMs are supported in the IRS tool</font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <hr/>
        <table width="100%">
            <tr>
                <td align="right">
                  <input type="button" value="OK" id="cmdOK" name="cmdOK" language="javascript" onclick="return cmdOK_onclick()">&nbsp;
                  <%if Request.Cookies("PreferredLayout2") <> "pulsar2" then%>
                   <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" language="javascript" onclick="return cmdCancel_onclick()">
                  <%end if%>
                </td>
            </tr>
        </table>
        <input type="hidden" id="txtEmployeeID" name="txtEmployeeID" value="<%=strEmployeeID%>">
        <input type="hidden" id="preferredLayout" value="<%=Request.Cookies("PreferredLayout2")%>" />
    </form>
</body>
</html>
