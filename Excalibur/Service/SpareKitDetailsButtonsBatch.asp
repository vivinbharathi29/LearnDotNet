<%@  language="VBScript" %>
<html>
<head>
    <meta name="VI60_DefaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">

    <script id="clientEventHandlersJS" type="text/javascript">
<!--

        var sBatchUpdSummary = "";

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

        function isDate(sDateValue) {
            var bResult=true;
            
            try {
                var theDate = new Date(sDateValue);
                
                if (theDate == "NaN") bResult = false;

            }
            catch(Error)
            {
                bResult=false;
                alert(Error.Description);
            }

            return bResult;
        }


        function showDatePicker(target) {
            var bResult = false;

            var strID;
            var txtDateField = target;
            strID = window.showModalDialog("../MobileSE/Today/calDraw1.asp", txtDateField.value, "dialogWidth:300px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                txtDateField.value = strID;
                bResult = true;
            }

            return bResult;
        }


        function VerifySave() {

            sBatchUpdSummary = "";
            var bDirty = false;
            var sUpdMatrix = "";

            var oUpdMatrix = window.parent.frames['UpperWindow'].document.getElementById('updateMatrix');
            var oUpdGEOs = window.parent.frames['UpperWindow'].document.getElementById('rdoUpdGEOSYes');
            var oAppIntNotes = window.parent.frames['UpperWindow'].document.getElementById('rdoNotesA');
            var oAppRSLComm = window.parent.frames['UpperWindow'].document.getElementById('rdoCommA');

            var oCSRList = window.parent.frames['UpperWindow'].document.getElementById('spsCsrLevel');
            var oDispList = window.parent.frames['UpperWindow'].document.getElementById('spsDisposition');
            var oWLTList = window.parent.frames['UpperWindow'].document.getElementById('spsWarranty');
            var oLSAList = window.parent.frames['UpperWindow'].document.getElementById('spsLocalStockAdvice');
            var oGEONA = window.parent.frames['UpperWindow'].document.getElementById('spsGeosNa');
            var oGEOLA = window.parent.frames['UpperWindow'].document.getElementById('spsGeosLa');
            var oGEOAPJ = window.parent.frames['UpperWindow'].document.getElementById('spsGeosApj');
            var oGEOEMEA = window.parent.frames['UpperWindow'].document.getElementById('spsGeosEmea');
            var oFSD = window.parent.frames['UpperWindow'].document.getElementById('spsFirstServiceDt');
            var oIntNotes = window.parent.frames['UpperWindow'].document.getElementById('spsNotes');
            var oRSLComm = window.parent.frames['UpperWindow'].document.getElementById('spsComments');

            // AT LEAST ONE ITEM NEEDS TO BE A VALUE OTHER THAN ITS RESPECTIVE DEFAULT
            // CSR Level, Disposition, Warranty Labour Tier, Local Stock Advice, GEOS (4 checkboxes), First Service Dt., and RSL Comments

            // CSR Level
            if (oCSRList.value != "0") {
                sUpdMatrix += "1";
                sBatchUpdSummary += "CSR Level = '" + oCSRList.options[oCSRList.selectedIndex].text + "'\n";
            }
            else sUpdMatrix += "0";
            
            // Disposition
            if (oDispList.value != "0") {
                sUpdMatrix += "1";
                sBatchUpdSummary += "Disposition = '" + oDispList.options[oDispList.selectedIndex].text + "'\n";
            }
            else sUpdMatrix += "0";
            
            // Warranty Labour Tier
            if (oWLTList.value != "0") {
                sUpdMatrix += "1";
                sBatchUpdSummary += "Warranty Labour Tier = '" + oWLTList.options[oWLTList.selectedIndex].text + "'\n";
            }
            else sUpdMatrix += "0";
            
            // Local Stock Advice
            if (oLSAList.value != "0") {
                sUpdMatrix += "1";
                sBatchUpdSummary += "Local Stock Advice = '" + oLSAList.options[oLSAList.selectedIndex].text + "'\n";
            }
            else sUpdMatrix += "0";

            if (oUpdGEOs.checked) {

                sUpdMatrix += "1111";

                // GEOS - NA
                if (oGEONA.checked) sBatchUpdSummary += "GEOS NA = 'true'\n";
                else sBatchUpdSummary += "GEOS NA = 'false'\n";

                // GEOS LA
                if (oGEOLA.checked) sBatchUpdSummary += "GEOS LA = 'true'\n";
                else sBatchUpdSummary += "GEOS LA = 'false'\n";

                // GEOS APJ
                if (oGEOAPJ.checked) sBatchUpdSummary += "GEOS APJ = 'true'\n";
                else sBatchUpdSummary += "GEOS APJ = 'false'\n";

                // GEOS EMEA
                if (oGEOEMEA.checked) sBatchUpdSummary += "GEOS EMEA = 'true'\n";
                else sBatchUpdSummary += "GEOS EMEA = 'false'\n";

            }
            else if (
            (!oUpdGEOs.checked) && ((oGEONA.checked) || (oGEOLA.checked) || (oGEOAPJ.checked) || (oGEOEMEA.checked))
            ){
                if (window.confirm("One or more of the GEOS check boxes are currently flagged.  Do you wish to Update all the selected records with the current GEOS settings?")) {
                    sUpdMatrix += "1111";
                    oUpdGEOs.checked = true;

                    // GEOS - NA
                    if (oGEONA.checked) sBatchUpdSummary += "GEOS NA = 'true'\n";
                    else sBatchUpdSummary += "GEOS NA = 'false'\n";

                    // GEOS LA
                    if (oGEOLA.checked) sBatchUpdSummary += "GEOS LA = 'true'\n";
                    else sBatchUpdSummary += "GEOS LA = 'false'\n";

                    // GEOS APJ
                    if (oGEOAPJ.checked) sBatchUpdSummary += "GEOS APJ = 'true'\n";
                    else sBatchUpdSummary += "GEOS APJ = 'false'\n";

                    // GEOS EMEA
                    if (oGEOEMEA.checked) sBatchUpdSummary += "GEOS EMEA = 'true'\n";
                    else sBatchUpdSummary += "GEOS EMEA = 'false'\n";


                } else {
                    sUpdMatrix += "0000";
                    oGEONA.checked=false;
                    oGEOLA.checked=false;
                    oGEOAPJ.checked = false;
                    oGEOEMEA.checked = false;
                }
            }
            else {
                sUpdMatrix += "0000";
            }

            // First Service Date
            if (oFSD.value.toString().trim().length > 0) {
                if (isDate(oFSD.value.toString().trim())) {
                    sUpdMatrix += "1";
                    sBatchUpdSummary += "First Service Dt. = '" + oFSD.value.toString().trim() + "'\n";
                }
                else {
                    alert("Invalid Date Specified!  Use the Date Picker.");
                    oFSD.value = "";
                    oFSD.style.backgroundColor = "lightsteelblue";
                    oFSD.focus();

                    if (showDatePicker(oFSD)) {
                        sUpdMatrix += "1";
                        sBatchUpdSummary += "First Service Dt. = '" + oFSD.value.toString().trim() + "'\n";
                    }
                    else sUpdMatrix += "0";

                    oFSD.style.backgroundColor = "white";
                }
            } else sUpdMatrix += "0";

            // Internal Notes
            if (oIntNotes.value.toString().trim().length > 0) {
                oIntNotes.value = oIntNotes.value.toString().replace("|", " ").trim(); // use RegExp instead?
                sUpdMatrix += "1";
                if(oAppIntNotes.checked)
                    sBatchUpdSummary += "Internal Notes (Append) = '" + oIntNotes.value + "'\n";
                else
                    sBatchUpdSummary += "Internal Notes (Overwrite)= '" + oIntNotes.value + "'\n";
            }
            else sUpdMatrix += "0";

            // RSL Comments
            if (oRSLComm.value.toString().trim().length > 0) {
                oRSLComm.value=oRSLComm.value.toString().replace("|", " ").trim(); // use RegExp instead?
                sUpdMatrix += "1";
                if (oAppRSLComm.checked)
                    sBatchUpdSummary += "RSL Comments (Append) = '" + oRSLComm.value + "'\n";
                else
                    sBatchUpdSummary += "RSL Comments (Overwrite) = '" + oRSLComm.value + "'\n";
            }
            else sUpdMatrix += "0";

            // Determine if any selections/entries have been made
            if (sUpdMatrix.indexOf("1") > -1) bDirty = true;

            // Update the updMatrix input value
            oUpdMatrix.value = sUpdMatrix;
            
            return bDirty;
        }


        function cmdEditCancel_onclick() {
            //if (window.confirm ("Are you sure you want to exit this screen without saving your changes?") == true)   
            window.parent.frames['UpperWindow'].document.getElementById('action').value = "cancel";
            window.parent.returnValue = "cancel";                     
            window.parent.close();
        }

        function cmdSubmit_onclick() {
            if (VerifySave()) {

                if (window.confirm("UPDATE SUMMARY\n\nThe selected record(s) will be updated in the following manner:\n\n" + sBatchUpdSummary + "\nProceed?") == true) {
                    cmdEditCancel.disabled = true;
                    cmdSubmit.disabled = true;
                    window.parent.frames['UpperWindow'].document.getElementById('action').value = "save";
                    window.parent.frames["UpperWindow"].document.getElementById('frmMain').submit();
                }
            } else {

                alert("At least ONE value must be specified in order to submit a Batch Update!");
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
    Dim SFPN : SFPN =trim(Request.QueryString("SFPN"))
%>
</head>
<body>
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
</body>
</html>
