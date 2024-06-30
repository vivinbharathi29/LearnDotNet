<%@  language="VBScript" %>

<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Purpose:	Pulsar RTM Wizard, main page in the <frame>.
'''Created By:	? , ?
'''Modified By: Herb, 01/29/2016, PBI 16008 Add SCMX standalone RTM.
'''             Herb, 03/11/2016, PBI 16623 Add "SCMX comments" input field.
'''             Herb, 05/27/2016, PBI 20383 Add Firmware standalone RTM.
'''             VHarris, 11/22/2016, PBI 29646 Miscellaneous: Convert Additional Dialogs 
'''             Herb, 07/28/2017, PBI 144920: Able to select Images Localization based on tiering when create RTM Document for Pulsar Products
'''             Herb, 08/30/2017, urgent fix for FailoverServer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Dim AppRoot : AppRoot = Session("ApplicationRoot")
	  
%>

<html>
<head>
    <title>Product RTM</title>
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="../includes/client/json2.js"></script>
    <script type="text/javascript" src="../includes/client/json_parse.js"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
<!--
    function UploadZip(ID) {
        //save ID for return function: ---
        globalVariable.save(ID, 'main_uploadzip_ID');

        var sURL = "<%=AppRoot %>/PMR/SoftpaqFrame.asp?Title=Upload SCMX File&Page=<%=AppRoot %>/common/fileupload.aspx&KeepLocal=true";
        modalDialog.open({ dialogTitle: 'Upload SCMX File', dialogURL: '' + sURL + '', dialogHeight: 250, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
    }

    function UploadZip_return(strPath) {
        var strServer;
        var PathArray;
        var ID = globalVariable.get('main_uploadzip_ID');

        if (typeof (strPath) != "undefined") {
            PathArray = strPath.split("|");

            $("#UploadAddLinks" + ID).hide();
            $("#UploadRemoveLinks" + ID).show();
            $("#txtUploadPath" + ID).val(PathArray[0].substr(0, PathArray[0].lastIndexOf("\\")));
            $("#txtAttachmentPath" + ID).val(PathArray[1]);
            $("#UploadPath" + ID).text(PathArray[0].substr(PathArray[0].lastIndexOf("\\") + 1, PathArray[0].length));
        }
    }

    function RemoveUpload(ID) {
        $("#UploadAddLinks" + ID).show();
        $("#UploadRemoveLinks" + ID).hide();
        $("#txtAttachmentPath" + ID).val("");
        $("#txtUploadPath" + ID).val("");
        $("#UploadPath" + ID).text("");
    }

    function txtRTMComments_onfocus() {
        frmMain.txtRTMComments.style.fontStyle = "normal";
        frmMain.txtRTMComments.style.color = "black";
        if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value) {
            frmMain.txtRTMComments.select();
        }
    }


    function txtRTMComments_onblur() {
        if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value) {
            frmMain.txtRTMComments.style.fontStyle = "italic";
            frmMain.txtRTMComments.style.color = "blue";
        }
    }

    function txtRestoreComments_onfocus() {
        frmMain.txtRestoreComments.style.fontStyle = "normal";
        frmMain.txtRestoreComments.style.color = "black";
        if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value)
            frmMain.txtRestoreComments.select();
    }


    function txtRestoreComments_onblur() {
        if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value) {
            frmMain.txtRestoreComments.style.fontStyle = "italic";
            frmMain.txtRestoreComments.style.color = "blue";
        }
    }

    function txtImageComments_onfocus() {
        frmMain.txtImageComments.style.fontStyle = "normal";
        frmMain.txtImageComments.style.color = "black";
        if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value)
            frmMain.txtImageComments.select();
    }

    function txtImageComments_onblur() {
        if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value) {
            frmMain.txtImageComments.style.fontStyle = "italic";
            frmMain.txtImageComments.style.color = "blue";
        }
    }

    function txtBIOSComments_onfocus() {
        frmMain.txtBIOSComments.style.fontStyle = "normal";
        frmMain.txtBIOSComments.style.color = "black";
        if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value)
            frmMain.txtBIOSComments.select();
    }


    function txtBIOSComments_onblur() {
        if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value) {
            frmMain.txtBIOSComments.style.fontStyle = "italic";
            frmMain.txtBIOSComments.style.color = "blue";
        }
    }

    function txtFWComments_onfocus() {
        frmMain.txtFWComments.style.fontStyle = "normal";
        frmMain.txtFWComments.style.color = "black";
        if (frmMain.txtFWComments.value == frmMain.txtFWCommentsTemplate.value)
            frmMain.txtFWComments.select();
    }


    function txtFWComments_onblur() {
        if (frmMain.txtFWComments.value == frmMain.txtFWCommentsTemplate.value) {
            frmMain.txtFWComments.style.fontStyle = "italic";
            frmMain.txtFWComments.style.color = "blue";
        }
    }

    function txtPatchComments_onfocus() {
        frmMain.txtPatchComments.style.fontStyle = "normal";
        frmMain.txtPatchComments.style.color = "black";
        if (frmMain.txtPatchComments.value == frmMain.txtPatchCommentsTemplate.value)
            frmMain.txtPatchComments.select();
    }


    function txtPatchComments_onblur() {
        if (frmMain.txtPatchComments.value == frmMain.txtPatchCommentsTemplate.value) {
            frmMain.txtPatchComments.style.fontStyle = "italic";
            frmMain.txtPatchComments.style.color = "blue";
        }
    }

    function cmdAdd_onclick() {
        modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../Email/AddressBook.asp?AddressList=' + frmMain.txtNotify.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('txtNotify', 'email_field');
    }

    function trim(varText) {
        var i = 0;
        var j = varText.length - 1;

        for (i = 0; i < varText.length; i++) {
            if (varText.substr(i, 1) != " " &&
                varText.substr(i, 1) != "\t")
                break;
        }


        for (j = varText.length - 1; j >= 0; j--) {
            if (varText.substr(j, 1) != " " &&
                varText.substr(j, 1) != "\t")
                break;
        }

        if (i <= j)
            return (varText.substr(i, (j + 1) - i));
        else
            return ("");
    }



    var KeyString = "";

    function combo_onkeypress() {
        if (event.keyCode == 13) {
            KeyString = "";
        }
        else {
            KeyString = KeyString + String.fromCharCode(event.keyCode);
            event.keyCode = 0;
            var i;
            var regularexpression;

            for (i = event.srcElement.length - 1; i >= 0; i--) {
                regularexpression = new RegExp("^" + KeyString, "i")
                if (regularexpression.exec(event.srcElement.options[i].text) != null) {
                    event.srcElement.selectedIndex = i;
                };

            }
            return false;
        }
    }

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

    var isDate = function (date) {
        return ((new Date(date)).toString() !== "Invalid Date") ? true : false;
    }

    function ValidateTab(strState) {
        var blnSuccess;
        var FocusOn;
        var i;
        var blnFound;
        var intCount;
        var strNewVersion = "";
        var BIOSArray = new Array();
        var RestoreArray = new Array();
        var PatchArray = new Array();
        var FWArray = new Array();

        blnSuccess = true;

        switch (strState) {
            case "General":
                blnFound = false;
                for (i = 0; i < cboTitles.length; i++)
                    if (cboTitles[i].text.toLowerCase().replace(/^\s+|\s+$/g, "") == frmMain.txtRTMName.value.toLowerCase().replace(/^\s+|\s+$/g, "")) {
                        if (frmMain.txtProductRTMID.value != "0") {
                            blnFound = false;
                        }
                        else {
                            blnFound = true;
                        }
                        break;
                    }

                if ((frmMain.txtRTMName.value == "") && blnSuccess) {
                    window.alert("RTM Title is required.");
                    FocusOn = frmMain.txtRTMName;
                    blnSuccess = false;
                }
                else if ((frmMain.txtRTMName.value != "" && blnFound) && blnSuccess) {
                    window.alert("The RTM tile you entered was used on a previous RTM for this Product.");
                    FocusOn = frmMain.txtRTMName;
                    blnSuccess = false;
                }
                else if ((frmMain.txtRTMDate.value == "") && blnSuccess) {
                    window.alert("RTM Date is required.");
                    FocusOn = frmMain.txtRTMDate;
                    blnSuccess = false;
                }
                else if ((!isDate(frmMain.txtRTMDate.value)) && blnSuccess) {
                    window.alert("RTM Date must be a valid Date format.");
                    FocusOn = frmMain.txtRTMDate;
                    blnSuccess = false;
                }
                else if ((!frmMain.chkBIOS.checked && !frmMain.chkRestore.checked && !frmMain.chkImages.checked && !frmMain.chkPatch.checked && !frmMain.chkSCMX.checked && !frmMain.chkFW.checked) && blnSuccess) {
                    window.alert("You must select the items to RTM.");
                    FocusOn = frmMain.txtRTMName;
                    blnSuccess = false;
                }
                else if ((frmMain.chkBIOS.checked && (!frmMain.optPhaseIn[0].checked) && (!frmMain.optPhaseIn[1].checked) && (!frmMain.optPhaseIn[2].checked)) && blnSuccess) {
                    window.alert("You must select BIOS Affectivity.");
                    FocusOn = frmMain.txtRTMName;
                    blnSuccess = false;
                }
                else if ((frmMain.chkFW.checked && (!frmMain.optPhaseIn[0].checked) && (!frmMain.optPhaseIn[1].checked) && (!frmMain.optPhaseIn[2].checked)) && blnSuccess) {
                    window.alert("You must select FW Affectivity.");
                    FocusOn = frmMain.txtRTMName;
                    blnSuccess = false;
                }
                break;
            case "BIOS":

                if (typeof (frmMain.chkBIOSList.length) == "undefined")
                    BIOSArray[0] = frmMain.chkBIOSList;
                else
                    BIOSArray = frmMain.chkBIOSList;
                blnFound = false;
                for (i = 0; i < BIOSArray.length; i++) {
                    if (BIOSArray[i].checked) {
                        blnFound = true;
                        break;
                    }
                }

                if (!blnFound) {
                    window.alert("You must select at least one BIOS version.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
            case "Restore":

                if (typeof (frmMain.chkRestoreList.length) == "undefined")
                    RestoreArray[0] = frmMain.chkRestoreList;
                else
                    RestoreArray = frmMain.chkRestoreList;
                blnFound = false;
                for (i = 0; i < RestoreArray.length; i++) {
                    if (RestoreArray[i].checked) {
                        blnFound = true;
                        break;
                    }
                }

                if (!blnFound) {
                    window.alert("You must select at least one Restore Media version.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
            case "Patches":

                if (typeof (frmMain.chkPatchList.length) == "undefined")
                    PatchArray[0] = frmMain.chkPatchList;
                else
                    PatchArray = frmMain.chkPatchList;
                blnFound = false;
                for (i = 0; i < PatchArray.length; i++) {
                    if (PatchArray[i].checked) {
                        blnFound = true;
                        break;
                    }
                }

                if (!blnFound) {
                    window.alert("You must select at least one Patch.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
            case "FW":

                if (typeof (frmMain.chkFWList.length) == "undefined")
                    FWArray[0] = frmMain.chkFWList;
                else
                    FWArray = frmMain.chkFWList;
                blnFound = false;
                for (i = 0; i < FWArray.length; i++) {
                    if (FWArray[i].checked) {
                        blnFound = true;
                        break;
                    }
                }

                if (!blnFound) {
                    window.alert("You must select at least one FW version.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
            case "Alerts":
                if (!frmMain.chkBuildLevel.checked) {
                    window.alert("You must signoff on the Build Level Alerts.");
                    FocusOn = frmMain.chkBuildLevel;
                    blnSuccess = false;
                }
                else if (!frmMain.chkDistribution.checked) {
                    window.alert("You must signoff on the Distribution Alerts.");
                    FocusOn = frmMain.chkDistribution;
                    blnSuccess = false;
                }
                else if (!frmMain.chkCertification.checked) {
                    window.alert("You must signoff on the Certification Alerts.");
                    FocusOn = frmMain.chkCertification;
                    blnSuccess = false;
                }
                else if (!frmMain.chkWorkflow.checked) {
                    window.alert("You must signoff on the Workflow Alerts.");
                    FocusOn = frmMain.chkWorkflow;
                    blnSuccess = false;
                }
                else if (!frmMain.chkAvailability.checked) {
                    window.alert("You must signoff on the Availability Alerts.");
                    FocusOn = frmMain.chkAvailability;
                    blnSuccess = false;
                }
                else if (!frmMain.chkDeveloper.checked) {
                    window.alert("You must signoff on the Developer Alerts.");
                    FocusOn = frmMain.chkDeveloper;
                    blnSuccess = false;
                }
                else if (!frmMain.chkRoot.checked) {
                    window.alert("You must signoff on the Root Deliverable Alerts.");
                    FocusOn = frmMain.chkRoot;
                    blnSuccess = false;
                }
                else if (!frmMain.chkOTSPrimary.checked) {
                    window.alert("You must signoff on the Primary OTS Alerts.");
                    FocusOn = frmMain.chkOTSPrimary;
                    blnSuccess = false;
                }

                break;
            case "Images":
                blnFound = false;
                for (i = 0; i < document.all.length; i++) {
                    if (document.all(i).className == "chkBase" || document.all(i).className == "chkDrop") {
                        if (document.all(i).checked || document.all(i).indeterminate) {
                            blnFound = true;
                            break;
                        }
                    }
                }

                if (!blnFound) {
                    window.alert("You must select at least one Image.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
            case "SCMX":
                blnFound = false;

                //validate file uploaded
                if (document.getElementById("UploadRemoveLinks1").style.display != "none") {
                    blnFound = true;
                }

                if (!blnFound) {
                    window.alert("You must upload a file.");
                    FocusOn = window.document;
                    blnSuccess = false;
                }


                break;
        }


        if(typeof (frmMain.txtRTMComments) != "undefined"){
            if ((frmMain.txtRTMComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtRTMComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtRTMComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtBIOSComments) != "undefined"){
            if ((frmMain.txtBIOSComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtBIOSComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtBIOSComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtBuildLevelComments) != "undefined"){
            if ((frmMain.txtBuildLevelComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtBuildLevelComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtBuildLevelComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtDistributionComments) != "undefined"){
            if ((frmMain.txtDistributionComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtDistributionComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtDistributionComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtCertificationComments) != "undefined"){
            if ((frmMain.txtCertificationComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtCertificationComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtCertificationComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtWorkflowComments) != "undefined"){
            if ((frmMain.txtWorkflowComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtWorkflowComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtWorkflowComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtAvailabilityComments) != "undefined"){
            if ((frmMain.txtAvailabilityComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtAvailabilityComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtAvailabilityComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtDeveloperComments) != "undefined"){
            if ((frmMain.txtDeveloperComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtDeveloperComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtDeveloperComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtRootComments) != "undefined"){
            if ((frmMain.txtRootComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtRootComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtRootComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtOTSPrimaryComments) != "undefined"){
            if ((frmMain.txtOTSPrimaryComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtOTSPrimaryComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtOTSPrimaryComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtNotify) != "undefined"){
            if ((frmMain.txtNotify.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtNotify.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtNotify;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtImageComments) != "undefined"){
            if ((frmMain.txtImageComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtImageComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtImageComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtRestoreComments) != "undefined"){
            if ((frmMain.txtRestoreComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtRestoreComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtRestoreComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtFWComments) != "undefined"){
            if ((frmMain.txtFWComments.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtFWComments.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtFWComments;
                blnSuccess = false;
            }
        }
        if(typeof (frmMain.txtEmailPreview) != "undefined"){
            if ((frmMain.txtEmailPreview.value.match(/[^\x00-\x7F\r\n]/)) && blnSuccess) {
                $("#dialog-charError").dialog("open");
                $("#CharErrorMsg").text(frmMain.txtEmailPreview.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                FocusOn = frmMain.txtEmailPreview;
                blnSuccess = false;
            }
        }


        if (blnSuccess == false) {
            if (CurrentState != strState) {
                CurrentState = strState;
                ProcessState();
            }

            FocusOn.focus();
        }


        return blnSuccess;
    }

    function window_onload() {
        $("#dialog-charError").dialog({
            height: 600,
            width: 600,
            modal: true,
            autoOpen: false
        });

        var i;
        var strID;
        var strName;
        if (typeof (frmMain) != "undefined") {

            if (frmMain.txtProductRTMID.value != "0") {
                
                if (frmMain.chkBuildLevel.checked) {

                    if (frmMain.chkBIOS.checked && !frmMain.chkRestore.checked && !frmMain.chkImages.checked)
                        document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value + "&ReportType=2";
                    else
                        document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                    BuildLevelAlertDetails.style.display = "none";
                    frmMain.txtBuildLevelComments.focus();
                }
                else {
                    BuildLevelAlertDetails.style.display = "";
                    document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                }
                if (frmMain.chkDistribution.checked) {
                    document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;
                    DistributionAlertDetails.style.display = "none";
                    frmMain.txtDistributionComments.focus();
                }
                else {
                    document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;
                    DistributionAlertDetails.style.display = "";
                }
                if (frmMain.chkCertification.checked) {
                    document.all.CertificationIFrame.src = frmMain.txtCertificationIFramesrc.value;
                    CertificationAlertDetails.style.display = "none";
                    frmMain.txtCertificationComments.focus();
                }
                else {
                    document.all.CertificationIFrame.src = frmMain.txtCertificationIFramesrc.value;
                    CertificationAlertDetails.style.display = "";
                }
                if (frmMain.chkWorkflow.checked) {
                    document.all.WorkflowIFrame.src = frmMain.txtWorkflowIFramesrc.value;
                    WorkflowAlertDetails.style.display = "none";
                    frmMain.txtWorkflowComments.focus();
                }
                else {
                    document.all.WorkflowIFrame.src = frmMain.txtWorkflowIFramesrc.value;
                    WorkflowAlertDetails.style.display = "";
                }
                if (frmMain.chkAvailability.checked) {
                    document.all.AvailabilityIFrame.src = frmMain.txtAvailabilityIFramesrc.value;
                    AvailabilityAlertDetails.style.display = "none";
                    frmMain.txtAvailabilityComments.focus();
                }
                else {
                    document.all.AvailabilityIFrame.src = frmMain.txtAvailabilityIFramesrc.value;
                    AvailabilityAlertDetails.style.display = "";
                }
                if (frmMain.chkDeveloper.checked) {
                    document.all.DeveloperIFrame.src = frmMain.txtDeveloperIFramesrc.value;
                    DeveloperAlertDetails.style.display = "none";
                    frmMain.txtDeveloperComments.focus();
                }
                else {
                    document.all.DeveloperIFrame.src = frmMain.txtDeveloperIFramesrc.value;
                    DeveloperAlertDetails.style.display = "";
                }
                if (frmMain.chkRoot.checked) {
                    document.all.RootIFrame.src = frmMain.txtRootIFramesrc.value;
                    RootAlertDetails.style.display = "none";
                    frmMain.txtRootComments.focus();
                }
                else {
                    document.all.RootIFrame.src = frmMain.txtRootIFramesrc.value;
                    RootAlertDetails.style.display = "none";
                }
                if (frmMain.chkOTSPrimary.checked) {
                    document.all.OTSPrimaryIFrame.src = frmMain.txtOTSPrimaryIFramesrc.value;
                    OTSPrimaryAlertDetails.style.display = "none";
                    frmMain.txtOTSPrimaryComments.focus();
                }
                else {
                    document.all.OTSPrimaryIFrame.src = frmMain.txtOTSPrimaryIFramesrc.value;
                    OTSPrimaryAlertDetails.style.display = "none";
                }
                var strAttachmentPath = frmMain.txtAttachmentPath1.value;
                if (strAttachmentPath != "" && strAttachmentPath != "undefined") {
                    var strServer;
                    var PathArray;
                    var ID = 1;
                    PathArray = strAttachmentPath.split("|");
                    $("#UploadAddLinks" + ID).hide();
                    $("#UploadRemoveLinks" + ID).show();
                    $("#txtUploadPath" + ID).val(PathArray[0].substr(0, PathArray[0].lastIndexOf("\\")));
                    $("#txtAttachmentPath" + ID).val(PathArray[0]);
                    $("#UploadPath" + ID).text(PathArray[0].substr(PathArray[0].lastIndexOf("\\") + 1, PathArray[0].length));
                }
            }

                if (txtImageCount.value == "0") {
                    frmMain.chkImages.checked = false;
                    frmMain.chkImages.disabled = true;
                    ImagesDisabled.innerHTML = "&nbsp;";//"&nbsp;(None&nbsp;Available)"
                    ImagesTextColor.color = "darkgray";
                }
                else {
                    frmMain.chkImages.disabled = false;
                    ImagesDisabled.innerHTML = "";
                    ImagesTextColor.color = "black";
                }

                if (txtBIOSCount.value == "0") {
                    frmMain.chkBIOS.checked = false;
                    frmMain.chkBIOS.disabled = true;
                    BIOSDisabled.innerHTML = "&nbsp;";//"&nbsp;(None&nbsp;Available)"
                    BIOSTextColor.color = "darkgray";
                }
                else {
                    frmMain.chkBIOS.disabled = false;
                    BIOSDisabled.innerHTML = "";
                    BIOSTextColor.color = "black";
                }

                if (txtRestoreCount.value == "0") {
                    frmMain.chkRestore.checked = false;
                    frmMain.chkRestore.disabled = true;
                    RestoreDisabled.innerHTML = "&nbsp;";//"&nbsp;(None&nbsp;Available)"
                    RestoreTextColor.color = "darkgray";
                }
                else {
                    frmMain.chkRestore.disabled = false;
                    RestoreDisabled.innerHTML = "";
                    RestoreTextColor.color = "black";
                }


                if (txtPatchCount.value == "0") {
                    frmMain.chkPatch.checked = false;
                    frmMain.chkPatch.disabled = true;
                    PatchDisabled.innerHTML = "&nbsp;";//"&nbsp;(None&nbsp;Available)"
                    PatchTextColor.color = "darkgray";
                }
                else {
                    frmMain.chkPatch.disabled = false;
                    PatchDisabled.innerHTML = "";
                    PatchTextColor.color = "black";
                }

                if (txtFWCount.value == "0") {
                    frmMain.chkFW.checked = false;
                    frmMain.chkFW.disabled = true;
                    FWDisabled.innerHTML = "&nbsp;";
                    FWTextColor.color = "darkgray";
                }
                else {
                    frmMain.chkFW.disabled = false;
                    FWDisabled.innerHTML = "";
                    FWTextColor.color = "black";
                }

            if (frmMain.txtProductRTMID.value != "0") {
                if (frmMain.chkSCMX.checked) {
                    BIOSAffectivityRow.style.display = "none";

                    frmMain.chkBIOS.checked = false;
                    frmMain.chkBIOS.disabled = true;

                    BIOSTextColor.color = "darkgray";

                    frmMain.chkFW.checked = false;
                    frmMain.chkFW.disabled = true;

                    FWTextColor.color = "darkgray";

                    frmMain.chkRestore.checked = false;
                    frmMain.chkRestore.disabled = true;

                    RestoreTextColor.color = "darkgray";

                    frmMain.chkImages.checked = false;
                    frmMain.chkImages.disabled = true;

                    ImagesTextColor.color = "darkgray";

                    frmMain.chkPatch.checked = false;
                    frmMain.chkPatch.disabled = true;

                    PatchTextColor.color = "darkgray";

                }

                if (frmMain.chkFW.checked) {
                    showAffectivityRow(frmMain.chkFW.checked);
                    frmMain.chkBIOS.checked = false;
                    frmMain.chkBIOS.disabled = true;
                    BIOSTextColor.color = "darkgray";

                    frmMain.chkSCMX.checked = false;
                    frmMain.chkSCMX.disabled = true;
                    SCMXTextColor.color = "darkgray";

                    frmMain.chkRestore.checked = false;
                    frmMain.chkRestore.disabled = true;
                    RestoreTextColor.color = "darkgray";

                    frmMain.chkPatch.checked = false;
                    frmMain.chkPatch.disabled = true;
                    PatchTextColor.color = "darkgray";

                }
                if (frmMain.chkBIOS.checked || frmMain.chkRestore.checked || frmMain.chkImages.checked || frmMain.chkPatch.checked) {

                    frmMain.chkSCMX.checked = false;
                    frmMain.chkSCMX.disabled = true;
                    SCMXTextColor.color = "darkgray";

                }
                if ( frmMain.chkRestore.checked ||  frmMain.chkPatch.checked) {

                    showAffectivityRow(false);
                    frmMain.chkFW.checked = false;
                    frmMain.chkFW.disabled = true;
                    FWTextColor.color = "darkgray";
                }

                if (frmMain.chkBIOS.checked) {
                    frmMain.chkFW.checked = false;
                    frmMain.chkFW.disabled = true;
                    showAffectivityRow(true);
                }
            }

               
            CurrentState = "General";
            ProcessState();
            FormLoading = false;
        }
        else
            window.parent.frames["LowerWindow"].cmdNext.disabled = true;

        //Add modal dialog code to body tag: ---
        modalDialog.load();

        //load date picker: ---
        load_datePicker();        
        if (frmMain.txtPartnerId.value!="1") {
            window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
        }        
    }

    function LoadAlerts(strType) {
        if (!frmMain.chkBuildLevel.checked)
            document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value + strType;
        if (!frmMain.chkDistribution.checked)
            document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value + strType;
        if (!frmMain.chkCertification.checked)
            document.all.CertificationIFrame.src = frmMain.txtCertificationIFramesrc.value + strType;
        if (!frmMain.chkWorkflow.checked)
            document.all.WorkflowIFrame.src = frmMain.txtWorkflowIFramesrc.value + strType;
        if (!frmMain.chkAvailability.checked)
            document.all.AvailabilityIFrame.src = frmMain.txtAvailabilityIFramesrc.value + strType;
        if (!frmMain.chkDeveloper.checked)
            document.all.DeveloperIFrame.src = frmMain.txtDeveloperIFramesrc.value + strType;
        if (!frmMain.chkRoot.checked)
            document.all.RootIFrame.src = frmMain.txtRootIFramesrc.value + strType;
        if (!frmMain.chkOTSPrimary.checked)
            document.all.OTSPrimaryIFrame.src = frmMain.txtOTSPrimaryIFramesrc.value + strType;

        if (strType == "&ReportType=2") {
            OTSPrimaryTypeText.innerText = "System BIOS deliverables"
            RootTypeText.innerText = "(System BIOS deliverables)"
            DeveloperTypeText.innerText = "(System BIOS deliverables)"
            AvailabilityTypeText.innerText = "(System BIOS deliverables)"
            WorkflowTypeText.innerText = "(System BIOS deliverables)"
            CertificationTypeText.innerText = "(System BIOS deliverables)"
            DistributionTypeText.innerText = "(System BIOS deliverables)"
            BuildLevelTypeText.innerText = "(System BIOS deliverables)"
        }
        else {
            //  OTSRelatedTypeText.innerText = "SW, FW, and Doc deliverables"
            OTSPrimaryTypeText.innerText = "SW, FW, and Doc deliverables"
            RootTypeText.innerText = "(SW, FW, and Doc deliverables)"
            DeveloperTypeText.innerText = "(SW, FW, and Doc deliverables)"
            AvailabilityTypeText.innerText = "(SW, FW, and Doc deliverables)"
            WorkflowTypeText.innerText = "(SW, FW, and Doc deliverables)"
            CertificationTypeText.innerText = "(SW, FW, and Doc deliverables)"
            DistributionTypeText.innerText = "(SW, FW, and Doc deliverables)"
            BuildLevelTypeText.innerText = "(SW, FW, and Doc deliverables)"
        }
    }

    var CurrentState;
    var FormLoading = true;


    function ProcessState() {
        var steptext;
        var strPreview;
        var strApprovers;

        switch (CurrentState) {
            case "General":
                lblTitle.innerText = "Enter General RTM information";
                if (frmMain.txtRTMComments.value == "") {
                    frmMain.txtRTMComments.value = frmMain.txtRTMCommentsTemplate.value;
                    frmMain.txtRTMComments.style.fontStyle = "italic";
                    frmMain.txtRTMComments.style.color = "blue";
                }
                tabGeneral.style.display = "";
                tabPatch.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabRestore.style.display = "none";
                tabImages.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = true;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.txtRTMName.focus();
                window.scrollTo(0, 0);
                break;

            case "BIOS":
                lblTitle.innerText = "Select System BIOS to RTM";

                if (frmMain.txtBIOSComments.value == "") {
                    frmMain.txtBIOSComments.value = frmMain.txtBIOSCommentsTemplate.value;
                    frmMain.txtBIOSComments.style.fontStyle = "italic";
                    frmMain.txtBIOSComments.style.color = "blue";
                }

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "";
                tabPatch.style.display = "none";
                tabRestore.style.display = "none";
                tabImages.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtBIOSComments.focus();
                window.scrollTo(0, 0);
                break;

            case "FW":
                lblTitle.innerText = "Select Firmware to RTM";

                if (frmMain.txtFWComments.value == "") {
                    frmMain.txtFWComments.value = frmMain.txtFWCommentsTemplate.value;
                    frmMain.txtFWComments.style.fontStyle = "italic";
                    frmMain.txtFWComments.style.color = "blue";
                }

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "";
                tabPatch.style.display = "none";
                tabRestore.style.display = "none";
                tabImages.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                window.scrollTo(0, 0);
                break;

            case "Patches":
                lblTitle.innerText = "Select Patches to RTM";

                if (frmMain.txtPatchComments.value == "") {
                    frmMain.txtPatchComments.value = frmMain.txtPatchCommentsTemplate.value;
                    frmMain.txtPatchComments.style.fontStyle = "italic";
                    frmMain.txtPatchComments.style.color = "blue";
                }

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabPatch.style.display = "";
                tabRestore.style.display = "none";
                tabImages.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                window.scrollTo(0, 0);
                break;

            case "Restore":

                //document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                //document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;

                lblTitle.innerText = "Select Restore Media to RTM";

                if (frmMain.txtRestoreComments.value == "") {
                    frmMain.txtRestoreComments.value = frmMain.txtRestoreCommentsTemplate.value;
                    frmMain.txtRestoreComments.style.fontStyle = "italic";
                    frmMain.txtRestoreComments.style.color = "blue";
                }

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabPatch.style.display = "none";
                tabRestore.style.display = "";
                tabImages.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtRestoreComments.focus();
                window.scrollTo(0, 0);
                break;

            case "Images":

                //document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                //document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;

                lblTitle.innerText = "Select Images to RTM";

                if (frmMain.txtImageComments.value == "") {
                    frmMain.txtImageComments.value = frmMain.txtImageCommentsTemplate.value;
                    frmMain.txtImageComments.style.fontStyle = "italic";
                    frmMain.txtImageComments.style.color = "blue";
                }

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabRestore.style.display = "none";
                tabPatch.style.display = "none";
                tabImages.style.display = "";
                trImageComments.style.display = "";
                divImageTable.style.display = "";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtImageComments.focus();
                window.scrollTo(0, 0);
                break;

            case "SCMX":

                lblTitle.innerText = "Upload SCMX file to RTM";

                if (frmMain.txtImageComments.value == "") {
                    frmMain.txtImageComments.value = frmMain.txtImageCommentsTemplate.value;
                    frmMain.txtImageComments.style.fontStyle = "italic";
                    frmMain.txtImageComments.style.color = "blue";
                }
                trImageComments.style.display = "";
                divImageTable.style.display = "none";

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabRestore.style.display = "none";
                tabPatch.style.display = "none";
                tabImages.style.display = "";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "none";

                if (typeof (noImage) != "undefined")
                    noImage.style.display = "none";



                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtImageComments.focus();
                window.scrollTo(0, 0);
                break;

            case "Alerts":
                //Temp
                /*
                frmMain.chkAvailability.checked=true;
                frmMain.chkBuildLevel.checked=true;
                frmMain.chkCertification.checked=true;
                frmMain.chkDeveloper.checked=true;
                frmMain.chkDistribution.checked=true;
                frmMain.chkOTSPrimary.checked=true;
                frmMain.chkRoot.checked=true;
                frmMain.chkWorkflow.checked=true;
                */
                //temp
                lblTitle.innerText = "Review Alerts";

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabRestore.style.display = "none";
                tabPatch.style.display = "none";
                tabImages.style.display = "none";
                tabAlerts.style.display = "";
                tabPreview.style.display = "none";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    window.parent.frames["LowerWindow"].cmdSave.disabled = true;
                }
                window.document.focus();
                window.scrollTo(0, 0);
                break;

            case "Preview":
                lblTitle.innerText = "Review Selected Information";
                PopulatePreview();

                tabGeneral.style.display = "none";
                tabBIOS.style.display = "none";
                tabFW.style.display = "none";
                tabRestore.style.display = "none";
                tabImages.style.display = "none";
                tabPatch.style.display = "none";
                tabAlerts.style.display = "none";
                tabPreview.style.display = "";
                if (!FormLoading) {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = true;
                    if (frmMain.txtPartnerId.value!="1") {
                            window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                    }
                    else {
                          window.parent.frames["LowerWindow"].cmdFinish.disabled = false;
                    }
                    window.parent.frames["LowerWindow"].cmdSave.disabled = false;
                }

                //frmMain.txtPreview.focus();
                window.document.focus();
                window.scrollTo(0, 0);
                break;
        }
    }

    function PopulatePreview() {
        var strPreview = "";
        var i;
        var strVersions = "";
        var ImageValueArray;
        var ControlArray = new Array();
        var RestoreArray = new Array();
        var PatchArray = new Array();
        var ImageArray = new Array();
        var isFusion = false;

        if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value)
            frmMain.txtRTMComments.value = "";

        if (typeof (frmMain.txtBIOSComments) != "undefined")
            if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value)
                frmMain.txtBIOSComments.value = "";

        if (typeof (frmMain.txtFWComments) != "undefined")
            if (frmMain.txtFWComments.value == frmMain.txtFWCommentsTemplate.value)
                frmMain.txtFWComments.value = "";

        if (typeof (frmMain.txtRestoreComments) != "undefined")
            if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value)
                frmMain.txtRestoreComments.value = "";

        if (typeof (frmMain.txtImageComments) != "undefined")
            if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value)
                frmMain.txtImageComments.value = "";

        strPreview = "<font size=2 face=verdana><b>General RTM Information</b></font><table class=EmbeddedTable bgcolor=white width=100% border=1 cellpadding=2 cellspacing=0>"
        strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>RTM Title:</b></td><td width='100%'>" + frmMain.txtRTMName.value + "&nbsp;</td>";
        strPreview = strPreview + "<td nowrap valign=top bgcolor=gainsboro><b>RTM Date:</b></td><td width='120'>" + frmMain.txtRTMDate.value + "</td></tr>";

        if ((frmMain.txtAttachmentPath1.value != "") && (!frmMain.chkImages.checked) && (!frmMain.chkSCMX.checked)) {
            RemoveUpload(1);
        }
        if (!UploadPath1.innerText) {
            if (document.getElementById("chkSCMX").checked) {
                strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>SCMx File:</b></td><td colspan=3 width='100%'>" + "None" + "&nbsp;</td></tr>";
            }
        } else if (UploadPath1.innerText != "") {
            strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>SCMx File:</b></td><td colspan=3 width='100%'>" + "<a target='_blank' href='" + frmMain.txtUploadPath1.value + "'>" + UploadPath1.innerText + "</a>" + "&nbsp;</td></tr>";
        }

        if (frmMain.txtRTMComments.value)
            strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>RTM Comments:</b></td><td colspan=3>" + frmMain.txtRTMComments.value.replace(/\r\n/g, "<BR>") + "</td></tr>";
        strPreview = strPreview + "</table>"
        if (frmMain.chkBIOS.checked) {
            strVersions = "";

            if (typeof (frmMain.chkBIOSList.length) == "undefined")
                ControlArray[0] = frmMain.chkBIOSList;
            else
                ControlArray = frmMain.chkBIOSList;

            for (i = 0; i < ControlArray.length; i++)
                if (ControlArray[i].checked) {
                    strVersions = strVersions + "<tr><td>" + ControlArray[i].PreviewID + "</td><td>" + ControlArray[i].PreviewName + "</td><td>" + ControlArray[i].PreviewVersion + "</td>"
                    if (frmMain.optCutIn.checked)
                        strVersions = strVersions + "<td>Immediate (Rework All Units)</td></tr>";
                    else if (frmMain.optWebOnly.checked)
                        strVersions = strVersions + "<td>Web Release Only</td></tr>";
                    else
                        strVersions = strVersions + "<td>Phase-in</td></tr>";
                }
            strPreview = strPreview + "<BR><font size=2 face=verdana><b>System BIOS to RTM</b></font><BR>";
            if (frmMain.txtBIOSComments.value) {
                strPreview = strPreview + "<font size=1 face=verdana color=black><BR><i>" + frmMain.txtBIOSComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
                // strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1>"
                // strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>Comments:</b></td><td width='100%'>" + frmMain.txtBIOSComments.value.replace(/\r\n/g,"<BR>") + "</td></tr></table><BR>";
                //   strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1>"
                //   strPreview = strPreview + "<tr><td width='100%'>" + frmMain.txtBIOSComments.value.replace(/\r\n/g,"<BR>") + "</td></tr></table><BR>";
            }
            strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Affectivity</b></td></tr>"
            strPreview = strPreview + strVersions + "</table>";
        }

        if (frmMain.chkFW.checked) {
            strVersions = "";

            if (typeof (frmMain.chkFWList.length) == "undefined")
                ControlArray[0] = frmMain.chkFWList;
            else
                ControlArray = frmMain.chkFWList;

            for (i = 0; i < ControlArray.length; i++)
                if (ControlArray[i].checked) {
                    strVersions = strVersions + "<tr><td>" + ControlArray[i].PreviewID + "</td><td>" + ControlArray[i].PreviewName + "</td><td>" + ControlArray[i].PreviewVersion + "</td>"
                    if (frmMain.optCutIn.checked)
                        strVersions = strVersions + "<td>Immediate (Rework All Units)</td></tr>";
                    else if (frmMain.optWebOnly.checked)
                        strVersions = strVersions + "<td>Web Release Only</td></tr>";
                    else
                        strVersions = strVersions + "<td>Phase-in</td></tr>";
                }
            strPreview = strPreview + "<BR><font size=2 face=verdana><b>Firmware to RTM</b></font><BR>";
            if (frmMain.txtFWComments.value) {
                strPreview = strPreview + "<font size=1 face=verdana color=black><BR><i>" + frmMain.txtFWComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
            }
            strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Affectivity</b></td></tr>"
            strPreview = strPreview + strVersions + "</table>";
        }

        if (frmMain.chkRestore.checked) {
            strVersions = "";
            if (typeof (frmMain.chkRestoreList.length) == "undefined")
                RestoreArray[0] = frmMain.chkRestoreList;
            else
                RestoreArray = frmMain.chkRestoreList;

            for (i = 0; i < RestoreArray.length; i++)
                if (RestoreArray[i].checked) {
                    strVersions = strVersions + "<tr><td valign=top>" + RestoreArray[i].PreviewID + "</td><td valign=top>" + RestoreArray[i].PreviewName + "</td><td valign=top>" + RestoreArray[i].PreviewVersion + "</td><td valign=top>" + RestoreArray[i].PreviewPart + "&nbsp;</td><td valign=top><a target=_blank href=\"http://houcmitrel02.auth.hpicorp.net:81/cdpmr/CDQuery.aspx?ExcalID=" + RestoreArray[i].PreviewID + "\">" + RestoreArray[i].PreviewPMR + "</a>&nbsp;</td></tr>"
                }
            strPreview = strPreview + "<BR><font size=2 face=verdana><b>Restore Media to RTM</b></font><BR>";
            if (frmMain.txtRestoreComments.value) {
                strPreview = strPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtRestoreComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
            }
            strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Part</b></td><td><b>PMR&nbsp;Date</b></td></tr>"
            strPreview = strPreview + strVersions + "</table>";
        }

        if (frmMain.chkPatch.checked) {
            strVersions = "";
            if (typeof (frmMain.chkPatchList.length) == "undefined")
                PatchArray[0] = frmMain.chkPatchList;
            else
                PatchArray = frmMain.chkPatchList;

            for (i = 0; i < PatchArray.length; i++)
                if (PatchArray[i].checked) {
                    strVersions = strVersions + "<tr><td valign=top>" + PatchArray[i].PreviewID + "</td><td valign=top>" + PatchArray[i].PreviewName + "</td><td valign=top>" + PatchArray[i].PreviewVersion + "</td><td valign=top>" + PatchArray[i].PreviewContents + "&nbsp;</td><td valign=top><a target=_blank href='../Image/PatchImages.asp?ProdID=" + frmMain.txtProductID.value + "&DelID=" + PatchArray[i].PreviewID + "'>View</a>&nbsp;</td></tr>"
                }
            strPreview = strPreview + "<BR><font size=2 face=verdana><b>Patches to RTM</b></font><BR>";
            if (frmMain.txtPatchComments.value) {
                strPreview = strPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtPatchComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
            }
            strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Patch&nbsp;Contents</b></td><td><b>Images</b></td></tr>"
            strPreview = strPreview + strVersions + "</table>";
        }

        if (frmMain.chkSCMX.checked) {

            strPreview = strPreview + "<BR><font size=2 face=verdana><b>SCMX to RTM</b></font><BR>";
            if (frmMain.txtImageComments.value) {
                strPreview = strPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtImageComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
            }

        }

        if (frmMain.chkImages.checked) {
            strVersions = "";
            if (typeof (frmMain.txtImagePreview.length) == "undefined")
                ImageArray[0] = frmMain.txtImagePreview;
            else
                ImageArray = frmMain.txtImagePreview;
            for (i = 0; i < ImageArray.length; i++) {
                ImageValueArray = ImageArray[i].value.split("|")
                if (ImageValueArray[0] == "FUSION") {
                    isFusion = true;
                    if (document.all("chkImage" + ImageValueArray[1]).checked) {
                        strVersions = strVersions + "<tr><td>" + ImageValueArray[2] + "&nbsp;</td><td>" + ImageValueArray[7] + "&nbsp;-&nbsp;" + ImageValueArray[6] + "&nbsp;</td><td>" + ImageValueArray[3] + "&nbsp;</td><td>" + ImageValueArray[4] + "&nbsp;</td><td>" + ImageValueArray[5] + "&nbsp;</td></tr>"
                    }
                }
                else {
                    if (document.all("chkImage" + ImageValueArray[0]).checked) {
                        strVersions = strVersions + "<tr><td>" + ImageValueArray[1] + "&nbsp;</td><td>" + ImageValueArray[8] + "&nbsp;-&nbsp;" + ImageValueArray[7] + "&nbsp;</td><td>" + ImageValueArray[2] + "&nbsp;</td><td>" + ImageValueArray[3] + "&nbsp;</td><td>" + ImageValueArray[4] + "&nbsp;</td><td>" + ImageValueArray[5] + "&nbsp;</td></tr>"
                    }
                }
            }
            strPreview = strPreview + "<BR><font size=2 face=verdana><b>Images to RTM</b></font><BR>";
            if (frmMain.txtImageComments.value) {
                strPreview = strPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtImageComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
            }
            if (isFusion)
                strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>Product&nbsp;Drop</b></td><td><b>Region</b></td><td><b>Brands</b></td><td><b>OS</b></td><td><b>Comments</b></td></tr>"
            else
                strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>SKU</b></td><td><b>Region</b></td><td><b>Model</b></td><td><b>OS</b></td><td><b>Apps Bundle</b></td><td><b>BTO/CTO</b></td></tr>"
            strPreview = strPreview + strVersions + "</table>";
        }

        //Alert
        strPreview = strPreview + "<div id=AlertPreviewSection><BR><font size=2 face=verdana><b>Alerts Reviewed</b></font><BR>";

        strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>Alert</b></td><td><b>Count</b></td><td width='100%'><b>Comments</b></td></tr>"
        try {
            strPreview = strPreview + "<tr><td nowrap>Build Level</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + BuildLevelIFrame.RecordID.innerText + "'>" + GetAlertCount(BuildLevelIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtBuildLevelComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Build Level</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Distribution</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + DistributionIFrame.RecordID.innerText + "'>" + GetAlertCount(DistributionIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtDistributionComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Distribution</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Certification</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + CertificationIFrame.RecordID.innerText + "'>" + GetAlertCount(CertificationIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtCertificationComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Certification</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Workflow</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + WorkflowIFrame.RecordID.innerText + "'>" + GetAlertCount(WorkflowIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtWorkflowComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Workflow</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Availability</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + AvailabilityIFrame.RecordID.innerText + "'>" + GetAlertCount(AvailabilityIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtAvailabilityComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Availability</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Developer</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + DeveloperIFrame.RecordID.innerText + "'>" + GetAlertCount(DeveloperIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtDeveloperComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Developer</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>Root</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + RootIFrame.RecordID.innerText + "'>" + GetAlertCount(RootIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtRootComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>Root</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        try {
            strPreview = strPreview + "<tr><td nowrap>OTS Primary</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + OTSPrimaryIFrame.RecordID.innerText + "'>" + GetAlertCount(OTSPrimaryIFrame.document.body.innerHTML) + "</a></td><td>" + frmMain.txtOTSPrimaryComments.value + "&nbsp;</td></tr>"
        } catch (err) {
            strPreview = strPreview + "<tr><td nowrap>OTS Primary</td><td nowrap align=center>skip</td><td>No data</td></tr>"
        }
        // strPreview = strPreview + "<tr><td nowrap>OTS Related</td><td nowrap align=center><a href=''>" + GetAlertCount(OTSRelatedIFrame.document.body.innerHTML) + "</td><td>" + frmMain.txtOTSRelatedComments.value + "&nbsp;</td></tr>"

        strPreview = strPreview + "</table></div>";

        PreviewPage.innerHTML = strPreview;

    }


    function GetAlertCount(strString) {
        if (strString.indexOf("<TD>None Found</TD></TR>") > -1 && strString.toLowerCase().split(/<tr>/g).length == 1)
            return 0;
        else
            return strString.toLowerCase().split(/<tr>/g).length - 1;
    }

    function chkSCMX_onclick() {
        if (typeof (frmMain) != "undefined") {

            if (frmMain.chkSCMX.checked) {

                frmMain.chkBIOS.checked = false;
                frmMain.chkBIOS.disabled = true;

                BIOSTextColor.color = "darkgray";

                frmMain.chkFW.checked = false;
                frmMain.chkFW.disabled = true;

                FWTextColor.color = "darkgray";

                frmMain.chkRestore.checked = false;
                frmMain.chkRestore.disabled = true;

                RestoreTextColor.color = "darkgray";

                frmMain.chkImages.checked = false;
                frmMain.chkImages.disabled = true;

                ImagesTextColor.color = "darkgray";

                frmMain.chkPatch.checked = false;
                frmMain.chkPatch.disabled = true;

                PatchTextColor.color = "darkgray";

            } else {

                if (txtBIOSCount.value != "0") {
                    frmMain.chkBIOS.disabled = false;

                    BIOSTextColor.color = "black"
                }

                if (txtFWCount.value != "0") {
                    frmMain.chkFW.disabled = false;

                    FWTextColor.color = "black"
                }

                if (txtRestoreCount.value != "0") {
                    frmMain.chkRestore.disabled = false;

                    RestoreTextColor.color = "black"
                }

                if (txtImageCount.value != "0") {
                    frmMain.chkImages.disabled = false;

                    ImagesTextColor.color = "black"
                }

                if (txtPatchCount.value != "0") {
                    frmMain.chkPatch.disabled = false;

                    PatchTextColor.color = "black"
                }
            }
            showAffectivityRow(false);
        }
    }



    function showAffectivityRow(blnEnable) {

        if (blnEnable) {
            BIOSAffectivityRow.style.display = "";
        } else {
            BIOSAffectivityRow.style.display = "none";
        }

    }

    function chkBIOS_onclick() {
        showAffectivityRow(frmMain.chkBIOS.checked);
    }

    function chkFW_onclick() {
        showAffectivityRow(frmMain.chkFW.checked);

        if (typeof (frmMain) != "undefined") {

            if (frmMain.chkFW.checked) {

                frmMain.chkBIOS.checked = false;
                frmMain.chkBIOS.disabled = true;
                BIOSTextColor.color = "darkgray";

                frmMain.chkSCMX.checked = false;
                frmMain.chkSCMX.disabled = true;
                SCMXTextColor.color = "darkgray";

                frmMain.chkRestore.checked = false;
                frmMain.chkRestore.disabled = true;
                RestoreTextColor.color = "darkgray";

                frmMain.chkPatch.checked = false;
                frmMain.chkPatch.disabled = true;
                PatchTextColor.color = "darkgray";

            } else {

                if (txtBIOSCount.value != "0") {
                    frmMain.chkBIOS.disabled = false;

                    BIOSTextColor.color = "black"
                }

                if (!frmMain.chkImages.checked) {
                    frmMain.chkSCMX.disabled = false;
                    SCMXTextColor.color = "black";
                }

                if (txtRestoreCount.value != "0") {
                    frmMain.chkRestore.disabled = false;

                    RestoreTextColor.color = "black"
                }

                if (txtPatchCount.value != "0") {
                    frmMain.chkPatch.disabled = false;

                    PatchTextColor.color = "black"
                }
            }

        }

    }

    function chkStandalone_onclick() {

        if (frmMain.chkBIOS.checked || frmMain.chkRestore.checked || frmMain.chkImages.checked || frmMain.chkPatch.checked) {

            frmMain.chkSCMX.checked = false;
            frmMain.chkSCMX.disabled = true;
            SCMXTextColor.color = "darkgray";

        } else {
            if (!frmMain.chkFW.checked) {
                frmMain.chkSCMX.disabled = false;
                SCMXTextColor.color = "black";
            }
        }


        if (frmMain.chkBIOS.checked || frmMain.chkRestore.checked || frmMain.chkPatch.checked) {

            frmMain.chkFW.checked = false;
            frmMain.chkFW.disabled = true;
            FWTextColor.color = "darkgray";

        } else {

            if (txtFWCount.value != "0") {
                frmMain.chkFW.disabled = false;
                FWTextColor.color = "black"
            }

        }



    }


    function chkBuildLevel_onclick() {
        if (frmMain.chkBuildLevel.checked) {
            BuildLevelAlertDetails.style.display = "none";
            frmMain.txtBuildLevelComments.focus();
            //frmMain.txtBuildLevelComments.scrollIntoView(true);
        }
        else
            BuildLevelAlertDetails.style.display = "";
    }

    function chkDistribution_onclick() {
        if (frmMain.chkDistribution.checked) {
            DistributionAlertDetails.style.display = "none";
            frmMain.txtDistributionComments.focus();
        }
        else
            DistributionAlertDetails.style.display = "";
    }

    function chkAvailability_onclick() {
        if (frmMain.chkAvailability.checked) {
            AvailabilityAlertDetails.style.display = "none";
            frmMain.txtAvailabilityComments.focus();
        }
        else
            AvailabilityAlertDetails.style.display = "";
    }

    function chkRoot_onclick() {
        if (frmMain.chkRoot.checked) {
            RootAlertDetails.style.display = "none";
            frmMain.txtRootComments.focus();
        }
        else
            RootAlertDetails.style.display = "";
    }

    function chkOTSPrimary_onclick() {
        if (frmMain.chkOTSPrimary.checked) {
            OTSPrimaryAlertDetails.style.display = "none";
            frmMain.txtOTSPrimaryComments.focus();
        }
        else
            OTSPrimaryAlertDetails.style.display = "";
    }

    /*function chkOTSRelated_onclick(){
        if (frmMain.chkOTSRelated.checked)
            OTSRelatedAlertDetails.style.display="none";
        else
            OTSRelatedAlertDetails.style.display="";
    }
    */
    function chkDeveloper_onclick() {
        if (frmMain.chkDeveloper.checked) {
            DeveloperAlertDetails.style.display = "none";
            frmMain.txtDeveloperComments.focus();
        }
        else
            DeveloperAlertDetails.style.display = "";
    }

    function chkWorkflow_onclick() {
        if (frmMain.chkWorkflow.checked) {
            WorkflowAlertDetails.style.display = "none";
            frmMain.txtWorkflowComments.focus();
        }
        else
            WorkflowAlertDetails.style.display = "";
    }

    function chkCertification_onclick() {
        if (frmMain.chkCertification.checked) {
            CertificationAlertDetails.style.display = "none";
            frmMain.txtCertificationComments.focus();
        }
        else
            CertificationAlertDetails.style.display = "";
    }

    function cmdDate_onclick(FieldID) {
        var strID;


        strID = window.showModalDialog("../mobilese/today/caldraw1.asp", frmMain.txtRTMDate.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) != "undefined")
            frmMain.txtRTMDate.value = strID;
    }

    function BaseRow_onmouseover(strID) {
        if (window.event.srcElement.id == strID) {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("BaseRow" + strID).bgColor = "lightsteelblue";
    }

    function DropRow_onmouseover(strID) {
        if (window.event.srcElement.id == strID) {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("DropRow" + strID).bgColor = "lightsteelblue";
    }

    function BaseRow_onmouseout(strID) {
        document.all("BaseRow" + strID).bgColor = "cornsilk";
    }

    function DropRow_onmouseout(strID) {
        document.all("DropRow" + strID).bgColor = "cornsilk";
    }

    function BaseRow_onclick(strID) {

        if (window.event.srcElement.id == strID)
            return;


        if (document.all("ImageRow" + strID).style.display == "")
            document.all("ImageRow" + strID).style.display = "none";
        else
            document.all("ImageRow" + strID).style.display = "";
    }

    function DropRow_onclick(strID) {

        if (window.event.srcElement.id == strID)
            return;


        if (document.all("DropContents" + strID).style.display == "")
            document.all("DropContents" + strID).style.display = "none";
        else
            document.all("DropContents" + strID).style.display = "";
    }


    function chkAll_onclick() {
        var i;

        if (typeof (frmMain.chkImage.length) == "undefined") {
            if (frmMain.chkImage.indeterminate) // && frmMain.chkAll.checked
            {
                frmMain.chkImage.indeterminate = 0;
                document.all("Row" + frmMain.chkImage.value).bgColor = "ivory";
            }
            frmMain.chkImage.checked = frmMain.chkAll.checked;
        }
        else {
            for (i = 0; i < frmMain.chkImage.length; i++) {
                if (frmMain.chkImage(i).indeterminate) //&& frmmain.chkAll.checked
                {
                    frmMain.chkImage(i).indeterminate = 0;
                    document.all("Row" + frmMain.chkImage(i).value).bgColor = "ivory";
                }
                frmMain.chkImage(i).checked = frmMain.chkAll.checked;
                if (document.all("Base" + frmMain.chkImage(i).className).indeterminate) //&& frmMain.chkAll.checked
                    document.all("Base" + frmMain.chkImage(i).className).indeterminate = 0;
                document.all("Base" + frmMain.chkImage(i).className).checked = frmMain.chkAll.checked;
            }
        }

    }

    function chkBase_onclick() {
        var i;

        if (typeof (frmMain.chkImage.length) == "undefined") {
            if (frmMain.chkImage.indeterminate && window.event.srcElement.checked) {
                frmMain.chkImage.indeterminate = 0;
                document.all("Row" + frmMain.chkImage.value).bgColor = "ivory";
            }
            frmMain.chkImage.checked = window.event.srcElement.checked;
        }
        else {
            for (i = 0; i < frmMain.chkImage.length; i++) {
                if (frmMain.chkImage(i).className == window.event.srcElement.id) {
                    if (frmMain.chkImage(i).indeterminate && window.event.srcElement.checked) {
                        frmMain.chkImage(i).indeterminate = 0;
                        document.all("Row" + frmMain.chkImage(i).value).bgColor = "ivory";
                    }
                    frmMain.chkImage(i).checked = window.event.srcElement.checked;
                }
            }
        }

    }

    function chkDrop_onclick() {
        var i;

        if (typeof (frmMain.chkImage.length) == "undefined") {
            frmMain.chkImage.checked = window.event.srcElement.checked;
        }
        else {
            for (i = 0; i < frmMain.chkImage.length; i++) {
                if (frmMain.chkImage(i).className == window.event.srcElement.id) {
                    frmMain.chkImage(i).checked = window.event.srcElement.checked;
                }
            }
        }

    }

    function UpdateBase(chkClicked) {
        var i;
        var blnAllSame = true;

        for (i = 0; i < frmMain.chkImage.length; i++) {

            if (frmMain.chkImage(i).className != "")
                if (frmMain.chkImage(i).className == chkClicked.className) {
                    if ((frmMain.chkImage(i).checked != chkClicked.checked) || frmMain.chkImage(i).indeterminate) {
                        blnAllSame = false;
                    }
                }
        }

        if (blnAllSame) {
            document.all("Base" + chkClicked.className).indeterminate = 0;
            document.all("Base" + chkClicked.className).checked = chkClicked.checked;
        }
        else
            document.all("Base" + chkClicked.className).indeterminate = -1;

        if (chkClicked.checked) {
            if (document.all("Row" + chkClicked.value) != null)
                document.all("Row" + chkClicked.value).bgColor = "ivory";
        }

    }



    function chkImage_onclick() {
        UpdateBase(window.event.srcElement);
    }

    function SelectAllSameTierImages(intTier) {
        var boolIniCheckForTier;
        var chkImageNodes;
        chkImageNodes = document.getElementsByName("chkImage");

        if (typeof (chkImageNodes.length) != "undefined") {

            try {
                boolIniCheckForTier = document.getElementById(window.event.srcElement.getAttribute("data-chkimageid")).checked;
            } catch (e) {
                boolIniCheckForTier = false; // if error, always set the checkboxes ON.
            }

            for (i = 0; i < chkImageNodes.length; i++) {
                if (chkImageNodes[i].getAttribute("data-tier") == intTier) {
                    chkImageNodes[i].checked = !boolIniCheckForTier;
                    UpdateBase(chkImageNodes[i]);
                }
            };

        }

    }

//-->
    </script>
    <style type="text/css">
        A:visited {
            COLOR: blue
        }

        A:hover {
            COLOR: red
        }

        .EmbeddedTable TBODY TD {
            FONT-FAMILY: Verdana;
        }

        .EmbeddedTable TBODY TD {
            Font-Size: xx-small;
        }

        input {
            FONT-SIZE: 10pt;
            FONT-FAMILY: Verdana;
        }

        textarea {
            FONT-SIZE: 10pt;
            FONT-FAMILY: Verdana;
        }

        .ImageTable TBODY TD {
            BORDER-TOP: gray thin solid;
            FONT-SIZE: xx-small;
            FONT-FAMILY: verdana;
        }

        .ImageTable TH {
            FONT-SIZE: xx-small;
            FONT-FAMILY: verdana;
        }

        .imagerows TBODY TD {
            BORDER-TOP: none;
            FONT-SIZE: xx-small;
            FONT-FAMILY: verdana;
        }

        .imagerows THEAD TD {
            BORDER-TOP: none;
            FONT-SIZE: xx-small;
            FONT-FAMILY: verdana;
        }
    </style>
</head>
<body bgcolor="ivory" language="javascript" onload="return window_onload()">
    <link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
    <%
	dim cn
	dim rs
	dim blnFound 
	dim i
	dim cm
	dim p
	dim strCategories
	dim CurrentUser
	dim CurrentUserID
	dim CurrentWorkgroupID
	dim strSEPMID
	dim strPMID
	dim strTestLeadID
	dim blnPOR
	dim blnEditOK
	dim strShowEditBoxes
	dim strVersion
	dim strProductName
    dim blnFusion
	dim strEmployees
	dim strDevCenter
	dim strLastRoot
	dim BIOSCount
    dim FWCount
	dim RestoreCount
    dim PatchCount
	dim CurrentUserEmail
	dim strPMRDate
	dim RTMCommentsTemplate
	dim BIOSCommentsTemplate
    dim FWCommentsTemplate
	dim RestoreCommentsTemplate
	dim ImageCommentsTemplate
    dim strCDPartNumber
    dim strCDPartNumber2
    dim partnerId
    
	
	'	RTMCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
'							"1. Reason for RTM" & vbcrlf & _
'							"2. Special circumstances, patches, schedule updates, future expectations." & vbcrlf & _
'							"3. Any factory holds affected/lifted and actions surrounding the issue." & vbcrlf & _
'							"4. Any special rules set by DCR/AVs/EAs/Factory Memos/Etc." & vbcrlf & _
'							"5. Any platform specific additional comments."

    RTMCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                          "1. Reason for RTM" & vbcrlf & _
                          "2. Special circumstances, patches, schedule updates, future expectations." & vbcrlf & _
                          "3. Any factory holds affected/lifted and actions surrounding the issue." & vbcrlf & _
                          "4. Any special rules set by DCR/AVs/EAs/Factory Memos/Etc." & vbcrlf & _
                          "5. Any platform specific additional comments."

	
	BIOSCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
							"1. Any special instructions regarding BIOS cut-ins, staggered schedules for release, parallel releases, risk releases." & vbcrlf & _
							"2. Any information regarding updates to VBIOS, ME, AMT, MRC, etc that are noteworthy." & vbcrlf & _
							"3. Any platform BIOS specific additional comments." 

	FWCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
							"1. Any special instructions regarding Firmware cut-ins, staggered schedules for release, parallel releases, risk releases." & vbcrlf & _
                            "2. Firmware Name " & vbcrlf & _
                            "3. Firmware version (i.e. 1.00 A,5)" & vbcrlf & _
							"4. Any information regarding updates to Firmware that are noteworthy." & vbcrlf & _
							"5. Any platform Firmware specific additional comments." 

	RestoreCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                              "1. Any part numbers for older media that may not be present in wizard." & vbcrlf & _
                              "2. Verification of media transfer to replication houses and date of arrival." & vbcrlf & _
                              "3. Name of deliverable (i.e. ODIE_1.0_Win7_DRDVD)" & vbcrlf & _
                              "4. Part Number for media to be RTM'd." & vbcrlf & _
                              "5. Version of restore media (i.e 1.28 A,1)" & vbcrlf & _
                              "6. PMR ID for restore media" & vbcrlf & _
                              "7. Any other restore media specific additional comments."

	ImageCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                            "1. Final Rev numbers of the images being released." & vbcrlf & _
                            "2. Verification of image transfer to the production servers or schedule for final arrival." & vbcrlf & _
                            "3. Name of the production servers on which the images are housed." & vbcrlf & _
                            "4. Any special requirements and/or dependencies for the images being release (BIOS/FW/Patches/MSCU/etc)." & vbcrlf & _
                            "5. System WHQL completion dates per required OS." & vbcrlf & _
                            "6. MDA log completion dates per required OS." & vbcrlf & _
                            "7. 2PP requirements for all supported images." & vbcrlf & _
                            "8. Marketing Name." & vbcrlf & _
                            "9. PCR File Tested (Not for production)." & vbcrlf & _
                            "10.Be sure to attach the .SCMX file" & vbcrlf & _
                            "11. Provide the SCMX file name." & vbcrlf & _
                            "12. Provide the PCR file name" & vbcrlf & _
                            "13. Provide the released ML numbers after updating the final version online." & vbcrlf & _
                            "14. Any other image specific additional comments."


	PatchCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                            "1.	Patch Name (i.e. CSI Patch – 3PP Addon)" & vbcrlf & _
                            "2.	CSI Patch Part Number" & vbcrlf & _
                            "3.	CSI Patch version (i.e. 1.00 A,5)" & vbcrlf & _
                            "4.	CSI Patch PRISM Revision number"
 
    Dim strRTMTitle		     :  strRTMTitle=""
    Dim strRTMDate		     :  strRTMDate=  formatdatetime(now,vbshortdate)	
    Dim stroptCutIn		     :  stroptCutIn=""		
    Dim stroptPhaseIn		 :  stroptPhaseIn=""	
    Dim stroptWebOnly		 :  stroptWebOnly=""
    Dim strchkBIOS		     :  strchkBIOS=""		
    Dim strchkRestore		 :  strchkRestore=""		
    Dim strchkPatch		     :  strchkPatch=""	
    Dim strchkFW		     :  strchkFW=""	
    Dim strchkImages		 :  strchkImages=""
    Dim strchkSCMX	         :  strchkSCMX=""
    Dim strAttachment1	     :  strAttachment1=""
    ' For the Product RTM Section :
    Dim strBuildLevelAlerts		     :  strBuildLevelAlerts=""	
    Dim strDistributionAlerts	     :  strDistributionAlerts=""
    Dim strCertificationAlerts	     :  strCertificationAlerts=""		
    Dim strWorkflowAlerts		     :  strWorkflowAlerts=""		
    Dim strAvailabilityAlerts	     :  strAvailabilityAlerts=""	
    Dim strDeveloperAlerts		     :  strDeveloperAlerts=""	
    Dim strRootDeliverableAlerts     :  strRootDeliverableAlerts=""
    Dim strOTSPrimaryAlerts          :  strOTSPrimaryAlerts=""

    Dim strBuildLevelCmmts		     :  strBuildLevelCmmts=""	
    Dim strDistributionCmmts	     :  strDistributionCmmts=""
    Dim strCertificationCmmts	     :  strCertificationCmmts=""		
    Dim strWorkflowCmmts		     :  strWorkflowCmmts=""		
    Dim strAvailabilityCmmts	     :  strAvailabilityCmmts=""	
    Dim strDeveloperCmmts		     :  strDeveloperCmmts=""	
    Dim strRootDeliverableCmmts      :  strRootDeliverableCmmts=""
    Dim strOTSPrimaryCmmts           :  strOTSPrimaryCmmts=""

    Dim strRTMCommentsTemplate		 :  strRTMCommentsTemplate=""	
    Dim strBIOSCommentsTemplate	     :  strBIOSCommentsTemplate=""
    Dim strFWCommentsTemplate	     :  strFWCommentsTemplate=""		
    Dim strPatchCommentsTemplate	 :  strPatchCommentsTemplate=""		
    Dim strRestoreCommentsTemplate	 :  strRestoreCommentsTemplate=""	
    Dim strImageCommentsTemplate	 :  strImageCommentsTemplate=""	

	BIOSCount = 0
    FWCount = 0
	RestoreCount = 0
    PatchCount=0
	
	strProductName = ""
    blnFusion = false
	strLastRoot = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentWorkgroupID = rs("WorkgroupID") & ""
		CurrentUserEmail = rs("Email") & ""
        partnerId = rs("PartnerID") & ""
	end if
	rs.Close
	
	if request("ID") = "" then
		blnprodFound = false
	else

      if  request("ProductRTMID") <> "0" Then		
         		
        rs.open "spGetProductRTM " & clng(request("ProductRTMID")),cn		
		if not (rs.EOF and rs.BOF) then		
		    strRTMTitle = rs("Title")		
            strAttachment1=rs("Attachment1")	
            strRTMDate=  formatdatetime(rs("rtmDate"),vbshortdate)		
            ItemsToRTMArray = Split(rs("typeid"),",")		
         		
             For i = 0 to ubound(ItemsToRTMArray)		
                    If ItemsToRTMArray(i)="1" Then		
                        strchkBIOS="checked"		
                    ElseIf ItemsToRTMArray(i)="2" Then		
                        strchkRestore="checked"		
                    ElseIf ItemsToRTMArray(i)="3" Then		
                        strchkPatch="checked"		
                    ElseIf ItemsToRTMArray(i)="4" Then		
                        strchkFW="checked"		
                    End If		
             Next		
        
        if trim(rs("Affectivity")) = "Immediate (Rework All Units)" then		
			        stroptCutIn = "checked"		
			elseif trim(rs("Affectivity")) = "Phase-in" then		
			        stroptPhaseIn ="checked"		
            elseif trim(rs("Affectivity")) ="Web Release Only" then		
			        stroptWebOnly = "checked"		
			end if		
           
        end if		

        if trim(rs("comments"))<>"" then
            strRTMCommentsTemplate=trim(rs("comments"))
        else
            strRTMCommentsTemplate=RTMCommentsTemplate
        end if
        if trim(rs("bioscomments"))<>"" then
            strBIOSCommentsTemplate=trim(rs("bioscomments"))
        else    
            strBIOSCommentsTemplate=BIOSCommentsTemplate
        end if
        if trim(rs("FWComments"))<>"" then
            strFWCommentsTemplate=trim(rs("FWComments"))
        else    
            strFWCommentsTemplate=FWCommentsTemplate
        end if
        if trim(rs("patchcomments"))<>"" then
            strPatchCommentsTemplate=trim(rs("patchcomments"))
        else    
            strPatchCommentsTemplate=PatchCommentsTemplate
        end if
        if trim(rs("restorecomments"))<>"" then
            strRestoreCommentsTemplate=trim(rs("restorecomments"))
        else
            strRestoreCommentsTemplate=RestoreCommentsTemplate
        end if
        if trim(rs("imagecomments"))<>"" then
            strImageCommentsTemplate=trim(rs("imagecomments"))
        else        
            strImageCommentsTemplate=ImageCommentsTemplate
        end if

        rs.Close		

        rs.open "spGetProductRTMImageSCMX " & clng(request("ProductRTMID")),cn
		if not (rs.EOF and rs.BOF) then		
                    If clng(rs("ImageCount"))>0 Then		
                        strchkImages="checked"
                    ElseIf clng(rs("SCMXCount"))>0 Then		
                        strchkSCMX="checked"		
                    End If		
        
        end if		
        rs.Close

        

         set rsnewalert = server.CreateObject("ADODB.recordset")
         rsnewalert.open "splistProductRTMAlerts " & clng(request("ProductRTMID")),cn		
		if not (rsnewalert.EOF and rsnewalert.BOF) then		
           do while not rsnewalert.eof
              if trim(rsnewalert("name")) = "Build Level" then		
			          strBuildLevelAlerts = "checked"	
                      strBuildLevelCmmts=rsnewalert("Comments")	
		       elseif trim(rsnewalert("name")) = "Distribution" then		
			          strDistributionAlerts ="checked"	
                      strDistributionCmmts=rsnewalert("Comments")
               elseif trim(rsnewalert("name")) ="Certification" then		
			          strCertificationAlerts = "checked"	
                      strCertificationCmmts=rsnewalert("Comments")
		       elseif trim(rsnewalert("name")) = "Workflow" then		
			          strWorkflowAlerts ="checked"		
                      strWorkflowCmmts=rsnewalert("Comments")
               elseif trim(rsnewalert("name")) ="Availability" then		
			          strAvailabilityAlerts = "checked"
                      strAvailabilityCmmts=rsnewalert("Comments")
               elseif trim(rsnewalert("name")) = "Developer" then		
			          strDeveloperAlerts ="checked"		
                      strDeveloperCmmts=rsnewalert("Comments")
               elseif trim(rsnewalert("name")) ="Root" then		
			          strRootDeliverableAlerts = "checked"
                      strRootDeliverableCmmts=rsnewalert("Comments")
               elseif trim(rsnewalert("name")) = "OTS Primary" then		
			          strOTSPrimaryAlerts ="checked"		
                      strOTSPrimaryCmmts=rsnewalert("Comments")
          
              end if
            rsnewalert.movenext
          loop
      end if		
      rsnewalert.Close		
      set rsnewalert =nothing
    end if

		rs.Open "spGetProductVersion " & clng(request("ID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strSEPMID = rs("SEPMID") & ""
			strTestLeadID = rs("SeTestLead") & ""
			strPMID = rs("PMID") & ""
            blnFusion = rs("Fusion") & ""
			strPMEmail = rs("SEPMEmail") & ""
			strPMName = rs("SEPMName") & ""
			strProductName = rs("Name") & " " & rs("Version")
			strDevCenter = rs("DevCenter") & ""
			if trim(rs("RTMNotifications") & "") = "" then
			    strDistribution = CurrentUserEmail & ";" & trim(rs("Distribution") & "")
			elseif instr(trim(rs("RTMNotifications") & ""),CurrentUserEmail) < 1 then
			    strDistribution = CurrentUserEmail & ";" & trim(rs("RTMNotifications") & "")
            else
			    strDistribution = trim(rs("RTMNotifications") & "")
			end if
			blnProdFound = true
			'strMSG = "Please SMR the following Deliverables for " & strProductName & " as soon as possible. The versions listed must be released because they are required to support the factory images. Additional “upgrade” versions may be released at your discretion." & vbcrlf & vbcrlf & "[Their Deliverables Listed Here]" & vbcrlf & vbcrlf
		else
			blnProdFound = false
		end if
		rs.Close
		
	end if
	

    if not blnProdFound then
        response.write "Unable to find the requested product."
    else
    %>

    <font size="4" face="verdana"><b>Product RTM Wizard for <%=strProductName%></b></font>
    <br>
    <br>
    <font size="2" face="verdana"><b><label ID="lblTitle">Enter General RTM Information</label></b></font>
    <%

    %>

    <form id="frmMain" method="post" action="MilestoneSignoffSave.asp">
        <div id="tabGeneral" style="display: ">
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120"><font size="2" face="verdana"><b>RTM&nbsp;Title:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                    <td>
                        <input style="width: 100%" id="txtRTMName" name="txtRTMName" type="text" value="<%=strRTMTitle%>" maxlength="120">
                    </td>
                </tr>
                <tr>
                    <td width="120"><font size="2" face="verdana"><b>RTM&nbsp;Date:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                    <td>
                        <input style="width: 120px" id="txtRTMDate" name="txtRTMDate" type="text" class="dateselection" value="<%=strRTMDate%>">
                    </td>
                </tr>
                <tr>
                    <td width="120"><font size="2" face="verdana"><b>Items&nbsp;to&nbsp;RTM:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td>
                        <table>
                            <tr>
                                <td>
                                    <input id="chkBIOS" value="1" <%=strchkBIOS%> name="chkBIOS" type="checkbox" onclick="chkBIOS_onclick(); chkStandalone_onclick();">&nbsp;<font id="BIOSTextColor" color="black">System&nbsp;BIOS</font><font color="red" size="1" face="verdana" id="BIOSDisabled">&nbsp;</font>
                                </td>
                                <td>
                                    <input id="chkRestore" value="1" <%=strchkRestore%> name="chkRestore" type="checkbox" onclick="chkStandalone_onclick();">&nbsp;<font id="RestoreTextColor" color="black">Restore&nbsp;Media</font>&nbsp;<font color="red" size="1" face="verdana" id="RestoreDisabled">&nbsp;</font>
                                </td>
                                <td>
                                    <input id="chkPatch" value="1" <%=strchkBIOS%> name="chkPatch" type="checkbox" onclick="chkStandalone_onclick();">&nbsp;<font id="PatchTextColor" color="black">Patches</font>&nbsp;<font color="red" size="1" face="verdana" id="PatchDisabled">&nbsp;</font>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <input id="chkFW" value="1" <%=strchkFW%> name="chkFW" type="checkbox" onclick="chkFW_onclick();">&nbsp;<font id="FWTextColor" color="black">Firmware</font><font color="red" size="1" face="verdana" id="FWDisabled">&nbsp;</font>
                                </td>
                                <td>
                                    <input id="chkImages" value="1" <%=strchkImages%> name="chkImages" type="checkbox" onclick="chkStandalone_onclick();">&nbsp;<font id="ImagesTextColor" color="black">Images</font>&nbsp;<font color="red" size="1" face="verdana" id="ImagesDisabled">&nbsp;</font>
                                </td>
                                <td>
                                    <input id="chkSCMX" value="1" <%=strchkSCMX%> name="chkSCMX" type="checkbox" onclick="chkSCMX_onclick();">&nbsp;<font id="SCMXTextColor" color="black">SCMX Only</font>&nbsp;<font color="red" size="1" face="verdana" id="SCMXDisabled">&nbsp;</font>

                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr id="BIOSAffectivityRow" style="display: none">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Affectivity:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                    <td>
                        <input id="optCutIn" <%=stroptCutIn%> name="optPhaseIn" type="radio" value="0">
                        Immediate (Rework All Units)
            <input id="optPhaseIn" <%=stroptPhaseIn%> name="optPhaseIn" type="radio" value="1">
                        Phase-in
            <input id="optWebOnly" <%=stroptWebOnly%> name="optPhaseIn" type="radio" value="2">
                        Web Release Only
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtRTMComments" name="txtRTMComments" cols="80" rows="15" onfocus="return txtRTMComments_onfocus()" onblur="return txtRTMComments_onblur()"><%=strRTMCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtRTMCommentsTemplate" name="txtRTMCommentsTemplate"><%=RTMCommentsTemplate%></textarea>
                    </td>
                </tr>
                <%
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    'if currentuserid = 30 then 
                %>
                <!--//<TR>
		<TD width=120 valign=top><font size=2 face=verdana><b>SCMX:</b></font></td>
		<td>
            <a href="">Upload</a>
       </TD>
	</TR>//-->
                <%'end if %>
            </table>
        </div>


        <div id="tabBIOS" style="display: none">
            <%
            rs.open "spListBIOSVersions4Productrtm " & clng(request("ID")) & ",3,"& clng(request("ProductRTMID")),cn
            if not(rs.eof and rs.bof) then
            %>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td width="100%">
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtBIOSComments" name="txtBIOSComments" cols="80" rows="7" onfocus="return txtBIOSComments_onfocus()" onblur="return txtBIOSComments_onblur()"><%=strBIOSCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtBIOSCommentsTemplate" name="txtBIOSCommentsTemplate"><%=BIOSCommentsTemplate%></textarea>
                    </td>
                </tr>
            </table>
            <br>
            <table style="border-left: gray thin solid; border-right: gray thin solid; border-bottom: gray thin solid" class="ImageTable" cellpadding="2" cellspacing="0" bgcolor="cornsilk" width="100%">
                <%
            else
                Response.write "<font size=2 color=red face=verdana>There are no System BIOS delverables available to RTM on this product.</font>"
            end if
            strLastRoot = ""
            do while not rs.eof
        	    BIOSCount = BIOSCount + 1
                if trim(strLastRoot) <> trim(rs("DeliverableName") & "") then
                    response.Write "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=6 bgcolor=wheat><b>&nbsp;" & rs("DeliverableName") & "</b></td></tr>"        
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""25%""><b>&nbsp;Notes&nbsp;</b></TD><TD width=""75%""><b>&nbsp;RTMed&nbsp;</b></TD></tr>"
                end if
                strLastRoot = trim(rs("DeliverableName") & "")
                response.Write "<tr bgcolor=ivory>"    
                if rs("deliverabletargeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSList"" name=""chkBIOSList""  type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSList"" name=""chkBIOSList""  type=""checkbox""></td>"        
                end if
                response.Write "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("Version") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>&nbsp;" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>" 
                if rs("rtmed") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSListRTMed"" name=""chkBIOSListRTMed"" disabled=""true"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSListRTMed"" name=""chkBIOSListRTMed"" disabled=""true"" type=""checkbox""></td>"        
                end if
                response.Write "</tr>" & vbcrlf        
                rs.movenext
            loop
            if not(rs.eof and rs.bof) then
                response.Write "</table>"        
            end if
            rs.close
                %>
        </div>


        <div id="tabFW" style="display: none">
            <%
            rs.open "spListFW4ProductRTM " & clng(request("ID")) &","& clng(request("ProductRTMID")),cn
            if not(rs.eof and rs.bof) then
            %>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td width="100%">
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtFWComments" name="txtFWComments" cols="80" rows="7" onfocus="return txtFWComments_onfocus()" onblur="return txtFWComments_onblur()"><%=strFWCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtFWCommentsTemplate" name="txtFWCommentsTemplate"><%=FWCommentsTemplate%></textarea>
                    </td>
                </tr>
            </table>
            <br>
            <table style="border-left: gray thin solid; border-right: gray thin solid; border-bottom: gray thin solid" class="ImageTable" cellpadding="2" cellspacing="0" bgcolor="cornsilk" width="100%">
                <%
            else
                Response.write "<font size=2 color=red face=verdana>There are no Firmware delverables available to RTM on this product.</font>"
            end if
            strLastRoot = ""
            do while not rs.eof
        	    FWCount = FWCount + 1
                if trim(strLastRoot) <> trim(rs("DeliverableName") & "") then
                    response.Write "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=6 bgcolor=wheat><b>&nbsp;" & rs("DeliverableName") & "</b></td></tr>"        
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""25%""><b>&nbsp;Notes&nbsp;</b></TD><TD width=""75%""><b>&nbsp;RTMed&nbsp;</b></TD></tr>"
                end if
                strLastRoot = trim(rs("DeliverableName") & "")
                response.Write "<tr bgcolor=ivory>"    
                if rs("deliverabletargeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkFWList"" name=""chkFWList"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkFWList"" name=""chkFWList"" type=""checkbox""></td>"        
                end if
                response.Write "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("Version") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>&nbsp;" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>" 
                if rs("rtmed") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkFWListRTMed"" name=""chkFWListRTMed"" disabled=""true"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkFWListRTMed"" name=""chkFWListRTMed"" disabled=""true"" type=""checkbox""></td>"        
                end if
                response.Write "</tr>" & vbcrlf        
                rs.movenext
            loop
            if not(rs.eof and rs.bof) then
                response.Write "</table>"        
            end if
            rs.close
                %>
        </div>

        <div id="tabPatch" style="display: none">
            <%
            rs.open "spListPatches4ProductRTM " & clng(request("ID"))&","& clng(request("ProductRTMID")),cn
            if not(rs.eof and rs.bof) then
            %>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td width="100%">
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtPatchComments" name="txtPatchComments" cols="80" rows="5" onfocus="return txtPatchComments_onfocus()" onblur="return txtPatchComments_onblur()"><%=strPatchCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtPatchCommentsTemplate" name="txtPatchCommentsTemplate"><%=PatchCommentsTemplate%></textarea>
                    </td>
                </tr>
            </table>
            <br>
            <table style="border-left: gray thin solid; border-right: gray thin solid; border-bottom: gray thin solid" class="ImageTable" width="100%" id="PatchTable" cellspacing="0" cellpadding="2">

                <%
            else
                Response.write "<font size=2 color=red face=verdana>There are no Patch delverables available to RTM on this product.</font>" & vbcrlf
            end if
            strLastRoot = ""
            do while not rs.eof
        	    PatchCount = PatchCount + 1
                if trim(strLastRoot) <> trim(rs("Name") & "") then
                    response.Write vbcrlf & "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=6 bgcolor=wheat><b>&nbsp;" & rs("Name") & "</b></td></tr>" & vbcrlf       
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""25%""><b>&nbsp;Notes&nbsp;</b></TD><TD width=""75%""><b>&nbsp;RTMed&nbsp;</b></TD></tr>" & vbcrlf
                end if
                strLastRoot = trim(rs("Name") & "")
                response.Write "<tr bgcolor=ivory>" & vbcrlf
                
                strPatchContents = "" 
               	set rs2 = server.CreateObject("ADODB.recordset")
                rs2.open "spGetSelectedDepends " & clng(rs("ID")),cn
                do while not rs2.eof
                    if strPatchContents <> "" then
                        strPatchContents = strPatchContents & "<BR>" 
                    end if
                    strPatchContents = strPatchContents & rs2("Name") & " [" & rs2("Version")
                    if trim(rs2("revision")&"") <> "" then
                        strPatchContents = strPatchContents & "," & rs2("revision")
                    end if
                    if trim(rs2("pass")&"") <> "" then
                        strPatchContents = strPatchContents & "," & rs2("pass")
                    end if
                    rs2.movenext
                loop
                rs2.close   
            	set rs2 = nothing

                if trim(strPatchContents) = "" then
                    strPatchContents = "&nbsp;"
                else
                    strPatchContents = strPatchContents & "]"
                    strPatchContents = server.HTMLEncode(strPatchContents)
                end if
                  
                if rs("deliverabletargeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchList"" name=""chkPatchList"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchList"" name=""chkPatchList"" type=""checkbox""></td>"        
                end if
                response.Write "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("Version") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>&nbsp;" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>"  
                if rs("rtmed") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchListRTMed"" name=""chkPatchListRTMed"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchListRTMed"" name=""chkPatchListRTMed"" type=""checkbox""></td>"        
                end if
                response.Write "</tr>" & vbcrlf          
                rs.movenext
            loop
            if not(rs.eof and rs.bof) then
                response.Write "</table>"        
            end if
            rs.close
                %>
        </div>

        <div id="tabRestore" style="display: none">
            <%
            rs.open "spListRestoreMedia4ProductRTM " & clng(request("ID"))&","& clng(request("ProductRTMID")),cn
            if not(rs.eof and rs.bof) then
            %>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td width="100%">
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtRestoreComments" name="txtRestoreComments" cols="80" rows="9" onfocus="return txtRestoreComments_onfocus()" onblur="return txtRestoreComments_onblur()"><%=strRestoreCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtRestoreCommentsTemplate" name="txtRestoreCommentsTemplate"><%=RestoreCommentsTemplate%></textarea>
                    </td>
                </tr>
            </table>
            <br>
            <table style="border-left: gray thin solid; border-right: gray thin solid; border-bottom: gray thin solid" class="ImageTable" width="100%" id="TableRestore" cellspacing="0" cellpadding="2">
                <%
            else
                Response.write "<font size=2 color=red face=verdana>There are no Restore Media delverables available to RTM on this product.</font>"
            end if
            strLastRoot = ""
            do while not rs.eof
            	RestoreCount = RestoreCount + 1
                if trim(strLastRoot) <> trim(rs("Name") & "") then
                    response.Write "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=8 bgcolor=wheat><b>&nbsp;" & rs("Name") & "</b></td></tr>"        
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;Part&nbsp;</b></TD><TD><b>&nbsp;PMR&nbsp;Date&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""25%""><b>&nbsp;Notes&nbsp;</b></TD><TD width=""75%""><b>&nbsp;RTMed&nbsp;</b></TD></tr>"
                end if
                if rs("MultiLanguage") = 0 then
                    strCDPartNumber = ""
                    strCDPartNumber2 = ""
                    set rs2 = server.CreateObject("ADODB.recordset")
                    rs2.open "spGetSelectedLanguages " & rs("ID"),cn,adOpenStatic
                    do while not rs2.EOF
                        if trim(rs2("cdpartnumber") & "") <> "" then
                            if strCDPartNumber <> "" then
                                strCDPartNumber = strCDPartNumber & "<BR>" & rs2("Abbreviation") & ":&nbsp;" & server.HTMLEncode(replace(rs2("cdPartNumber")," ",""))
                                strCDPartNumber2 = strCDPartNumber2 & "<BR>&nbsp;" & rs2("Abbreviation") & ":&nbsp;" & server.HTMLEncode(replace(rs2("cdPartNumber")," ",""))
                            else
                                strCDPartNumber = rs2("Abbreviation") & ":&nbsp;" & server.HTMLEncode(replace(rs2("cdPartNumber")," ",""))
                                strCDPartNumber2 = rs2("Abbreviation") & ":&nbsp;" & server.HTMLEncode(replace(rs2("cdPartNumber")," ",""))
                            end if
                        end if
                        rs2.MoveNext
                    loop
                    rs2.Close
                    set rs2 = nothing
                else
                    strCDPartNumber = server.HTMLEncode(replace(replace(rs("cdPartNumber")," ",""),",","<BR>"))
                    strCDPartNumber2 = server.HTMLEncode(replace(replace(rs("cdPartNumber")," ",""),",","<BR>&nbsp;"))
                end if
                strLastRoot = trim(rs("Name") & "")
                strVersion = rs("Version") & ""
                if trim(rs("Revision") & "") <> "" then
                    strversion = strVersion & "," & rs("Revision")
                end if
                if trim(rs("Pass") & "") <> "" then
                    strversion = strVersion & "," & rs("pass")
                end if
                response.Write "<tr bgcolor=ivory>"  
                
                if trim(rs("PMRDate") & "") = "" then
                    strPMRDate = ""
                else
                    strPMRDate = formatdatetime(rs("PMRDate"),vbshortdate)
                end if
                if rs("deliverabletargeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & server.HTMLEncode(rs("Name")) & """ PreviewVersion=""" & server.HTMLEncode(strVersion) & """ PreviewPMR=""" & server.HTMLEncode(strPMRDate) & """ PreviewPart=""" & strCDPartNumber & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkRestoreList"" name=""chkRestoreList"" checked type=""checkbox"" /></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & server.HTMLEncode(rs("Name")) & """ PreviewVersion=""" & server.HTMLEncode(strVersion) & """ PreviewPMR=""" & server.HTMLEncode(strPMRDate) & """ PreviewPart=""" & strCDPartNumber & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkRestoreList"" name=""chkRestoreList"" type=""checkbox"" /></td>"        
                end if
                response.Write "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & strVersion & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td nowrap valign=top>&nbsp;" & strcdPartNumber2 & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & strPMRDate & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>"   
                 if rs("rtmed") then      
                    response.Write "<td valign=top><input PreviewName=""" & server.HTMLEncode(rs("Name")) & """ PreviewVersion=""" & server.HTMLEncode(strVersion) & """ PreviewPMR=""" & server.HTMLEncode(strPMRDate) & """ PreviewPart=""" & strCDPartNumber & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkRestoreListRTMed"" name=""chkRestoreListRTMed"" checked type=""checkbox"" /></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & server.HTMLEncode(rs("Name")) & """ PreviewVersion=""" & server.HTMLEncode(strVersion) & """ PreviewPMR=""" & server.HTMLEncode(strPMRDate) & """ PreviewPart=""" & strCDPartNumber & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkRestoreListRTMed"" name=""chkRestoreListRTMed"" type=""checkbox"" /></td>"        
                end if
                response.Write "</tr>" & vbcrlf          
                rs.movenext
            loop
            if not(rs.eof and rs.bof) then
                response.Write "</table>"        
            end if
            rs.close
                %>
        </div>
        <div id="tabImages" style="display: none">
            <%
    dim imagecount
    dim ChildRowCount
	strAllImages = ""
    'blnFusion=false
    if blnFusion and false then 'This was the method where they had to pick a whole product drop  ' this section never runs.
        imagecount=0
	    rs.open "spListImagesForProduct2RTM " & clng(request("ID"))&","& clng(request("ProductRTMID")),cn,adOpenForwardOnly
	    if rs.EOF and rs.BOF then
		    strAllImages = "<TR><TD colspan=11><FONT size=1 face=verdana>No images defined for this product.</font></td></tr>"
	    else
            strLastProductDrop = ""
            do while not rs.eof
                if lcase(trim(strLastProductDrop)) <> lcase(trim(rs("ProductDrop") & "")) then
'                    strAllImages = strAllImages & "<tr><td><input id=""chkProductDrop"" type=""checkbox"" /></td><td colspan=10>" & rs("ProductDrop") & "</td></tr>"
			        if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				        RowStyle = "style=display:none"
			        else
    				    rowStyle=""
			        end if
                    if strLastProductDrop <> "" then
                        strAllImages= strAllImages & "</table><br>"
                    end if
			         strAllImages= strAllImages &  "<TR " & rowStyle & " id=DropRow" & rs("ProductDropID") & " LANGUAGE=javascript onmouseover=""return DropRow_onmouseover(" & rs("ProductDropID") & ")"" onmouseout=""return DropRow_onmouseout(" & rs("ProductDropID") & ")"" onclick=""return DropRow_onclick(" & rs("ProductDropID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("ProductDropID") & """ name=""Drop" & rs("ProductDropID") & """ type=""checkbox"" class=""chkDrop"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkDrop_onclick()""></TD>" 
                     strAllImages= strAllImages & "<td>" & rs("ProductDrop") & "&nbsp;&nbsp;&nbsp;</td>"
                     strAllImages= strAllImages & "<td colspan=10>" & rs("OSList") & "</td>"
			         strAllImages= strAllImages &  "<TR style=""display:none"" id=DropContents" & rs("ProductDropID") & " bgcolor=cornsilk><td>&nbsp;</td>"
                     strAllImages= strAllImages & "<td colspan=10>" 
                     strAllImages= strAllImages & "<BR><table width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD><b>Region</b></TD><TD><b>Brands</b></TD><TD><b>OS</b></TD><TD><b>Comments</b></TD></THEAD>"
                end if

                if trim(rs("ProductDrop") & "") =  "" then
                    strimagePreview = "FUSION|" & rs("ID") & "|No Product Drop Number Defined|"
                else
                    strimagePreview = "FUSION|" & rs("ID") & "|" & ucase(rs("ProductDrop")) & "|" 
	            end if
                strimagePreview = server.HTMLEncode(strimagePreview & rs("Brands") & "|" & rs("OS") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))

                strAllImages= strAllImages & "<tr>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & "<INPUT class=""" & trim(rs("ProductDropID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " style=""display:none"" name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16"">" & rs("OptionConfig") & "&nbsp;-&nbsp;" & rs("Region") & "<input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """>" &  "</td>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & rs("Brands") & "</td>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & rs("OS") & "</td>"
                strCommentCollection = ""
                if trim(rs("DefinitionComments") & "" ) <> "" then
                    strCommentCollection = trim(rs("DefinitionComments") & "" )
                end if
                if trim(rs("Comments") & "" ) <> "" and strCommentCollection <> "" then
                    strCommentCollection = strCommentCollection & "<br>" & trim(rs("DefinitionComments") & "" )
                elseif trim(rs("Comments") & "" ) <> ""  then
                    strCommentCollection = trim(rs("DefinitionComments") & "" )
                end if
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & strCommentCollection & "&nbsp;</td>"
                strAllImages= strAllImages & "</tr>"
                strLastProductDrop= rs("ProductDrop") & "" 
                imagecount = imagecount + 1
                rs.movenext
            loop
            if imagecount > 0 then
                strAllImages= strAllImages & "</td></tr></table>" 
            end if
        end if
        rs.close
    elseif blnFusion then ' This section is for converged notebooks ' for IRS Images

	    rs.open "spListImagesForProduct2RTM " & clng(request("ID"))&","& clng(request("ProductRTMID")),cn,adOpenForwardOnly
	    lastDefinition = ""
	    imagecount=0
        ChildRowCount=0
	    if rs.EOF and rs.BOF then
		    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></tr>"
	    else
		    do while not rs.EOF
			    imagecount = imagecount + 1
			    if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
				    strAllImages = strAllImages & strBase
				    strAllImages = strAllImages &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Tier</b></TD><TD><b>Config</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD><TD><b>RTMed</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
				    strRows = ""
				    YesCount = 0
				    NoCount = 0
				    MixedCount=0
                    ChildRowCount=0
			    end if
			    lastdefinition = rs("DefinitionID")
			
			    strLanguageList = rs("OSLanguage")
			    if trim(rs("OtherLanguage") & "") <> "" then
				    strLanguageList = strLanguageList & "," & rs("OtherLanguage")
			    end if	
			
			   ' strSavedLanguages = getLanguages(strImages,rs("ID"))
			   ' if strSavedLanguages = "" then
				'    strSavedLanguages = strLanguageList
			   ' end if
            
                if trim(rs("ProductDrop") & "") =  "" then
                    strimagePreview = "FUSION|" & rs("ID") & "|No Product Drop Number Defined|"
                else
                    strimagePreview = "FUSION|" & rs("ID") & "|" & ucase(rs("ProductDrop")) & "|" 
	            end if
                strimagePreview = server.HTMLEncode(strimagePreview & rs("Brands") & "|" & rs("OS") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
			    strCellColor = "ivory"
			    if instr(strImages,", " & rs("ID") & ",") > 0 or not blnImages then
					strCellColor = "ivory"
					YesCount = YesCount + 1
					if request("Type") = "1" then
						strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "' LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
					else
                          if clng(request("ProductRTMID")) >0 then  
                            if rs("targeted") then  
						        strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "'  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
                                ChildRowCount=ChildRowCount+1
                            else    
						        strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox""  id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "'  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
				            end if
                          else
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "'  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
                         end if
                    end if
			    else
				    if request("Type") = "1" then
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "' value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
				    else
                        if clng(request("ProductRTMID")) >0 then  
                            if rs("targeted") then  
					            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "' value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                                ChildRowCount=ChildRowCount+1
                            else    
					            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "' value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                            end if
                        else
					        strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage data-tier='" & rs("Tier") & "' value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                        end if
				    end if
				    NoCount = NoCount + 1
			    end if
			
			    if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				    RowStyle = "style=display:none"
			    else
				    rowStyle=""
			    end if
			    if YesCount = 0  and MixedCount=0 then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
                        if clng(request("ProductRTMID")) >0 then  
                          'if rs("targeted") then 
                            if clng(ChildRowCount)>0 then 
			                  strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" checked class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                          else
                              strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                          end if
                        else
                          strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" checked class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        end if
                  end if
			    elseif NoCount=0 and MixedCount=0  then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
                     if clng(request("ProductRTMID")) >0 then  
                        'if rs("targeted") then 
                         if clng(ChildRowCount)>0 then 
				            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                          else 
				            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase""  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        end if
			         else 
				        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                     end if
                end if
			      TotalImageDefsChecked= TotalImageDefsChecked + 1
			    else
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      else 
                     if clng(request("ProductRTMID")) >0 then  
                        'if rs("targeted") then  
                         if clng(ChildRowCount)>0 then 
				            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
                        else
                            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
                        end if
                     else
                        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			         end if
                end if
			      TotalImageDefsChecked = TotalImageDefsChecked + 1
			    end if
			    strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("ProductDrop") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("Brands") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" &  rs("OS")  & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("Comments") & "&nbsp;</TD>"
			    strBase = strBase &  "</tr>"

                dim strTierLink
                strTierLink = "&nbsp;"
                if (rs("Tier") & "") <> "" then
                    strTierLink = "<input type='button' onclick='return SelectAllSameTierImages(" & rs("Tier") & ");' title='Select all Tier " & rs("Tier") & "' value='" & rs("Tier") & "' style='border:0px none;'  data-chkimageid='chkImage" & rs("ID") & "' >"
                end if

			    strRows = strRows & "<TD style='text-align: center;'>" & strTierLink & "</TD>"
			    strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
			    strRows = strRows & "<TD width=130>" & rs("Region") & "</TD>"
			    strRows = strRows & "<TD>" & rs("countrycode") & "</TD>"

			    if trim(rs("OtherLanguage") & "") <> "" then
			        strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			    else
				    strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			    end if				
			    strRows = strRows & "<TD width=65>" & rs("Keyboard") & "</TD>"
			    strRows = strRows & "<TD width=75>" & rs("powercord") & "</TD>"
                 if rs("rtmed") then      
                    strRows = strRows & "<TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImageRTMed" & rs("ID") & " name=chkImageRTMed data-tier='" & rs("Tier") & "'  value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"        
                else
                     strRows = strRows & "<TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox""  id=chkImageRTMed" & rs("ID") & " name=chkImageRTMed data-tier='" & rs("Tier") & "'  value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"        
                
                end if
			    strRows = strRows & "</tr>"

			    rs.MoveNext
		    loop
		    if imagecount = 0 then
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No active images defined for this product.</font></td></tr>"
		    end if
		    strAllImages = strAllImages & strBase
		    strAllimages = strAllImages & "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Tier</b></TD><TD><b>Config</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD><TD><b>RTMed</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
		    strRows = ""
	    end if	
	    rs.Close
    else ' This section is for legacy notebooks  ' For excalibure Images
	    rs.open "spListImagesForProduct2RTM " & clng(request("ID"))&","& clng(request("ProductRTMID")),cn,adOpenForwardOnly
	    lastDefinition = ""
	    imagecount=0
        ChildRowCount=0
	    if rs.EOF and rs.BOF then
		    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></tr>"
	    else
		    do while not rs.EOF
			    imagecount = imagecount + 1
			    if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
				    strAllImages = strAllImages & strBase
				    strAllImages = strAllImages &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
				    strRows = ""
				    YesCount = 0
				    NoCount = 0
				    MixedCount=0
                    ChildRowCount=0
			    end if
			    lastdefinition = rs("DefinitionID")
			
			    strLanguageList = rs("OSLanguage")
			    if trim(rs("OtherLanguage") & "") <> "" then
				    strLanguageList = strLanguageList & "," & rs("OtherLanguage")
			    end if	
			
			    strSavedLanguages = getLanguages(strImages,rs("ID"))
			    if strSavedLanguages = "" then
				    strSavedLanguages = strLanguageList
			    end if
            
                if trim(rs("Skunumber") & "") =  "" then
                    strimagePreview = server.HTMLEncode(rs("ID") & "|No SKU Defined|" & rs("Brand") & "|" & rs("OS") & "|" & rs("SW") & "|" & rs("ImageType") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
                else
                    strimagePreview = server.HTMLEncode(rs("ID") & "|" & 	replace(ucase(rs("SKUNumber")) & "","XX",mid(rs("Dash"),2,2)) & "|" & rs("Brand") & "|" & rs("OS") & "|" & rs("SW") & "|" & rs("ImageType") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
	            end if
			    strCellColor = "ivory"
			    if instr(strImages,", " & rs("ID") & ",") > 0 or not blnImages then
				    if strLanguageList <> strSavedLanguages then
					    strCellColor = "mistyrose"
					    MixedCount = MixedCount + 1
					    if request("Type") = "1" then
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
					    else
                            if clng(request("ProductRTMID")) >0 then  
                                if rs("targeted") then
						             strRows = strRows & "<TR bgcolor=ivory><TD><INPUT indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strimagepreview & """></TD>"
                                     ChildRowCount= ChildRowCount+1
                                else
						             strRows = strRows & "<TR bgcolor=ivory><TD><INPUT indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strimagepreview & """></TD>"
                                end if
                            else
						        strRows = strRows & "<TR bgcolor=ivory><TD><INPUT indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strimagepreview & """></TD>"
					        end if
                         end if
				    else
					    strCellColor = "ivory"
					    YesCount = YesCount + 1
					    if request("Type") = "1" then
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
					    else
                            if clng(request("ProductRTMID")) >0 then  
                                if rs("targeted") then
						            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
                                    ChildRowCount= ChildRowCount+1
                                else
						            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox""  id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
                                end if
                            else
					            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagepreview & """></TD>"							
                            end if
					    end if
                    end if
			    else
				    if request("Type") = "1" then
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
				    else
                        if clng(request("ProductRTMID")) >0 then  
                           if rs("targeted") then
					            strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                                ChildRowCount= ChildRowCount+1
                            else
        					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                            end if
                        else
					        strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """></TD>"
                        end if
				    end if
				    NoCount = NoCount + 1
			    end if
			
			    if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				    RowStyle = "style=display:none"
			    else
				    rowStyle=""
			    end if
			    if YesCount = 0  and MixedCount=0 then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
                    if clng(request("ProductRTMID")) >0 then  
                        'if rs("targeted") then
                         if clng(ChildRowCount) >0 then 
			                strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" checked class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        else
			                strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        end if
                    else
			                strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" checked class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                    end if
			      end if
			    elseif NoCount=0 and MixedCount=0  then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
                     if clng(request("ProductRTMID")) >0 then  
                        'if rs("targeted") then  
                          if clng(ChildRowCount) >0 then 
        				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        else
        				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                        end if
                     else
        				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
                     end if
                    
			      end if
			      TotalImageDefsChecked= TotalImageDefsChecked + 1
			    else
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      else 
                     if clng(request("ProductRTMID")) >0 then  
                        'if rs("targeted") then  
                          if clng(ChildRowCount) >0 then 
        		            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
                        else
                            strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
                        end if
                     else
        		        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
                     end if
			      end if
			      TotalImageDefsChecked = TotalImageDefsChecked + 1
			    end if
			    strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("brand") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
			    strBase = strBase &  "</tr>"
			
			    strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
			    strRows = strRows & "<TD width=130>" & rs("Region") & "</TD>"
			    strRows = strRows & "<TD width=50>" & rs("CountryCode") & "</TD>"
			    strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
			    if trim(rs("OtherLanguage") & "") <> "" then
			      if request("Type") = "1" then
			        strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			      else
				    strRows = strRows & "<TD ID=""Row" & rs("ID") & """ bgcolor=" & strCellColor & " width=70>" & strSavedLanguages & "</a></TD>"
			      end if
			    else
				    strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			    end if				
			    strRows = strRows & "<TD width=65>" & rs("Keyboard") & "</TD>"
			    strRows = strRows & "<TD width=75>" & rs("Powercord") & "</TD>"
                if rs("rtmed") then      
                   strRows = strRows & "<TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImageRTMed" & rs("ID") & " name=chkImageRTMed data-tier='" & rs("Tier") & "'   value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"        
               else
                    strRows = strRows & "<TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox""  id=chkImageRTMed" & rs("ID") & " name=chkImageRTMed data-tier='" & rs("Tier") & "'  value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"        
               
               end if
			    strRows = strRows & "</tr>"

			    rs.MoveNext
		    loop
		    if imagecount = 0 then
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No active images defined for this product.</font></td></tr>"
		    end if
		    strAllImages = strAllImages & strBase
		    strAllimages = strAllImages & "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
		    strRows = ""
	    end if	
	    rs.Close
    end if

	if TotalImageDefsChecked > 0 then
		strAllChecked="checked"
	else
		strAllChecked=""
	end if
            %>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>File&nbsp;Upload:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td>
                        <div id="UploadAfddLinks1"><a href="javascript: UploadZip(1);">Upload</a></div>
                        <div id="UploadRemoveLinks1" style="display: none"><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id="UploadPath1"></label></div>
                        <input id="txtAttachmentPath1" name="txtAttachmentPath1" type="hidden" value="<%=strAttachment1%>" />

                        <input id="txtUploadPath1" name="txtUploadPath1" type="hidden" value="" />
                    </td>
                </tr>
                <tr id="trImageComments">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td width="100%">
                        <textarea style="width: 100%; color: blue; font-style: italic" id="txtImageComments" name="txtImageComments" cols="80" rows="18" onfocus="return txtImageComments_onfocus()" onblur="return txtImageComments_onblur()"><%=strImageCommentsTemplate%></textarea>
                        <textarea style="display: none" id="txtImageCommentsTemplate" name="txtImageCommentsTemplate"><%=ImageCommentsTemplate%></textarea>
                    </td>
                </tr>
            </table>
            <br>
            <div id="divImageTable">
                <% 
    if imagecount <> 0 then
                %>

                <%if blnFusion then %>
                <table class="ImageTable" width="100%" border="0" cellspacing="0" cellpadding="1">
                    <thead bgcolor="Wheat">
                        <% if request("Type") = "1" then%>
                        <th align="left">
                            <input disabled type="checkbox" id="chkAll" name="chkAll" style="width: 16; height: 16" <%=strAllChecked%> language="javascript" onclick="return chkAll_onclick()"></th>
                        <%else%>
                        <th align="left">
                            <input type="checkbox" id="chkAll" name="chkAll" style="width: 16; height: 16" <%=strAllChecked%> language="javascript" onclick="return chkAll_onclick()"></th>
                        <%end if%>
                        <th align="left">ID</th>
                        <th align="left">Product Drop</th>
                        <th align="left">Brands</th>
                        <th align="left">OS</th>
                        <th align="left" colspan="6" width="100%">Comments</th>
                    </thead>
                    <%=strAllImages%>
                </table>

                <%else%>
                <table class="ImageTable" width="100%" border="0" cellspacing="0" cellpadding="1">
                    <thead bgcolor="Wheat">
                        <% if request("Type") = "1" then%>
                        <th align="left">
                            <input disabled type="checkbox" id="chkAll" name="chkAll" style="width: 16; height: 16" <%=strAllChecked%> language="javascript" onclick="return chkAll_onclick()"></th>
                        <%else%>
                        <th align="left">
                            <input type="checkbox" id="chkAll" name="chkAll" style="width: 16; height: 16" <%=strAllChecked%> language="javascript" onclick="return chkAll_onclick()"></th>
                        <%end if%>
                        <th align="left">ID</th>
                        <th align="left">SKU</th>
                        <th align="left">Model</th>
                        <th align="left">OS</th>
                        <th align="left">Apps&nbsp;Bundle</th>
                        <th align="left">BTO/CTO</th>
                        <th align="left">Comments</th>
                    </thead>
                    <%=strAllImages%>
                </table>
                <%end if%>
                <%else
                response.write "<label id=""UploadPath1""></label>"
                Response.write "<div id=noImage><font size=2 color=red face=verdana>There are no Images available to RTM on this product.</font></div>"
		end if%>
            </div>
            <input style="display: none" type="checkbox" id="chkAllChecked" name="chkAllChecked">
        </div>

        <div id="tabAlerts" style="display: none">
            <font size="2" face="verdana"><b>Build Level Alerts: </b>&nbsp;&nbsp;<font id=BuildLevelTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkBuildLevel" name="chkBuildLevel" <%=strBuildLevelAlerts%> type="checkbox" value="1" language="javascript" onclick="chkBuildLevel_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtBuildLevelComments" name="txtBuildLevelComments" rows="4" cols="90" style="width: 100%"><%=strBuildLevelCmmts%></textarea>
                    </td>
                </tr>
                <tr id="BuildLevelAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtBuildLevelIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=1&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="BuildLevelIFrame" name="BuildLevelIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>
            <br>
            <font size="2" face="verdana"><b>Distribution Alerts: </b>&nbsp;&nbsp;<font id=DistributionTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkDistribution" name="chkDistribution" <%=strDistributionAlerts%> type="checkbox" value="1" language="javascript" onclick="chkDistribution_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtDistributionComments" name="txtDistributionComments" rows="4" cols="90" style="width: 100%"><%=strDistributionCmmts%></textarea>
                    </td>
                </tr>
                <tr id="DistributionAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtDistributionIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=2&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="DistributionIFrame" name="DistributionIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>

            <br>
            <font size="2" face="verdana"><b>Certification Alerts: </b>&nbsp;&nbsp;<font id=CertificationTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkCertification" name="chkCertification" <%=strCertificationAlerts%> type="checkbox" value="1" language="javascript" onclick="chkCertification_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtCertificationComments" name="txtCertificationComments" rows="4" cols="90" style="width: 100%"><%=strCertificationCmmts%></textarea>
                    </td>
                </tr>
                <tr id="CertificationAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtCertificationIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=3&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="CertificationIFrame" name="CertificationIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>

            <br>
            <font size="2" face="verdana"><b>Workflow Alerts: </b>&nbsp;&nbsp;<font id=WorkflowTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkWorkflow" name="chkWorkflow" <%=strWorkflowAlerts%> type="checkbox" value="1" language="javascript" onclick="chkWorkflow_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></td>
                    <td>
                        <textarea id="txtWorkflowComments" name="txtWorkflowComments" rows="4" cols="90" style="width: 100%"><%=strWorkflowCmmts%></textarea>
                    </td>
                </tr>
                <tr id="WorkflowAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtWorkflowIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=4&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="WorkflowIFrame" name="WorkflowIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>


            <br>
            <font size="2" face="verdana"><b>Availability Alerts: </b>&nbsp;&nbsp;<font id=AvailabilityTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkAvailability" name="chkAvailability" <%=strAvailabilityAlerts%> type="checkbox" value="1" language="javascript" onclick="chkAvailability_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtAvailabilityComments" name="txtAvailabilityComments" rows="4" cols="90" style="width: 100%"> <%=strAvailabilityCmmts%></textarea>
                    </td>
                </tr>
                <tr id="AvailabilityAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtAvailabilityIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=5&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="AvailabilityIFrame" name="AvailabilityIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>

            <br>
            <font size="2" face="verdana"><b>Developer Alerts: </b>&nbsp;&nbsp;<font id=DeveloperTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkDeveloper" name="chkDeveloper" <%=strDeveloperAlerts%> type="checkbox" value="1" language="javascript" onclick="chkDeveloper_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtDeveloperComments" name="txtDeveloperComments" rows="4" cols="90" style="width: 100%"><%=strDeveloperCmmts%> </textarea>
                    </td>
                </tr>
                <tr id="DeveloperAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtDeveloperIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=6&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="DeveloperIFrame" name="DeveloperIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>

            <br>
            <font size="2" face="verdana"><b>Root Deliverable Alerts: </b>&nbsp;&nbsp;<font id=RootTypeText color=green></font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkRoot" name="chkRoot" <%=strRootDeliverableAlerts%> type="checkbox" value="1" language="javascript" onclick="chkRoot_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtRootComments" name="txtRootComments" rows="4" cols="90" style="width: 100%"><%=strRootDeliverableCmmts%></textarea>
                    </td>
                </tr>
                <tr id="RootAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtRootIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=7&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="RootIFrame" name="RootIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>

            <br>
            <font size="2" face="verdana"><b>OTS Primary Alerts - <%=strproductname%>:&nbsp;&nbsp;</b><font id=OTSPrimaryFilterText color=green>(P0/P1 observations for <label id=OTSPrimaryTypeText></label>)</font><BR></font>
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Signoff:</b></font></td>
                    <td>
                        <input id="chkOTSPrimary" name="chkOTSPrimary" <%=strOTSPrimaryAlerts%> type="checkbox" value="1" language="javascript" onclick="chkOTSPrimary_onclick();">
                        I have reviewed these alerts.
                    </td>
                </tr>
                <tr>
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                    <td>
                        <textarea id="txtOTSPrimaryComments" name="txtOTSPrimaryComments" rows="4" cols="90" style="width: 100%"><%=strOTSPrimaryCmmts  %></textarea>
                    </td>
                </tr>
                <tr id="OTSPrimaryAlertDetails">
                    <td width="120" valign="top"><font size="2" face="verdana"><b>Alerts:</b></font></td>
                    <td>
                        <textarea id="txtOTSPrimaryIFramesrc" style="display: none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=clng(request("ID"))%>&Sections=8&TableOnly=1&RTMSignoff=1&RTMID=<%=clng(request("ProductRTMID"))%></textarea>
                        <iframe id="OTSPrimaryIFrame" name="OTSPrimaryIFrame" marginwidth="0" width="100%" src="../maint/blank_loading.htm"></iframe>
                    </td>
                </tr>

            </table>
        </div>


        <input type="hidden" id="txtProductID" name="txtProductID" value="<%=trim(clng(request("ID")))%>">
        <input type="hidden" id="txtProductRTMID" name="txtProductRTMID" value="<%=trim(clng(request("ProductRTMID")))%>">
        <input type="hidden" id="txtIsRTMAsDraft" name="txtIsRTMAsDraft" value="">
        <input type="hidden" id="txtProductName" name="txtProductName" value="<%=strProductName%>">
        <input type="hidden" id="txtCurrentUserEmail" name="txtCurrentUserEmail" value="<%=CurrentUserEmail%>">
        <input type="hidden" id="txtPartnerId" name="txtPartnerId" value="<%=partnerId%>">
        <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
        <div id="tabPreview" style="display: none">
            <table width="100%" border="0">
                <tr>
                    <td valign="top"><b>Email:&nbsp;&nbsp;</b></td>
                    <td width="100%">
                        <textarea style="width: 100%" id="txtNotify" name="txtNotify" rows="3"><%=strDistribution%></textarea></td>
                    <td valign="top">
                        <button style="height: 50" id="cmdAdd" name="cmdAdd" language="javascript" onclick="return cmdAdd_onclick()">
                            Address<br>
                            Book</button>
                    </td>
                </tr>
            </table>
            <font size="2" face="verdana"><b>Preview:</b></font>
            <div style="padding-left: 10; padding-right: 10; padding-bottom: 10; padding-top: 10; border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: auto; border-left: steelblue 1px solid; border-bottom: steelblue 1px solid; height: 100%; background-color: white" id="PreviewPage">
            </div>
        </div>
        <textarea style="display: none" id="txtEmailPreview" name="txtEmailPreview"></textarea>

    </form>
    <%
    dim strExistingRTMTitles
    'Load existing RTm titles
        strExistingRTMTitles = ""
		rs.Open "spListProductRTMTitles " & clng(request("ID")),cn,adOpenForwardOnly
        do while not rs.eof
            strExistingRTMTitles = strExistingRTMTitles & "<option>" & rs("Title") & "</option>"
            rs.movenext
        loop
        rs.close
        
    end if

	cn.Close
	set cn = nothing
	set rs = nothing

function GetLanguages(strImages, strID)
	dim strTemp
	
	if instr(strImages,trim(strID) & "=") = 0 then
		GetLanguages = ""
	else
		strTemp = mid(strimages,instr(strImages,trim(strID) & "=")+ len(trim(strID) & "="))
		strTemp = mid(strTemp,1,instr(strTemp,")") -1) 'Strip off )...
		GetLanguages = strTemp
	end if
	
end function


    %>
    <select style="display: none" id="cboTitles" name="cboTitles">
        <%=strExistingRTMTitles%>
    </select>
    <div id="dialog-charError" title="Characters Error">
        <p><b>Invalid characters detected(location marked with _):</b></p>
        <textarea style="width: 90%;" id="CharErrorMsg" cols="60" rows="15"></textarea>
    </div>
    <input style="display: none" id="txtImageCount" type="text" value="<%=imagecount%>">
    <input style="display: none" id="txtBIOSCount" type="text" value="<%=BIOScount%>">
    <input style="display: none" id="txtFWCount" type="text" value="<%=FWcount%>">
    <input style="display: none" id="txtRestoreCount" type="text" value="<%=Restorecount%>">
    <input style="display: none" id="txtPatchCount" type="text" value="<%=Patchcount%>">
</body>
</html>
