<%@  language="VBScript" %>
<html>
<head>
<title></title> 
    <link href="../../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../../Scripts/verifyEmailAddress.js"></script>
    <script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        $(function () {
            $("input:button").button();
            if ($.urlParam("Layout") == "pulsar2") {
                $('#cmdEditCancel').css('display', 'none');
            }
        });

        $.urlParam = function (name) {
            var results = new RegExp('[\?&]' + name + '=([^&#]*)')
                .exec(decodeURI(window.location.search));
            return (results !== null) ? results[1].toLowerCase() || 0 : false;
        }
    </script>

    <script type="text/javascript" src="../../includes/Date.asp"></script>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
    <!--

        function ltrim(s) {
            return s.replace(/^\s*/, "")
        }

        function VerifySave() {
            var blnSuccess;
            var i;
            var Pending;
            var ApproverRows;

            var blnGEOS = false;
            var blnCategory = false;
            var blnBusiness = false;
            var blnDescription = false;
            var blnBios = true;

            if (window.parent.frames["UpperWindow"].ProgramInput.hidCurrentUserPartner.value == "1") {
                if (window.parent.frames["UpperWindow"].ProgramInput.chkIDChange.checked && window.parent.frames["UpperWindow"].ProgramInput.txtJustificationTemplate.value == ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value))
                    window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value = "";
            }

            if ((window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3")
	    && (window.parent.frames["UpperWindow"].ProgramInput.txtID.value == "")
	    && (window.parent.frames["UpperWindow"].ProgramInput.chkBiosChange.checked)
	    && ((window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkOtherChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkDocChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkCommodityChange.checked)
        || (window.parent.frames["UpperWindow"].ProgramInput.chkImageChange.checked)
        || (window.parent.frames["UpperWindow"].ProgramInput.chkCategoryBiosChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkSKUChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkReqChange.checked))) {

                blnBios = false;
                window.alert("You cannot select other change categories if you select a BIOS Change Request (BCR).");
            }

            if (blnBios
	    && (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3")
	    && (window.parent.frames["UpperWindow"].ProgramInput.txtID.value == "")
	    && (window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked)
	    && ((window.parent.frames["UpperWindow"].ProgramInput.chkBiosChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkOtherChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkDocChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkCommodityChange.checked)
        || (window.parent.frames["UpperWindow"].ProgramInput.chkImageChange.checked)
        || (window.parent.frames["UpperWindow"].ProgramInput.chkCategoryBiosChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkSKUChange.checked)
	    || (window.parent.frames["UpperWindow"].ProgramInput.chkReqChange.checked))) {

                blnBios = false;
                window.alert("You cannot select other change categories if you select a Software Change Request (SCR).");
            }

            if (window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked) {
                if ((window.parent.frames["UpperWindow"].ProgramInput.hidDeliverableRootId.value == "") || (window.parent.frames["UpperWindow"].ProgramInput.hidDeliverableRootId.value == "0")) {
                    blnBios = false;
                    window.alert("You must select a deliverable impacted by this Software Change Request");
                }
            }

            if (blnBios
	    && (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3")
	    && (window.parent.frames["UpperWindow"].ProgramInput.txtID.value == "")
	    && ((window.parent.frames["UpperWindow"].ProgramInput.chkBiosChange.checked)
	        || (window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked))) {
                if (!((window.parent.frames["UpperWindow"].ProgramInput.rbNewBiosFeature.checked)
            || (window.parent.frames["UpperWindow"].ProgramInput.rbChangeBiosFeature.checked))) {
                    blnBios = false;
                    window.alert("You must select feature change or request.");
                    window.parent.frames["UpperWindow"].ProgramInput.rbChangeBiosFeature.select;
                    window.parent.frames["UpperWindow"].ProgramInput.rbChangeBiosFeature.focus;
                }
            }            
            if ((window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3") && (!((window.parent.frames["UpperWindow"].ProgramInput.chkBiosChange.checked)
	        || (window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked) || (window.parent.frames["UpperWindow"].ProgramInput.chkIDChange.checked) ))) // Change Request
            {
                //Verify Geo Is Checked
                if (window.parent.frames["UpperWindow"].ProgramInput.chkNA.checked)
                    blnGEOS = true;
                else if (window.parent.frames["UpperWindow"].ProgramInput.chkLA.checked)
                    blnGEOS = true;
                else if (window.parent.frames["UpperWindow"].ProgramInput.chkAPJ.checked)
                    blnGEOS = true;
                else if (window.parent.frames["UpperWindow"].ProgramInput.chkEMEA.checked)
                    blnGEOS = true;

                //Verify Business Unit
                if (window.parent.frames["UpperWindow"].ProgramInput.txtID.value == "" || window.parent.frames["UpperWindow"].ProgramInput.txtID.value > 5607) {
                    if (window.parent.frames["UpperWindow"].ProgramInput.chkCommercial.checked)
                        blnBusiness = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkConsumer.checked)
                        blnBusiness = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkSMB.checked)
                        blnBusiness = true;
                }
                else
                    blnBusiness = true;

                // Verify Change Category
                if (window.parent.frames["UpperWindow"].document.getElementsByName("chkReqChange") || window.parent.frames["UpperWindow"].document.getElementsByName("chkSKUChange") || window.parent.frames["UpperWindow"].document.getElementsByName("chkImageChange")
                        || window.parent.frames["UpperWindow"].document.getElementsByName("chkCommodityChange") || window.parent.frames["UpperWindow"].document.getElementsByName("chkOtherChange")
                        || window.parent.frames["UpperWindow"].document.getElementsByName("chkSwChange"))
                {
                    if (window.parent.frames["UpperWindow"].ProgramInput.chkReqChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkSKUChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkImageChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkCategoryBiosChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkCommodityChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkOtherChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked)
                        blnCategory = true;
                    else if (window.parent.frames["UpperWindow"].ProgramInput.chkDocChange.checked)
                        blnCategory = true;
                }
                else {
                    blnCategory = true;
                }
            }
            else {
                blnGEOS = true;
                blnCategory = true;
                blnBusiness = true;
            }
            
            blnSuccess = blnBios;
            if (blnSuccess && (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3"))  //Change Request
            {
                var q = new Date();
                var m = q.getMonth();
                var d = q.getDate();
                var y = q.getFullYear();

                var Today = new Date(y, m, d);

                if (blnCategory == false) {
                    blnSuccess = false;
                    window.alert("You must specify at least one change category.");
                    window.parent.frames["UpperWindow"].ProgramInput.chkOtherChange.focus();
                }
                else if (blnGEOS == false) {
                    blnSuccess = false;
                    window.alert("You must select at least one Region.");
                }                
                else if (blnBusiness == false) {
                    blnSuccess = false;
                    window.alert("You must specify at least impacted business.");
                    window.parent.frames["UpperWindow"].ProgramInput.chkConsumer.focus();
                }
                
                else if (window.parent.frames["UpperWindow"].txtRecordLocked.value == "1") {
                    blnSuccess = false;
                    window.alert("You are not allowed to edit this DCR.");
                }
                
                else if (isDate(window.parent.frames["UpperWindow"].ProgramInput.txtAvailDate.value) == false && window.parent.frames["UpperWindow"].ProgramInput.txtAvailDate.value != "") {
                    blnSuccess = false;
                    window.alert("You must enter a valid date format if you enter a date.");
                    window.parent.frames["UpperWindow"].ProgramInput.txtAvailDate.select();
                    window.parent.frames["UpperWindow"].ProgramInput.txtAvailDate.focus();
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.chkZsrpRequired.checked) {
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.value == "") {
                        blnSuccess = false;
                        window.alert("You must enter a date for the ZSRP Ready Target Date.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.focus();
                    }
                    else if (!(isDate(window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.value)) && (window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.value != "")) {
                        blnSuccess = false;
                        window.alert("You must enter a valid date format in the date field.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyTargetDt.focus();
                    }
                    else if (!(isDate(window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyActualDt.value)) && (window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyActualDt.value != "")) {
                        blnSuccess = false;
                        window.alert("You must enter a valid date format in the date field.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyActualDt.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtZsrpReadyActualDt.focus();
                    }
                }                
                //check if at least RTP or the EM date is in the DCR when the workflow id = 7
                
                if ((window.parent.frames["UpperWindow"].ProgramInput.hdnWorkflowID.value == "7") &&
                    (window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.value == "") &&
                    (window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.value == ""))
                {
                    blnSuccess = false;
                    window.alert("You must enter either the RTP or the EM date when the DCR has the workflow of 'Change Product Life Cycle Dates'.");
                }

                if (window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.value != "")
                {
                    if (isDate_mmddyyyy(window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.value)) {
                        var RTPDate = new Date(window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.value);                        
                    }
                    else
                    {   // remove FCS from all areas - task 20243
                        blnSuccess = false;
                        window.alert("RTP/MR Date must be in mm/dd/yyyy format.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtRTPDate.focus();
                    }
                }
                if (window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.value != "")
                {
                    if (isDate_mmddyyyy(window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.value)) {
                        var RASDate = new Date(window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.value);
                        if (RASDate < Today) {
                            blnSuccess = false;
                            window.alert("The End of Manufacturing (EM) Date must not be in the past of today's date.");
                            window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.select();
                            window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.focus();
                        }
                    }
                    else
                    {
                        blnSuccess = false;
                        window.alert("End of Manufacturing (EM) Date must be in mm/dd/yyyy format.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtRASDiscoDate.focus();
                    }
                }
            }


            if (blnSuccess && (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5")) {
                window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value = ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtPositiveDescription.value) + String.fromCharCode(1) + ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtNegativeDescription.value);

                if (window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value.length == 1)
                    window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value = "";

                if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtCorrectiveActions.value) == "" && ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtPreventiveActions.value) == "") {
                    window.parent.frames["UpperWindow"].ProgramInput.txtActions.value = "";
                }
                else
                    window.parent.frames["UpperWindow"].ProgramInput.txtActions.value = ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtCorrectiveActions.value) + String.fromCharCode(1) + ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtPreventiveActions.value);
            }


            if (blnSuccess) {                
                if (window.parent.frames["UpperWindow"].ProgramInput.cboApproverStatus) {
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtSaveApproval.value != "0" && window.parent.frames["UpperWindow"].ProgramInput.cboApproverStatus.selectedIndex == 2 && window.parent.frames["UpperWindow"].ProgramInput.txtApproverComments.value == "") {
                        blnSuccess = false;
                        window.alert("You must enter comments if you disapprove.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtApproverComments.select();
                        window.parent.frames["UpperWindow"].ProgramInput.txtApproverComments.focus();
                    }
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.txtNotify.value.length > 8000) {
                    blnSuccess = false;
                    window.alert("The Notify on Approval field can not exceed 8000 characters.");
                    window.parent.frames["UpperWindow"].ProgramInput.txtNotify.select();
                    window.parent.frames["UpperWindow"].ProgramInput.txtNotify.focus();
                }
                else if ((!VerifyEmail(window.parent.frames["UpperWindow"].ProgramInput.txtNotify.value)) && (window.parent.frames["UpperWindow"].ProgramInput.txtNotify.value != "")) {
                    blnSuccess = false;
                    window.alert("You must enter a valid SMTP email address or clear the Notify field.");
                    window.parent.frames["UpperWindow"].ProgramInput.txtNotify.select();
                    window.parent.frames["UpperWindow"].ProgramInput.txtNotify.focus();
                }
                else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtSummary.value) == "") {
                    blnSuccess = false;
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5")
                        window.alert("You must enter an Issue/Accomplishment.");
                    else
                        window.alert("You must enter a summary.");
                    window.parent.frames["UpperWindow"].ProgramInput.txtSummary.focus();
                }
                else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtSummary.value).match(/[^\x00-\x7F\r\n]/)) {
                    blnSuccess = false;
                    window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtSummary.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                    window.parent.frames["UpperWindow"].ProgramInput.txtSummary.focus();
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.cboOwner.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.chkPreinstallDeliverable.checked) {
                    blnSuccess = false;
                    window.alert("You must select an owner if no product is selected.");
                    window.parent.frames["UpperWindow"].ProgramInput.cboOwner.focus();
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.lstPriority.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5") {
                    blnSuccess = false;
                    window.alert("You must select an Impact.");
                    window.parent.frames["UpperWindow"].ProgramInput.lstPriority.focus();
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.cboMetricImpact.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5") {
                    blnSuccess = false;
                    window.alert("You must select the Impacted Metric.");
                    window.parent.frames["UpperWindow"].ProgramInput.cboMetricImpact.focus();
                }
                else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value) == "" && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5") {
                    blnSuccess = false;
                    window.alert("You must describe the impact of this opportunity.");
                    window.parent.frames["UpperWindow"].ProgramInput.txtPositiveDescription.focus();
                }
                else if (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3") {
                    if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value) == window.parent.frames["UpperWindow"].ProgramInput.txtDescriptionTemplate.value || ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value) == "") {
                        blnSuccess = false;
                        window.alert("You must enter a Description for this change request.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtDescription.focus();
                    }
                    else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value).match(/[^\x00-\x7F\r\n]/)) {
                        blnSuccess = false;
                        window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                        window.parent.frames["UpperWindow"].ProgramInput.txtDescription.focus();
                    }
                }

                //Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234 - If Type 3, Check a Product has been selected from checkbox table: ---    

                if (window.parent.frames["UpperWindow"].ProgramInput.txtID.value == "" && blnSuccess) //Adding
                {
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtType.value != "3"){
                        //Validate Product Drop-Down List: ---
                        if ((window.parent.frames["UpperWindow"].ProgramInput.lstProducts.value == "" && !window.parent.frames["UpperWindow"].ProgramInput.lstProducts.disabled) && (!window.parent.frames["UpperWindow"].ProgramInput.chkPreinstallDeliverable.checked)) {
                            blnSuccess = false;
                            window.alert("You must select at least one program.");
                            window.parent.frames["UpperWindow"].ProgramInput.lstProducts.focus();
                        }
                    } else if(window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3"){
                        //Validate Product Checkbox Table: ---
                        var oCheckboxes = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");                        
                        var bPreinstallDeliverableChecked = window.parent.frames["UpperWindow"].ProgramInput.chkPreinstallDeliverable.checked;
                        var bChecked = false;
                        var bDisabled = false;
                        
                        //Are checkboxes disabled: ---
                        for (var i = 0; i < oCheckboxes.length; i++) {
                            if(oCheckboxes[i].disabled === true){
                                bDisabled = true;
                                break;
                            }
                        }
                        //Are checkboxes checked: ---
                        var bRChecked;
                        var ProductNames = "";
                        for (var i = 0; i < oCheckboxes.length; i++) {
                            bRChecked = false;
                            if (oCheckboxes[i].checked === true) {
                                var oRCheckboxes = window.parent.frames["UpperWindow"].document.getElementsByName("chkRelease_" + oCheckboxes[i].value);
                                var ReleaseCount = 0;                                
                               
                                for (var x = 0; x < oRCheckboxes.length; x++) {
                                    if (oRCheckboxes[x].type == "checkbox") {
                                        ReleaseCount++;
                                        if (oRCheckboxes[x].checked === true) {
                                            bRChecked = true;
                                            break;
                                        }
                                    }
                                }

                                if (bRChecked === false && ReleaseCount > 0) {
                                    if (ProductNames != "")
                                        ProductNames = ProductNames + ", ";

                                    ProductNames = ProductNames + $("#product_" + oCheckboxes[i].value, window.parent.frames["UpperWindow"].document).attr("data-productname");
                                }                               

                                bChecked = true;
                            }
                        }

                        //If PreinstallDeliverable NOT checked AND checkboxes are enabled BUT none are checked, display error message: ---
                        if (bPreinstallDeliverableChecked === false && (bDisabled === false && bChecked === false)){
                            blnSuccess = false;
                            window.alert("You must select at least one product.");
                        }
                        else {
                            if (bPreinstallDeliverableChecked === false && (bDisabled === false && ProductNames != ""))
                            {
                                blnSuccess = false;
                                window.alert("You must select at least one release for each selected product:  " + ProductNames + ".");
                            }
                        }
            
                        //Validate Justification: ---
                        if (!window.parent.frames["UpperWindow"].ProgramInput.chkIDChange.checked
                            && window.parent.frames["UpperWindow"].ProgramInput.hidCurrentUserPartner.value == "1")
                        {
                            if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value) == window.parent.frames["UpperWindow"].ProgramInput.txtJustificationTemplate.value
                                || ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value) == "") {
                                blnSuccess = false;
                                window.alert("You must enter a justification for this change request.");
                                window.parent.frames["UpperWindow"].ProgramInput.txtJustification.focus();
                            }
                            else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value).match(/[^\x00-\x7F\r\n]/)) {
                                blnSuccess = false;
                                window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                                window.parent.frames["UpperWindow"].ProgramInput.txtJustification.focus();
                            }
                        }
                    }
                }
                else if (blnSuccess)//editing
                {
                    if (!(window.parent.frames["UpperWindow"].ProgramInput.chkSwChange.checked)) {
                        if (window.parent.frames["UpperWindow"].ProgramInput.lstProducts.selectedIndex == 0) {
                            blnSuccess = false;
                            window.alert("You must select a program.");
                            window.parent.frames["UpperWindow"].ProgramInput.lstProducts.focus();
                        }
                    }
                    if (window.parent.frames["UpperWindow"].ProgramInput.cboOwner.selectedIndex == 0) {
                        blnSuccess = false;
                        window.alert("You must select an owner.");
                        window.parent.frames["UpperWindow"].ProgramInput.cboOwner.focus();
                    }
                    else if (window.parent.frames["UpperWindow"].ProgramInput.cboCoreTeam.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.txtType.value != "5") {
                        blnSuccess = false;
                        window.alert("You must select a Core Team Representative.");
                        window.parent.frames["UpperWindow"].ProgramInput.cboCoreTeam.focus();
                    }
                    else if (window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 0) {
                        blnSuccess = false;
                        window.alert("You must select a status.");
                        window.parent.frames["UpperWindow"].ProgramInput.cboStatus.focus();
                    }
                    else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value).match(/[^\x00-\x7F\r\n]/)) {
                        blnSuccess = false;
                        window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                        window.parent.frames["UpperWindow"].ProgramInput.txtDescription.focus();
                    }
                    else if ((window.parent.frames["UpperWindow"].ProgramInput.hidCurrentUserPartner.value == "1") && (window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 2 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 4 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 5) && (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value) == "") && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5") {
                        blnSuccess = false;
                        window.alert("Root Cause is required for items that are closed.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtJustification.focus();
                    }
                    else if (window.parent.frames["UpperWindow"].ProgramInput.hidCurrentUserPartner.value == "1" && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "3" && ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value) == "" && !window.parent.frames["UpperWindow"].ProgramInput.chkIDChange.checked) {
                        blnSuccess = false;
                        window.alert("You must enter a justification for this change request.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtJustification.focus();
                    }
                    else if (window.parent.frames["UpperWindow"].ProgramInput.txtJustification != undefined) {
                        if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value).match(/[^\x00-\x7F\r\n]/)) {
                            blnSuccess = false;
                            window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtJustification.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                            window.parent.frames["UpperWindow"].ProgramInput.txtJustification.focus();
                        }                       
                    }
                    else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtDetails.value).match(/[^\x00-\x7F\r\n]/)) {
                        blnSuccess = false;
                        window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtDetails.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                        window.parent.frames["UpperWindow"].ProgramInput.txtDetails.focus();
                    }
                    else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtResolution.value).match(/[^\x00-\x7F\r\n]/)) {
                        blnSuccess = false;
                        window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.txtResolution.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                        window.parent.frames["UpperWindow"].ProgramInput.txtResolution.focus();
                    }
                    else if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.Textarea5.value).match(/[^\x00-\x7F\r\n]/)) {
                        blnSuccess = false;
                        window.alert("Invalid characters detected(location marked with _):\n." + window.parent.frames["UpperWindow"].ProgramInput.Textarea5.value.replace(/[^\x00-\x7F\r\n]/g, '_'));
                        window.parent.frames["UpperWindow"].ProgramInput.Textarea5.focus();
                    }
                    else if ((window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 2 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 4 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 5) && (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtActions.value) == "") && window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "5") {
                        blnSuccess = false;
                        if (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtCorrectiveActions.value) == "") {
                            window.parent.frames["UpperWindow"].ProgramInput.txtCorrectiveActions.focus();
                            window.alert("You must enter a list of corrective actions required for this item.");
                        }
                        else {
                            window.parent.frames["UpperWindow"].ProgramInput.txtPreventiveActions.focus();
                            window.alert("You must enter a list of preventive actions required for this item.");
                        }
                    }
                    else if ((window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 2 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 4 || window.parent.frames["UpperWindow"].ProgramInput.cboStatus.value == 5) && (ltrim(window.parent.frames["UpperWindow"].ProgramInput.txtResolution.value) == "") && window.parent.frames["UpperWindow"].ProgramInput.txtType.value != "4") {
                        blnSuccess = false;
                        window.alert("Resolution is required for items that are closed.");
                        window.parent.frames["UpperWindow"].ProgramInput.txtResolution.focus();
                    }
                }

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
                    if (window.parent.frames["UpperWindow"].document.all("chkDelete" + i).checked) {
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
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
            var iframeName = parent.window.name;
            if (iframeName != '') {
                    parent.window.parent.ClosePropertiesDialog();
            }else if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel();
            } else {                
                    window.parent.close();
                }
            }
        }

        function cmdClear_onclick() {
            window.parent.frames["UpperWindow"].ProgramInput.reset();
            window.parent.frames["UpperWindow"].ProgramInput.hdnClearProjectList.click();
            if (window.parent.frames["UpperWindow"].ProgramInput.hidCurrentUserPartner.value == "1") {
                window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.fontStyle = "italic";
                window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.color = "blue";
            }
        }

        function cmdSubmit_onclick() {
            if (VerifySave()) {
                //cmdEditCancel.value = "Close";
                if ($.urlParam("Layout") != "pulsar2") {
                    cmdEditCancel.disabled = true;
                }
                cmdSubmit.disabled = true;
                cmdClear.disabled = true;
                //cmdSubmit.style.display = "none";
                //cmdClear.style.display = "none";
                if (window.parent.frames["UpperWindow"].ProgramInput.txtType.value == "4")
                    window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value = window.parent.frames["UpperWindow"].frames.myEditor.document.body.innerHTML;
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
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")

    dim isDcrOwner
    dim isTopPM
    dim isApprover
    dim isODM
    dim isPC
    dim isSM
    dim isSubmitter
    dim intDcrStatus
    dim isStatusInvestigating
    dim bolTypeIsDCR '''BCR,SCR,ICR

    bolTypeIsDCR = false
    if TypeID = "3" then 
        bolTypeIsDCR = true
    end if

    dim disableOkButton

    isDcrOwner = false
    isTopPM = false
    isApprover = false
    isODM = false
    isPC = false
    isSM = false
    isStatusInvestigating = false

    disableOkButton = ""   ' or "disabled"

 
  	dim cn
	dim cm
	dim p

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.CommandTimeout = 180
	cn.Open
	set cm = server.CreateObject("ADODB.Command")

	Set cm.ActiveConnection = cn

	set rs = server.CreateObject("ADODB.recordset")

	Dim CurrentUser	
	Dim CurrentUserID
	dim CurrentUserGroup
	dim CurrentUserSysAdmin

	CurrentUserSysAdmin = false
	
	CurrentUserID = 0

	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))
	
	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenStatic
	'rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

    dim CurrentCMImpersonate
	dim CurrentPCImpersonate
	dim CurrentPhWebImpersonate
	dim CurrentMarketingImpersonate
    dim CurrentPMImpersonate

	set cm=nothing	
	if not (rs.EOF and rs.BOF) then
	    CurrentCMImpersonate = rs("CMImpersonate") 
		CurrentPCImpersonate = rs("PCImpersonate") 
		CurrentPhWebImpersonate = rs("PhWebImpersonate") 
		CurrentMarketingImpersonate = rs("MarketingImpersonate") 
        CurrentPMImpersonate = rs("PMImpersonate") 
		CurrentUserID = rs("ID")
		CurrentUserSysAdmin = rs("SystemAdmin")
		CurrentUserGroup = rs("WorkgroupID")
        CurrentUserPartner = rs("PartnerID") 
	end if
	rs.Close

    dim strImpersonateID
    strImpersonateID = 0

    if CurrentCMImpersonate <> 0 then
	    strImpersonateID = CurrentCMImpersonate
    end if

    if CurrentPCImpersonate <> 0 then
	    strImpersonateID = CurrentPCImpersonate
    end if

    if CurrentPhWebImpersonate <> 0 then
	    strImpersonateID = CurrentPhWebImpersonate
    end if

    if CurrentMarketingImpersonate <> 0 then
	    strImpersonateID = CurrentMarketingImpersonate
    end if

    if CurrentPMImpersonate <> 0 then
	    strImpersonateID = CurrentPMImpersonate
    end if

	if (strImpersonateID <> 0) then
        rs.Open "spGetEmployeeByID " & strImpersonateID,cn,adOpenStatic
        if not(rs.EOF and rs.BOF) then
            CurrentUserID = trim(strImpersonateID)
            CurrentUserPartner = rs("Partnerid") 
        end if
        rs.Close        
    end if 

    if CurrentUserPartner <> 1 then
        isODM = true
    end if


	if IssueID <> "" and bolTypeIsDCR then

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetActionProperties"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = IssueID
		cm.Parameters.Append p
	

		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		Dim strID 
		Dim strSubmitterID
		Dim strOwnerId
		Dim strPMID
		Dim strPCID 
		Dim strSMID
		Dim strCommodityManagerID 
		Dim PreinstallOwnerID 
		Dim strTDCCMID 
		Dim PVID 

        if not(rs.EOF and rs.BOF) then
			strID = IssueID & ""
			strSubmitterID = rs("SubmitterID") & ""
			strOwnerId = rs("OwnerID") & ""
			strPMID = rs("PMID") & ""
			strPCID = rs("PCID") & ""
			strSMID = rs("SMID") & ""
			strCommodityManagerID = rs("PDEID") & ""
			PreinstallOwnerID = rs("PreinstallOwnerID") & ""
			strTDCCMID = rs("TdcCmId") & ""
			PVID = rs("ID") & ""
            intDcrStatus = rs("Status")
        end if

		rs.Close

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovals"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = IssueID
		cm.Parameters.Append p
	
		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		do while not rs.EOF
            if trim(CurrentUserID) = trim(rs("ApproverID") & "" ) then
                isApprover = true
            end if
			rs.MoveNext
		loop			
		rs.Close

        disableOkButton = "disabled"

        if trim(CurrentUserID) = trim(strSubmitterID) then
            isSubmitter = true
        end if

        if trim(CurrentUserID) = trim(strOwnerId) then
            isDcrOwner = true
        end if

        if trim(CurrentUserID) = trim(strPMID) or trim(CurrentUserID) = trim(strTDCCMID) then
                isTopPM = true
        end if

        if trim(CurrentUserID) = trim(strPCID) then
                isPC = true
        end if

        if trim(CurrentUserID) = trim(strSMID) then
                isSM = true
        end if

        if intDcrStatus = 6 then
            isStatusInvestigating = true
        end if

        if isODM then
            disableOkButton = "disabled"
        end if

        '''''' Set the "OK" button
        if isDcrOwner or isTopPM or isApprover or isPC or isSM or (isSubmitter and isStatusInvestigating) then
            disableOkButton = ""
        end if

    end if  ''' IssueID <> "" and bolTypeIsDCR 



	Set cm = Nothing	
	cn.Close
	set cn=nothing


    %>
</head>
<body style="border-top: 2px solid #b2b2b2;">

<div style="text-align:right;">
            <%if IssueID <> "" then%>
                <input type="button" value="OK" id="cmdSubmit" name="cmdSubmit" language="javascript" <%=disableOkButton %>
                    onclick="return cmdSubmit_onclick()">
                <input style="display: none" type="button" value="Clear Form" id="cmdClear" name="cmdClear"
                    language="javascript" onclick="return cmdClear_onclick()">
                <input type="button" value="Cancel" id="cmdEditCancel" name="cmdEditCancel" language="javascript"
                    onclick="return cmdEditCancel_onclick()">
            <%else%>
                <input type="button" value="Submit" id="cmdSubmit" name="cmdSubmit" language="javascript"
                    onclick="return cmdSubmit_onclick()">
                <input type="button" value="Clear Form" id="cmdClear" name="cmdClear" language="javascript"
                    onclick="return cmdClear_onclick()">
                <input type="button" value="Cancel" id="cmdEditCancel" name="cmdEditCancel" language="javascript"
                    onclick="return cmdEditCancel_onclick()">
            <%end if%>
            </div>
</body>
</html>