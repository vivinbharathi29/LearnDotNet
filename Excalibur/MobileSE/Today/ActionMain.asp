<%@  language="VBScript" %>
<%
  Option Explicit

  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Dim AppRoot : AppRoot = Session("ApplicationRoot")
' --- GLOBAL & OPTIONAL INCLUDES: ---%>
<!--#INCLUDE FILE="../../includes/oConnect.asp"-->
<!--#INCLUDE FILE="../../includes/orsProduct.asp"-->
<%

' Issue Types
' 1 = Issue
' 2 = Action ItemAdd workflow

' 3 = Change Request
' 4 = Status Note
' 5 = Improvement Opportunity  
' 6 = Test Request
' 7 = Service ECR

    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim TypeID : TypeID = regEx.Replace(Request.QueryString("Type"), "")
    Dim ProdID : ProdID = regEx.Replace(Request.QueryString("ProdID"), "")
    Dim IssueID : IssueID = regEx.Replace(Request.QueryString("ID"), "")
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")	  
    Dim Layout : Layout = Request.QueryString("Layout")

    Dim WorkflowID : WorkflowID = 0


Dim iChangeRequestID
'Harris, Valerie -  02/8/2016 - PBI 15660/ Task 16234 - If Type 3, create unique Change RequestID so products submitted on one change request are displayed toghether in edit mode
If Trim(TypeID) = "3" Then
    If Trim(IssueID) = "" Then
        iChangeRequestID = GetRandomNumber(1, 8000) 
    Else
        iChangeRequestID = GetChangeRequestGroupID(IssueID)
    End If
Else
    iChangeRequestID = 0
End If
%>
<!DOCTYPE html>
<html lang="en" data-browser="" data-version="">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <meta http-equiv="Pragma" content="no-cache" />
    <meta http-equiv="Expires" content="-1" />
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta name="VI60_defaultClientScript" content="JavaScript" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <title>ActionMain</title>
    <link rel="stylesheet" type="text/css" href="../style/programoffice.css" />
    <link rel="stylesheet" type="text/css" href="style/actionmain.css" />    
    <!-- #include file="../../includes/bundleConfig.inc" -->

    <script type="text/javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
    <script type="text/javascript" src="scripts/actionmain.js"></script>
<script id="clientEventHandlersJS" type="text/javascript" >
function chkCategoryBiosChange_onclick() {
    if (ProgramInput.chkCategoryBiosChange.checked) {
        ProgramInput.chkReqChange.disabled = true;
        ProgramInput.chkSKUChange.disabled = true;
        ProgramInput.chkImageChange.disabled = true;
        ProgramInput.chkCommodityChange.disabled = true;
        ProgramInput.chkDocChange.disabled = true;
        ProgramInput.chkOtherChange.disabled = true;
        $("#chkReqChange").prop('checked', false);
        $("#chkSKUChange").prop('checked', false);
        $("#chkImageChange").prop('checked', false);
        $("#chkCommodityChange").prop('checked', false);
        $("#chkDocChange").prop('checked', false);
        $("#chkOtherChange").prop('checked', false);
        $("#Notify").val("");       
    }
    else {
        ProgramInput.chkReqChange.disabled = false;
        ProgramInput.chkSKUChange.disabled = false;
        ProgramInput.chkImageChange.disabled = false;
        ProgramInput.chkCommodityChange.disabled = false;
        ProgramInput.chkDocChange.disabled = false;
        ProgramInput.chkOtherChange.disabled = false;
    } 
 }   

$(function () {
    $("#dialog-charError").dialog({
        height: 350,
        width: 600,
        modal: true,
        autoOpen: false
    });
    var CheckIllegalChars = function (e) {
   	var pastedData;

   	if (e.originalEvent.clipboardData === undefined)//IE
   		pastedData = clipboardData.getData('text');
   	else
   		pastedData = e.originalEvent.clipboardData.getData('text');

   	var clean = pastedData.replace(/[^\x00-\x7F\r\n]/g, '_');//replace any nont printable character

   	if (clean != pastedData) {
        $("#dialog-charError").dialog("open");
        $("#CharErrorMsg").text(clean);
   	}
   };


   var ua = window.navigator.userAgent;
   var msie = ua.indexOf("MSIE ");
   if (msie > 0) // If Internet Explorer
   {
   	$("body").off('paste');
   	$("body").on('paste', function (e) { CheckIllegalChars(e); });
   }
   else
   {
   	$(document).on('paste', function (e) { CheckIllegalChars(e); });
   }

    $("#chkReqChange").click(function () {
    if ($(this).is(':checked')) { ProgramInput.chkCategoryBiosChange.disabled = true; }
    else {
        
        if ($("#chkSKUChange").is(':checked') || $("#chkImageChange").is(':checked') || $("#chkCommodityChange").is(':checked') || $("#chkDocChange").is(':checked') || $("#chkOtherChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }        
    }
    });

    $("#chkSKUChange").click(function () {
    if ($(this).is(':checked')) ProgramInput.chkCategoryBiosChange.disabled = true;
    else {

        if ($("#chkReqChange").is(':checked') || $("#chkImageChange").is(':checked') || $("#chkCommodityChange").is(':checked') || $("#chkDocChange").is(':checked') || $("#chkOtherChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }
    }
});
$("#chkImageChange").click(function () {
    if ($(this).is(':checked')) ProgramInput.chkCategoryBiosChange.disabled = true;
    else {

        if ($("#chkReqChange").is(':checked') || $("#chkSKUChange").is(':checked') || $("#chkCommodityChange").is(':checked') || $("#chkDocChange").is(':checked') || $("#chkOtherChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }
    }
});
    $("#chkCommodityChange").click(function () {        
    if ($(this).is(':checked')) ProgramInput.chkCategoryBiosChange.disabled = true;
    else {

        if ($("#chkReqChange").is(':checked') || $("#chkSKUChange").is(':checked') || $("#chkImageChange").is(':checked') || $("#chkDocChange").is(':checked') || $("#chkOtherChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }
    }
});
$("#chkDocChange").click(function () {
    if ($(this).is(':checked')) ProgramInput.chkCategoryBiosChange.disabled = true;
    else {

        if ($("#chkReqChange").is(':checked') || $("#chkSKUChange").is(':checked') || $("#chkImageChange").is(':checked') || $("#chkCommodityChange").is(':checked') || $("#chkOtherChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }
    }
});
$("#chkOtherChange").click(function () {
    if ($(this).is(':checked')) ProgramInput.chkCategoryBiosChange.disabled = true;
    else {

        if ($("#chkReqChange").is(':checked') || $("#chkSKUChange").is(':checked') || $("#chkImageChange").is(':checked') || $("#chkCommodityChange").is(':checked') || $("#chkDocChange").is(':checked')) {
            ProgramInput.chkCategoryBiosChange.disabled = true;
        }
        else {
            ProgramInput.chkCategoryBiosChange.disabled = false;
        }
    }
});

  });

        function txtJustification_onfocus() {
            ProgramInput.txtJustification.style.fontStyle = "normal";
            ProgramInput.txtJustification.style.color = "black";
            ProgramInput.txtJustification.select();
        }

        function txtJustification_onblur() {
            if (ProgramInput.txtJustification.value == ProgramInput.txtJustificationTemplate.value) {
                ProgramInput.txtJustification.style.fontStyle = "italic";
                ProgramInput.txtJustification.style.color = "blue";
            }
        }

        function txtDescription_onfocus() {
            ProgramInput.txtDescription.style.fontStyle = "normal";
            ProgramInput.txtDescription.style.color = "black";
            ProgramInput.txtDescription.select();
        }

        function txtDescription_onblur() {
            if (ProgramInput.txtDescription.value == ProgramInput.txtDescriptionTemplate.value) {
                ProgramInput.txtDescription.style.fontStyle = "italic";
                ProgramInput.txtDescription.style.color = "blue";
            }
        }


        function cmdAllGeos_onclick() {
            ProgramInput.chkNA.checked = true;
            ProgramInput.chkLA.checked = true;
            ProgramInput.chkAPJ.checked = true;
            ProgramInput.chkEMEA.checked = true;
        }

        function cmdAllBus_onclick() {
            ProgramInput.chkConsumer.checked = true;
            ProgramInput.chkCommercial.checked = true;
            ProgramInput.chkSMB.checked = true;
        }

        function cboStatus_onchange() {
            var SelectedItem;
            SelectedItem = ProgramInput.cboStatus.options[ProgramInput.cboStatus.selectedIndex].text;
            if (SelectedItem == "Approved" || SelectedItem == "Disapproved" || SelectedItem == "Closed") {
                RequireResolution.style.display = "";
                if (ProgramInput.txtType.value == "5") {
                    RequireJustification.style.display = "";
                    RequireActions.style.display = "";
                }
            }
            else {
                RequireResolution.style.display = "none";
                if (ProgramInput.txtType.value == "5") {
                    RequireJustification.style.display = "none";
                    RequireActions.style.display = "none";
                }
            }
        }

        function cmdDate_onclick(target) {
            var strID;
            var txtDateField = document.getElementById(target);
            strID = window.showModalDialog("calDraw1.asp", txtDateField.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                txtDateField.value = strID;
            }
        }

        function cmdAvailDate_onclick() {
            var strID;
            strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtAvailDate.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                ProgramInput.txtAvailDate.value = strID;
            }
        }

        function cmdRTPDate_onclick() {
            var strID;
            strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtRTPDate.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                ProgramInput.txtRTPDate.value = strID;
            }
        }

        function cmdRASDiscoDate_onclick() {
            var strID;
            strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtRASDiscoDate.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                ProgramInput.txtRASDiscoDate.value = strID;
            }
        }
        
        function cmdOwnerAdd_onclick() {
            ChooseEmployee(ProgramInput.cboOwner);
        }


        function ChooseEmployee(myControl) {
            modalDialog.open({ dialogTitle: 'Select Employee', dialogURL: 'ChooseEmployee.asp', dialogHeight: 200, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
            globalVariable.save('1', 'employee_type');
            globalVariable.save(myControl.id, 'employee_dropdown');

            /*var ResultArray;
            ResultArray = window.showModalDialog("ChooseEmployee.asp", "", "dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")

            if (typeof (ResultArray) != "undefined") {
                if (ResultArray[0] != 0) {
                    myControl.options[myControl.length] = new Option(ResultArray[1], ResultArray[0]);
                    myControl.selectedIndex = myControl.length - 1;
                }
            }*/
        }

        function ChooseEmployeeResult() {
            ResultArray = modalDialog.getArgument('employee_query_array');
            ResultArray = JSON.parse(ResultArray);

            iType = globalVariable.get('employee_type');
            sSelectID = globalVariable.get('employee_dropdown');

            switch (iType) {
                case "1":
                    myControl = document.getElementById('' + sSelectID + '');
                    if (typeof (ResultArray) != "undefined") {
                        if (ResultArray[0] != 0) {
                            myControl.options[myControl.length] = new Option(ResultArray[1], ResultArray[0]);
                            myControl.selectedIndex = myControl.length - 1;
                        }
                    }
                    break;
                default: break;
            }
        }

        function lblStatus_onclick() {
            if (ProgramInput.chkReports.checked)
                ProgramInput.chkReports.checked = false;
            else
                ProgramInput.chkReports.checked = true;
        }

        function lblStatus_onmouseover() {
            window.event.srcElement.parentElement.style.cursor = "hand";

        }

        function DeleteApprover() {
            var i;
            var strAdding = "";

            for (i = parseInt(ProgramInput.txtApproversLoaded.value) + 1; i < ApproverTable.rows.length - 1; i++)
                if (document.all("chkDelete" + i).checked)
                document.all("Row" + i).style.display = "none";
            else
                strAdding = strAdding + document.all("cboApprover" + i).value + ",";
        }

        function AddApprover() {
            var NewRow;
            var NewCell;
            DeleteCell.style.display = "";
            for (i = 0; i < ProgramInput.txtApproversLoaded.value; i++)
                document.all("Del" + i).style.display = "";

            NewRow = ApproverTable.insertRow(ApproverTable.rows.length - 1);
            NewRow.name = "Row" + (ApproverTable.rows.length - 2);
            NewRow.id = "Row" + (ApproverTable.rows.length - 2);
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "<INPUT type=\"checkbox\" id=\"chkDelete" + (ApproverTable.rows.length - 2) + "\" name=\"chkDelete" + (ApproverTable.rows.length - 2) + "\">";
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "<SELECT size=1 id=cboApprover" + (ApproverTable.rows.length - 2) + " name=cboApprover" + (ApproverTable.rows.length - 2) + "  onkeypress=\"return combo_onkeypress()\" onfocus=\"return combo_onfocus()\" onclick=\"return combo_onclick()\" onkeydown=\"return combo_onkeydown()\">" + cboEmployee.innerHTML + "</SELECT>&nbsp;<INPUT type=\"button\" value=\"Add\" id=button1 name=button1 onclick=\"return ChooseEmployee(ProgramInput." + "cboApprover" + (ApproverTable.rows.length - 2) + ");\">";
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "<font face=verdana size=1>Approval Required</font>";
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "&nbsp;";
            window.document.all("cboApprover" + (ApproverTable.rows.length - 2)).focus();
        }

        function cboSubmitter_onchange() {
            //	window.alert(ProgramInput.txtSubmitter.value + ":" + ProgramInput.cboSubmitter.value);
            ProgramInput.txtSubmitter.value = ProgramInput.cboSubmitter.value;
        }

        function DeleteItem(strID) {
            alert("This function is temporarily disabled.  Contact Dave Whorton for assistance");
            /*
            var rc;
            if (window.confirm("Are you sure?")) {
                rc = window.showModalDialog("DeleteAction.asp?xxI1Iu4uT9Tg6gR2R=" + strID, "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                if (typeof (rc) != "undefined") {
                    if (rc == "1") {
                        window.returnValue = 1;
                        window.close();
                    }
                }
            }*/
        }

        function ChangeStatus(strID) {

            var rc;

            if (document.all("Status" + strID).innerHTML == "Cancelled") {
                if (confirm("Are you sure you want to reset this status to Requested?")) {
                    rc = window.showModalDialog("ApproverStatus.asp?ActionID=" + ProgramInput.txtID.value + "&ID=" + strID + "&Status=1", "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                    if (typeof (rc) != "undefined") {
                        if (rc == "1")
                            document.all("Status" + strID).innerHTML = "Approval Requested";
                        else
                            window.alert("Unable to update status.");
                    }
                }
            }
            else {
                if (confirm("Are you sure you want to cancel this approval request?")) {
                    rc = window.showModalDialog("ApproverStatus.asp?ActionID=" + ProgramInput.txtID.value + "&ID=" + strID + "&Status=4", "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                    if (typeof (rc) != "undefined") {
                        if (rc == "1")
                            document.all("Status" + strID).innerHTML = "Cancelled";
                        else
                            window.alert("Unable to update status.");
                    }
                }
            }
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

                //for (i=0;i<event.srcElement.length;i++)
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

        function chkPreinstallDeliverable_onclick() {
            if (ProgramInput.chkPreinstallDeliverable.checked) {
                if(ProgramInput.txtType.value != 3){
                ProgramInput.lstProducts.style.display = "none";
                multiselect.style.display = "none";
                }else{
                    //Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234 - Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
                    document.getElementById('divProductRelease').style.display = "none";
                }
            }
            else {
                if(ProgramInput.txtType.value != 3){
                ProgramInput.lstProducts.style.display = "";
                multiselect.style.display = "";
                }else{
                    //Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234 - Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
                    document.getElementById('divProductRelease').style.display = "";
                }
            }
        }

        function PDtext_onclick() {
            if (ProgramInput.chkPreinstallDeliverable.checked) {
                ProgramInput.chkPreinstallDeliverable.checked = false;
                chkPreinstallDeliverable_onclick();
            }
            else {
                ProgramInput.chkPreinstallDeliverable.checked = true;
                chkPreinstallDeliverable_onclick();
            }
        }

        function PDtext_onmouseover() {
            window.event.srcElement.style.cursor = "hand";
        }

        function AddApproverList() {
            var i;
            var NewCombo;

            if (cboProdApprovalList.length == 0)
                window.alert("No Product Approver List exists for " + txtProductName.value + ". You must add Approvers and click the \"Save Product Approver List\" link to create this list.");
            else {
                for (i = 0; i < cboProdApprovalList.length; i++) {
                    AddApprover();
                    NewCombo = window.document.all("cboApprover" + (ApproverTable.rows.length - 2))
                    for (j = 0; j < NewCombo.length; j++) {
                        if (NewCombo.options[j].value == cboProdApprovalList.options[i].text)
                            NewCombo.selectedIndex = j;
                    }
                }
            }
        }
        var ComboOptions;

        function UpdateApproverList() {
            var i;
            var strOut = "";
            ComboOptions = "<SELECT style=\"Display:none\" id=cboProdApprovalList name=cboProdApprovalList>";

            if (window.confirm("Are you sure you want to save the current Approver List as the default Approver list for " + txtProductName.value + "?")) {
                for (i = 1; i < ApproverTable.rows.length - 1; i++) {
                    if (!window.document.all("chkDelete" + i).checked) {
                        strOut = strOut + "," + window.document.all("cboApprover" + i).value;
                        ComboOptions = ComboOptions + "<option>" + window.document.all("cboApprover" + i).value + "</option>"
                    }
                }

                if (strOut == "")
                    spnAddListLink.style.display = "none";
                else
                    spnAddListLink.style.display = "";


                ComboOptions = ComboOptions + "</select>";

                jsrsExecute("ActionRSupdate.asp", myCallback, "ProductApprovers", Array(txtProductID.value, strOut.substr(1)));
            }
        }

        function myCallback(returnstring) {
            if (returnstring != 1)
                window.alert("Unable to update the Product Approver List.");
            else {
                window.alert("Product Approver List updated.");
                divApproverList.innerHTML = ComboOptions;
            }
        }


        function window_onload() {
            if (document.getElementById("txtRecordLocked")) {
                if (txtRecordLocked.value == "1") {
                    if (typeof (window.parent.frames["LowerWindow"].cmdSubmit) != "undefined") {
                        window.parent.frames["LowerWindow"].cmdSubmit.disabled = true;
                    }
                }
            }

            if (ProgramInput.txtType.value == "4") {
                frames.myEditor.document.body.contentEditable = "True";
                frames.myEditor.document.body.innerHTML = "<font face=verdana size=1>" + ProgramInput.txtDescription.value + "</font>";
                frames.myEditor.focus();
            }

            if (ProgramInput.txtID.value == "") {
                rbProductChange_onclick();
            }

            if ((ProgramInput.chkSwChange.checked) && ((ProgramInput.hidDeliverableRootId.value == "") || (ProgramInput.hidDeliverableRootId.value == "0"))) {
                RowRootDeliverable.style.display = "";
            }

            chkBiosChange_onclick();
            chkSwChange_onclick();
            chkZsrpRequired_onclick();

            if (ProgramInput.chkIDChange.checked) {
                rowOperatingSystem.style.display = "none";
                rowLocalization.style.display = "none";
                RequireJustification.style.display = "none";
            }

            //Add modal dialog code to body tag: ---
            modalDialog.load();

            //add datepicker to date fields
            load_datePicker();
        }

        function ApprovalChange(strID) {
            if (strID == 0) //Approve clicked
                ProgramInput.cboApproverStatus.selectedIndex = 1;
            else if (strID == 2)
                ProgramInput.cboApproverStatus.selectedIndex = 3;
            else
                ProgramInput.cboApproverStatus.selectedIndex = 2;
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

        function chkImageChange_onclick() {
            if (ProgramInput.chkImageChange.checked) {
                if (trim(ProgramInput.txtNotify.value) == "")
                    ProgramInput.txtNotify.value = "cheryl.fragnito@hp.com;"
                else
                    ProgramInput.txtNotify.value = "cheryl.fragnito@hp.com; " + ProgramInput.txtNotify.value;
            }
            else {
                ProgramInput.txtNotify.value = ProgramInput.txtNotify.value.replace("cheryl.fragnito@hp.com;", "");
                ProgramInput.txtNotify.value = ProgramInput.txtNotify.value.replace("cheryl.fragnito@hp.com", "");
            }
        }

        function chkBiosChange_onclick() {
            if ((ProgramInput.chkBiosChange.checked) && (!(ProgramInput.chkSwChange.checked))) {
                document.getElementById("rowBiosChange1").style.display = "";
                document.getElementById("rowBiosChange2").style.display = "";
                document.getElementById("rowBiosChange3").style.display = "";
                document.getElementById("divBiosImpact1").style.display = "none";
                document.getElementById("divBiosImpact2").style.display = "";
                document.getElementById("divBiosRequired1").style.visibility = "visible";

                document.getElementById("rowDistribution").style.display = "none";
                document.getElementById("rowLocalization").style.display = "none";
                document.getElementById("rowOperatingSystem").style.display = "none";
            }
            else {
                document.getElementById("rowBiosChange1").style.display = "none";
                document.getElementById("rowBiosChange2").style.display = "none";
                document.getElementById("rowBiosChange3").style.display = "none";
                document.getElementById("divBiosImpact1").style.display = "";
                document.getElementById("divBiosImpact2").style.display = "none";
                document.getElementById("divBiosRequired1").style.display = "none";
                document.getElementById("rowDistribution").style.display = "";
                document.getElementById("rowLocalization").style.display = "";
                document.getElementById("rowOperatingSystem").style.display = "";

            }
        }

        function chkSwChange_onclick() {
            if ((!(ProgramInput.chkBiosChange.checked)) && (ProgramInput.chkSwChange.checked)) {
                document.getElementById("rowBiosChange1").style.display = "";
                document.getElementById("rowBiosChange2").style.display = "";
                document.getElementById("rowBiosChange3").style.display = "";
                document.getElementById("divBiosImpact1").style.display = "none";
                document.getElementById("divBiosImpact2").style.display = "";
                document.getElementById("divBiosRequired1").style.visibility = "visible";
                document.getElementById("rowDistribution").style.display = "none";
                document.getElementById("rowLocalization").style.display = "none";
                document.getElementById("rowOperatingSystem").style.display = "none";
            }
            else {
                document.getElementById("rowBiosChange1").style.display = "none";
                document.getElementById("rowBiosChange2").style.display = "none";
                document.getElementById("rowBiosChange3").style.display = "none";
                document.getElementById("divBiosImpact1").style.display = "";
                document.getElementById("divBiosImpact2").style.display = "none";
                document.getElementById("divBiosRequired1").style.display = "none";
                document.getElementById("rowDistribution").style.display = "";
                document.getElementById("rowLocalization").style.display = "";
                document.getElementById("rowOperatingSystem").style.display = "";

            }
        }

        function rbBiosChange_onclick() {
            clearChangeCategories();
            ProgramInput.chkBiosChange.checked = true;

            if(document.getElementById("lstProducts")){
            ProgramInput.lstProducts.disabled = true;
            }

            RequireJustification.style.display = "";
            rowZsrp.style.display = "none";
            rowBusiness.style.display = "none";
            rowRegions.style.display = "none";
            rowImpact.style.display = "";
            SampleText.innerHTML = "Samples Available:";
            chkBiosChange_onclick();

            //Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
            if(window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts")){
                var oCheckProject = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");
                for (var i = 0; i < oCheckProject.length; i++) {
                    oCheckProject[i].disabled = true;
                }
            }
        }

        function rbSwChange_onclick() {
            clearChangeCategories();
            ProgramInput.chkSwChange.checked = true;
            
            if(document.getElementById("lstProducts")){
            ProgramInput.lstProducts.disabled = true;
            }

            RowRootDeliverable.style.display = "";
            RequireJustification.style.display = "";
            rowZsrp.style.display = "none";
            rowBusiness.style.display = "none";
            rowRegions.style.display = "none";
            rowImpact.style.display = "";
            SampleText.innerHTML = "Samples Available:";
            chkSwChange_onclick();

            //Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
            if(window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts")){
                var oCheckProject = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");
                for (var i = 0; i < oCheckProject.length; i++) {
                    oCheckProject[i].disabled = true;
                }
            }
        }

        function rbProductChange_onclick() {
            clearChangeCategories();
            chkSwChange_onclick();
            chkBiosChange_onclick();
            
            if (document.getElementById("lstProducts")) {
                ProgramInput.lstProducts.disabled = false;
            }

            RowProgram.style.display = "";
            RowChangeCategory.style.display = "";
            SampleText.innerHTML = "Samples Available:";
            RequireJustification.style.display = "";
            rowZsrp.style.display = "";
            rowBusiness.style.display = "";
            rowRegions.style.display = "";
            rowImpact.style.display = "";

            //Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
            if(window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts")){
                var oCheckProject = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");
                for (var i = 0; i < oCheckProject.length; i++) {
                    oCheckProject[i].disabled = false;
                    if (oCheckProject[i].checked) {
                        // Re-initiate default selected Product to set Program Coordinator email to Notify on Approval list
                        oCheckProject[i].click();
                        oCheckProject[i].click();
                    }
                }
            }
        }


        function rbIDChange_onclick() {
            clearChangeCategories();
            chkSwChange_onclick();
            chkBiosChange_onclick();
            ProgramInput.chkIDChange.checked = true;
           
            if(document.getElementById("lstProducts")){
            ProgramInput.lstProducts.disabled = true;
            }

            RowProgram.style.display = "none";
            rowZsrp.style.display = "none";
            rowBusiness.style.display = "";
            SampleText.innerHTML = "Deadline:";
            rowRegions.style.display = "none";
            rowDistribution.style.display = "none";
            rowLocalization.style.display = "none";
            rowOperatingSystem.style.display = "none";
            rowImpact.style.display = "none";
            RequireJustification.style.display = "none";

            //Hide checkox table if preinstall is checked -- Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
            if(window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts")){
                var oCheckProject = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");
                for (var i = 0; i < oCheckProject.length; i++) {
                    oCheckProject[i].disabled = true;
                }
            }
        }


        function clearChangeCategories() {
            RowChangeCategory.style.display = "none";
            RowProgram.style.display = "none"
            RowRootDeliverable.style.display = "none";

            ProgramInput.hidDeliverableRootId.value = "";
            ProgramInput.txtDeliverableRootName.value = "";

            ProgramInput.chkBiosChange.checked = false;
            ProgramInput.chkSwChange.checked = false;
            ProgramInput.chkIDChange.checked = false;
            ProgramInput.chkReqChange.checked = false;
            ProgramInput.chkSKUChange.checked = false;
            ProgramInput.chkImageChange.checked = false;
            ProgramInput.chkCategoryBiosChange.checked = false;
            ProgramInput.chkCommodityChange.checked = false;
            ProgramInput.chkDocChange.checked = false;
            ProgramInput.chkOtherChange.checked = false;
        }

        function SelectProducts(ProdList) {
            var ProdArray = ProdList.split(",");
            var i;
            var j;
            var UpdateCount = 0;
            var ProdCount = 0;

            for (i = 0; i < ProdArray.length; i++) {
                ProdCount++;
                for (j = 0; j < ProgramInput.lstProducts.length; j++) {
                    if (ProdArray[i] == ProgramInput.lstProducts[j].value) {
                        ProgramInput.lstProducts[j].selected = true;
                        UpdateCount++;
                        break;
                    }
                }
            }
            if (UpdateCount == ProdCount && UpdateCount == 1)
                alert("Automatically selected the product defined for this Product Group.\r\rPlease verify the product list was updated correctly.");
            else if (UpdateCount == ProdCount)
                alert("Automatically selected all " + UpdateCount + " products defined for this Product Group.\r\rPlease verify the product list was updated correctly.");
            else if (ProdCount == 1)
                alert("Unable to find the only product defined for this Product Group.  No products have been selected.");
            else
                alert("Automatically selected only " + UpdateCount + " of the " + ProdCount + " products defined for this Product Group.\r\rPlease verify the product list was updated correctly.");
        }

        function cmdAdd_onclick() {
            modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtNotify.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
            globalVariable.save('txtNotify', 'email_field');
            /*var strResult;
            strResult = window.showModalDialog("../../Email/AddressBook.asp?AddressList=" + ProgramInput.txtNotify.value, "", "dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
            if (typeof (strResult) != "undefined")
                ProgramInput.txtNotify.value = strResult;*/
        }
          
        function cmdFindRoot_onclick() {
            var strResult;
            strResult = window.showModalDialog("../../Common/DeliverableRoot.asp?DeliverableList=" + ProgramInput.hidDeliverableRootId.value, "", "dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
		if (typeof (strResult) != "undefined") {
                if (strResult.indexOf("|") > 0) {
                    var strResultArray = strResult.split("|")
                    ProgramInput.hidDeliverableRootId.value = strResultArray[0];
                    ProgramInput.txtDeliverableRootName.value = strResultArray[1];
                }
            }
        }

        function chkZsrpRequired_onclick() {

            if (ProgramInput.chkZsrpRequired.checked) {
                rowZsrpActual.style.display = "";
                rowZsrpTarget.style.display = "";
                divZsrpRequired.style.visibility = "visible";
            }
            else {
                rowZsrpActual.style.display = "none";
                rowZsrpTarget.style.display = "none";
                divZsrpRequired.style.visibility = "hidden";
            }
        }

        function AddDCRWorkflow(ID, CurrentUserID, PVID) {
            var strID, RTPDate="", EMDate="";
            RTPDate = $('#txtRTPDate').val();
            EMDate = $("#txtRASDiscoDate").val();
            strID = window.parent.showModalDialog("DCRWorkflowFrame.asp?DCRID=" + ID + "&UserID=" + CurrentUserID + "&PVID=" + PVID + "&AddNew=1" + "&RTPDate=" + RTPDate + "&EMDate=" + EMDate , "", "dialogWidth:650px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;")    
        }

        function ViewDCRWorkflowStatus(ID, CurrentUserID, PVID) {
            var strID;
            strID = window.parent.showModalDialog("DCRWorkflowFrame.asp?DCRID=" + ID + "&UserID=" + CurrentUserID + "&PVID=" + PVID + "&AddNew=0", "", "dialogWidth:920px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;")
            //document.location.reload();
        }

        function GroupHeaderClicked(ID) {
            if (document.all("ProgramGroupList" + ID).style.display == "none")
                document.all("ProgramGroupList" + ID).style.display = "";
            else
                document.all("ProgramGroupList" + ID).style.display = "none";
        }

        function hdnClearProjectList_onclick() {
            $('#rbTypeDCR').click();

            var oCheckProject = window.parent.frames["UpperWindow"].document.getElementsByName("chkProducts");
            for (var i = 0; i < oCheckProject.length; i++) {
                oCheckProject[i].checked = false;
                var oCheckRelease = window.parent.frames["UpperWindow"].document.getElementsByName("chkRelease_" + oCheckProject[i].value);
                for (var j = 0; j < oCheckRelease.length; j++) {
                    oCheckRelease[j].checked = false;
                }
            }

            ProgramInput.txtNotify.value = "";
            return;
        }

        function UploadZip(ID) {
            //save ID for return function: ---
            globalVariable.save(ID, 'main_uploadzip_ID');

            var url = "<%=AppRoot %>/PMR/SoftpaqFrame.asp?Page=<%=AppRoot %>/common/fileupload.aspx&Title=Upload DCR Attachments";
            modalDialog.open({ dialogTitle: 'Upload', dialogURL: '' + url + '', dialogHeight: 250, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
        }


        function UploadZip_return(strPath) {
            var ID = globalVariable.get('main_uploadzip_ID');
            if (typeof (strPath) != "undefined") {
                $("#UploadAddLinks" + ID).hide();
                $("#UploadRemoveLinks" + ID).show();
                $("#hplAttachment" + ID).hide();
                $("#UploadPath" + ID).text(strPath.substr(strPath.lastIndexOf("\\") + 1, strPath.length));
                $("#txtAttachmentPath" + ID).val(strPath);
            }
        }

        function RemoveUpload(ID) {
            $("#UploadAddLinks" + ID).show();
            $("#UploadRemoveLinks" + ID).hide();
            $("#UploadPath" + ID).text("");
            $("#txtAttachmentPath" + ID).val("");
        }
    </script>
    <link rel="stylesheet" type="text/css" href="../Style/programoffice.css" />

    <style type="text/css">
A:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
h3
{
    font-family: Verdana;
    font-size: small;
    color: Black;
}
TABLE
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana
}

</style>
</head>
<%
	if trim(TypeID) = "" then
		Response.Write "<body bgcolor=""Ivory""><INPUT type=""hidden"" id=txtRecordLocked name=txtRecordLocked value=""""><FONT face=verdana size=2>Not enough information supplied to display this page.</font>"
		Response.End()
	end if

%>
<body bgcolor="Ivory" onload="return window_onload()">
    <form action="ActionSave.asp?layout=<%=Layout%>" method="post" name="ProgramInput">
        <% 

    function LongName(strName)
	dim FirstName
	dim LastName
	dim GroupName
'		LongName = strName
	if instr(strName,",")>0 then
		FirstName = mid(strName,instr(strName,",")+2)
		LastName = left(strName, instr(strName,",")-1)
		if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
			GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
			FirstName  = left(FirstName,instr(FirstName,"(")-2)
		end if
		if right(Firstname,6) = "&nbsp;" then
			Firstname = left(firstname,len(firstname)-6)
		end if
		LongName = FirstName & "&nbsp;" & LastName & GroupName
	else
		LongName = strName
	end if	
	end function
	
	dim rs 
	dim cn
	dim strproducts
	dim strOwners
	dim DisplayForAdd
	dim DisplayForChangeOnly
	dim JustificationTemplate
	dim strID
	dim strPMID
	dim strPCID
	dim strSummary
	dim strReps
	dim strSubmitter
	dim strSubmitterID
	dim strSubmitted
	dim strTarget
	dim strActual
	dim strNotify
	dim strAction
	dim strJustification
	dim strDescription
	dim strResolution
	dim strProgramID
	dim strOwnerID
	dim strCommercial
	dim strConsumer
	dim strSMB
    dim strNA
    dim strLA
	dim strEMEA
	dim strAPJ
	dim strCoreTeamRep
	dim strStatus
	dim strStatuses	
	dim ClosureLabel
	dim strDisplayReport
	dim strOnlineReports
	dim strReportValue
	dim strApprovals
	dim strStatusText
	dim NoApprovals 'perhaps useless
	dim strSaveApprovals
	dim strApproverComments
	dim strDistribution
	dim strCTODate
	dim strBTODate
	dim DisplayDistribution
	dim BTOYes
	dim CTOYes
	dim BTONo
	dim CTONo
	dim DisplayBTODate
	dim DisplayCTODate 
	dim strAddChange
	dim strModifyChange
	dim strRemoveChange
	dim ApproversLoaded
	dim ApproversPending
	dim DescriptionHeight
	dim strOwner
	dim DescriptionTemplate
	dim DisplayRestore
	dim LanguageList
	dim strPriority
	dim strPriorityOptions
	dim strEditSubmitter
	dim strCustomers
	dim blnSubmitterFound

	dim PreinstallOwnerID
	dim strPreinstallOwnerList
	dim blnProdFound
	dim strAvailableForTest
    dim strTargetApprovalDate
	dim strRTPDate
	dim strRASDiscoDate
	dim strAvailableNotes
	dim strProdApprovalList
	dim ProdApprovalArray
	dim strApproverListLink
	dim strProductname
	dim strReqChange
	dim strOtherChange
	dim strDocChange
	dim strSKUChange
	dim strImageChange
	dim strCommodityChange
	dim strSMID
	dim strECNDate
	dim strRecordLocked
	dim SustainingProduct
	dim strNetAffect
	dim strOSList
	dim strCommodityManagerID
	dim strDetails
	dim strStatusValue
	dim strBiosChange
	dim strSwChange
    dim strIDChange
	dim strDeliverableRootId
	dim strDeliverableRootName
	dim strZsrpRequired
	dim strZsrpReadyTargetDt
	dim strZsrpReadyActualDt
	dim isPC
	dim strTdcCmId
	Dim PVID
	dim strDevCenter
    dim devcenterid
    dim CurrentCMImpersonate
	dim CurrentPCImpersonate
	dim CurrentPhWebImpersonate
	dim CurrentMarketingImpersonate
    dim CurrentPMImpersonate
    Dim sListProducts
    dim isDcrOwner
    dim isTopPM
    dim isODM
    dim isSM
    dim isSustainingTeam  '''Sustaining System Team  
    dim isDcrApprover
    dim strAVRequired
    dim isSubmitter
    dim strQualificationRequired
    dim strImportant
    dim strCategoryBiosChange
    dim strAttachment1
    dim strAttachment2
    dim strAttachment3
    dim strAttachment4
    dim strAttachment5
    dim AttachmentArray

    isDcrOwner = false
    isTopPM = false
    isODM = false
    isSM = false
    isSustainingTeam = false
    isDcrApprover = false
    isSubmitter = false
	
    '--Declare ODM Test Status Variable: --
    dim bODMTestLeadStatus
	
	if trim(TypeID) = "4" then
		DescriptionHeight = 200	
	else
		DescriptionHeight = 120	
	end if
	if trim(IssueID) = "" then
		DisplayForAdd = "none"		
	else
		DisplayForAdd = ""
	end if
	DisplayRestore  = "none"
	if trim(TypeID) = "3"  then 'and IssueID = "" then
		DisplayForChangeOnly = ""
		ClosureLabel = "Target Approval:"
		if IssueID = "" then
			DisplayRestore = ""
		end if
	else
		DisplayForChangeOnly = "none"
		ClosureLabel = "Target Date:"
	end if
	
	
	JustificationTemplate = "DCR Business justification that describes impact to the following whether accepted or rejected:" & vbcrlf & _
							"1. Potential revenue impact and/or impact to unit forecast. (Include timeframes)" & vbcrlf & _
							"2. Customer impact/satisfaction" & vbcrlf & _
							"3. Impact on product commonality" & vbcrlf & _
							"4. Contractual obligations"
	
	if trim(TypeID) = "4" then
		DescriptionTemplate = "<strong> Last Weeks Accomplishments:</strong><ul><li></li><li></li><li></li></ul><strong> Next Weeks Plans:</strong><ul><li></li><li></li><li></li></ul>"  
	else 
        if trim(TypeID) = "3" then
            DescriptionTemplate = "Please include sensitive information such as business justifications, revenue, volume, etc in the 'Justification' section of DCR." & vbcrlf & _
                                  "Do not put in the 'Description' section as external partners may have visibility to this area." & vbcrlf & vbcrlf & _
                                  "If pasting information to the Description or Justification field from a program such as Microsoft Word," & vbcrlf & _
                                  "please paste the text first to Notepad. Then copy and paste from Notepad into these fields to avoid getting " & vbcrlf & _
                                  "unrecognizable characters."
        else
		    DescriptionTemplate = ""
        end if
	end if
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	
	
	Dim CurrentUser	
	Dim CurrentUserID
	dim CurrentUserGroup
	dim CurrentUserSysAdmin
	dim cm
	dim p
	
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

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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
        CurrentUserPartner = trim(rs("PartnerID") & "" )
	end if
	rs.Close
	
	dim strImpersonateName
    strImpersonateName = ""

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
            CurrentUserID = strImpersonateID
            strImpersonateName = longname(rs("name") & "")
            CurrentUserPartner = trim(rs("Partnerid") & "" ) 
        end if
        rs.Close        
    end if 


    '--PBI 5621 - Define ODM Test Status Variable: --
    'bODMTestLeadStatus = IsODMTestLead(CurrentUserID) 
	
	blnProdFound = false	
	strPMID = ""
	strPCID = ""
	strSMID = ""
	
	PreinstallOwnerID = ""
	strPreinstallOwnerList = ""
	strID = ""
	strSummary=""
	strDescription = ""
	strSubmitter = ""
	strSubmitterID = ""
	strSubmitted = ""
	strtarget = ""
	strActual = ""
	strNotify = ""
	strDescription = DescriptionTemplate
	strAction = ""
	strResolution = ""
	strJustification = ""
	strProgramId = 0
	strOwnerID = ""
	strCommercial = ""
	strConsumer = ""
	strSMB = ""
    strNA = ""
    strLA = ""
	strEMEA = ""
	strAPJ = ""
	strCoreTeamRep = ""
	strRecordLocked = "0"
	strStatus = ""
	strStatuses = ""
	strStatusText = ""
	strPMID = ""
	strPCID = ""
	strSMID = ""
	strCommodityManagerID = ""
	strOnlineReports = ""
	strReportValue = ""
	strApprovals = ""
	strSaveApprovals = "0"
	strDistribution = ""
	strCTODate = ""
	strBTODate = ""
	BTOYes = ""
	CTOYes = ""
	BTONo = ""
	CTONo = ""
	DisplayBTODate = "none"
	DisplayCTODate = "none"
	strAddChange = ""
	strModifyChange = ""
	strRemoveChange = ""
	strOwner = ""	
	strPriority = ""
	LanguageList = ""
	strPriorityOptions	= ""
	strEditSubmitter = ""
	strCustomers = ""
	strAvailableForTest = ""
    strTargetApprovalDate = ""
	strRTPDate = ""
	strRASDiscoDate = ""
	strAvailableNotes = ""
	strProdApprovalList = ""
	strApproverListLink = ""
	strProductname= ""
	strReqChange = ""
	strBiosChange = ""
    strIDChange = ""
	strSwChange = ""
	strOtherChange = ""
	strDocChange = ""
	strECNDate = ""
	SustainingProduct = false
	strNetAffect = ""
	strDetails = ""
	strStatusValue = ""
	strDeliverableRootId = ""
	strDeliverableRootName = ""
	strZsrpRequired = ""
	strZsrpReadyTargetDt = ""
	strZsrpReadyActualDt = ""
	isPC = false
	strTdcCmId = ""
	strAVRequired = ""
    strQualificationRequired = ""
    strImportant = ""
    AttachmentArray = ""
    strAttachment1 = ""
    strAttachment2 = ""
    strAttachment3 = ""
    strAttachment4 = ""
    strAttachment5 = ""

	if CategoryID = "1" then
		strSKUChange = "checked"
	else
		strSKUChange = ""
	end if
	strImageChange = ""
	strCommodityChange = ""
    strCategoryBiosChange=""
	
	if IssueID <> "" then

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

		'rs.Open "spgetActionProperties " & IssueID,cn,adOpenForwardOnly
		if rs.eof and rs.bof then
			strID="0"
            response.write "<input id=""rowZsrpActual"" style=""display:none"" type=""checkbox"" />"
            response.write "<input id=""chkZsrpRequired"" style=""display:none"" type=""checkbox"" />"
            response.write "<input id=""chkIDChange"" style=""display:none"" type=""checkbox"" />"
            response.write "<input id=""chkSwChange"" style=""display:none"" type=""checkbox"" />"
            response.write "<input id=""chkBiosChange"" style=""display:none"" type=""checkbox"" />"
            response.write "<input id=""rowBiosChange1"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowBiosChange2"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowBiosChange3"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""divBiosImpact1"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""divBiosImpact2"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""divBiosRequired1"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowDistribution"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowLocalization"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowOperatingSystem"" style=""display:none"" type=""hidden"" />"


		else
			strID = IssueID & ""
			strAction  = rs("Actions") & ""
			strResolution  = rs("Resolution") & ""
			strSummary  = replace(rs("Summary") & "","""","&QUOT;")
			strDescription = rs("Description") & ""
			strSubmitter = rs("Submitter") & ""
			strSubmitterID = rs("SubmitterID") & ""
			strSubmitted = rs("Created") & ""
			strTarget = rs("TargetDate") & ""
			strActual = rs("ActualDate") & ""
			strNotify = rs("Notify") & ""
			strJustification = rs("Justification") & ""
			strProgramId = rs("ProductVersionID") & ""
			strOwnerId = rs("OwnerID") & ""
			strCommercial = rs("Commercial") & ""
			strConsumer = rs("Consumer") & ""
			strSMB = rs("SMB") & ""
            strNA = rs("NA")
            strLA = rs("LA")
			strAPJ = rs("APJ") & ""
			strEMEA = rs("EMEA") & ""
			strReqChange = rs("ReqChange") & ""
			strBiosChange = rs("BiosChange") & ""
			strSwChange = rs("SwChange") & ""
			strIDChange = rs("IDChange") & ""
			strOtherChange = rs("OtherChange") & ""
			strDocChange = rs("DocChange") & ""
			strSKUChange = rs("SKUChange") & ""
			strImageChange = rs("ImageChange") & ""
			strCommodityChange = rs("CommodityChange") & ""
			strCoreTeamRep = rs("CoreTeamRep") & ""
			strStatus  = rs("Status") & ""
			strPMID = rs("PMID") & ""
			strPCID = rs("PCID") & ""
			strSMID = rs("SMID") & ""
			strCommodityManagerID = rs("PDEID") & ""
			strOnlineReports = rs("OnlineReports") & ""
			strReportValue = rs("OnStatusReport") & ""
			strDistribution = rs("Distribution") & ""
			strCTODate = rs("CTODate") & ""
			strBTODate = rs("BTODate") & ""
			strECNDate = rs("ECNDate") & ""
			strAddChange = rs("AddChange") & ""
			strModifyChange = rs("ModifyChange") & ""
			strRemoveChange = rs("RemoveChange") & ""
			strPriority = rs("Priority") & ""
			strCustomers = rs("AffectsCustomers") & ""
			PreinstallOwnerID = rs("PreinstallOwnerID") & ""
			strAvailableForTest = rs("AvailableForTest") & ""
            strTargetApprovalDate = rs("TargetApprovalDate") & ""
			strRTPDate = rs("RTPDate") & ""
			strRASDiscoDate = rs("RASDiscoDate") & ""
			strAvailableNotes = rs("AvailableNotes") & ""
			strDetails = rs("Details") & ""
			strDeliverableRootID = rs("DeliverableRootID") & ""
			strZsrpReadyTargetDt = rs("ZsrpReadyTargetDt") & ""
			strZsrpReadyActualDt = rs("ZsrpReadyActualDt") & ""
			strZsrpRequired = rs("ZsrpRequired") & ""
            strAVRequired = rs("AVRequired") & ""
            strQualificationRequired = rs("QualificationRequired") & ""
			strTDCCMID = rs("TdcCmId") & ""
			PVID = rs("ID") & ""
            strImportant = rs("Important") & ""
			'if Trim(strDeliverableRootID) = "" Then strDeliverableRootID = "0"
            strCategoryBiosChange=rs("CategoryBiosChange") & ""
            strAttachment1 = rs("Attachment1") & ""
            strAttachment2 = rs("Attachment2") & ""
            strAttachment3 = rs("Attachment3") & ""
            strAttachment4 = rs("Attachment4") & ""
            strAttachment5 = rs("Attachment5") & ""

		end if
		rs.Close
        If strPCID <> "" Then 
            isPC = (CLng(CurrentUserID) = CLng(strPCID))
        End If

        if trim(CurrentUserID) = trim(strOwnerId) then
            isDcrOwner = true
        end if

        if trim(CurrentUserID) = trim(strSubmitterID) then
            isSubmitter = true
        end if

        if trim(CurrentUserID) = trim(strPMID) or trim(CurrentUserID) = trim(strTDCCMID) then
            isTopPM = true
        end if        

        if CurrentUserPartner <> "1" then
            isODM = true
        end if

        if trim(CurrentUserID) = trim(strSMID) then
            isSM = true
        end if
    

        if trim(strCoreTeamRep) = "12" then '''is Sustaining System Team
            isSustainingTeam = true
        end if


		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = strProgramId
		cm.Parameters.Append p
	

		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing


'		rs.Open "spGetProductVersion " & strProgramId,cn,adOpenForwardOnly
		if not(rs.EOF and rs.BOF) then
			if rs("Sustaining") = 1 then
				SustainingProduct = true
			end if
            devcenterid=rs("DevCenter")
		end if
		rs.Close

        Dim AddDCRNotificationList
        if strID = "0" then
        AddDCRNotificationList = ""
        else
	    rs.Open "SELECT AddDCRNotificationList FROM ProductVersion WITH (NOLOCK) WHERE ID=" & strProgramId, cn, adOpenStatic
        AddDCRNotificationList = rs("AddDCRNotificationList")
        rs.Close
        end if

		If strZsrpRequired = "" Then
		    strZsrpRequired = ""
		ElseIf strZsrpRequired Then
		    strZsrpRequired = "checked"
		Else
		    strZsrpRequired = ""
		End If

        If strAVRequired = "" Then
		    strAVRequired = ""
		ElseIf strAVRequired Then
		    strAVRequired = "checked"
		Else
		    strAVRequired = ""
		End If

        If strQualificationRequired = "" Then
		    strQualificationRequired = ""
		ElseIf strQualificationRequired Then
		    strQualificationRequired = "checked"
		Else
		    strQualificationRequired = ""
		End If
		
		if strReqChange = "" then
			strReqChange = ""
		elseif strReqChange  then
			strReqChange = "checked"
		else
			strReqChange = ""
		end if
    
        if strBiosChange = "" then
            strBiosChange = ""
        elseif strBiosChange then
            strBiosChange = "checked"
        else
            strBiosChange = ""
        end if
        
        If strSwChange = "" Then
            strSwChange = ""
        ElseIf strSwChange Then
            strSwChange = "checked"
        End If

        
        If strIDChange = "" or strIDChange = "False"  Then
            strIDChange = ""
        ElseIf strIDChange Then
            strIDChange = "checked"
        End If

		if strSKUChange = "" then
			strSKUChange = ""
		elseif strSKUChange  then
			strSKUChange = "checked"
		else
			strSKUChange = ""
		end if

		if strImageChange = "" then
			strImageChange = ""
		elseif strImageChange  then
			strImageChange = "checked"
		else
			strImageChange = ""
		end if

        if strCategoryBiosChange = "" then
			strCategoryBiosChange = ""
		elseif strCategoryBiosChange  then
			strCategoryBiosChange = "checked"
		else
			strCategoryBiosChange = ""
		end if

		if strCommodityChange = "" then
			strCommodityChange = ""
		elseif strCommodityChange  then
			strCommodityChange = "checked"
		else
			strCommodityChange = ""
		end if
		
		
		if strDocChange = "" then
			strDocChange = ""
		elseif strDocChange  then
			strDocChange = "checked"
		else
			strDocChange = ""
		end if		

		if strOtherChange = "" then
			strOtherChange = ""
		elseif strOtherChange  then
			strOtherChange = "checked"
		else
			strOtherChange = ""
		end if

        If strImportant = "" Then
		    strImportant = ""
		ElseIf strImportant Then
		    strImportant = "checked"
		Else
		    strImportant = ""
		End If		
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetApproverList"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = strProgramID
		cm.Parameters.Append p

		rs.CursorType = adOpenStatic
		'rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

        Dim i
		strProdApprovalList = ""
		'rs.Open "spGetApproverList " & strProgramID ,cn,adOpenForwardOnly
		if not(rs.EOF and rs.BOF) then
			ProdApprovalArray = split(rs("ApproverList") & "",",")
			for i = lbound(ProdApprovalArray) to ubound(ProdApprovalArray)
				if trim(ProdApprovalArray(i)) <> "" then
					strProdApprovalList = strProdApprovalList & "<option>" & ProdApprovalArray(i) & "</option>"
				end if
			next
		end if
		rs.Close
        
        Dim strApprovalListLink
		if strProdApprovalList <> "" then
			strApprovalListLink = "<span ID=spnAddListLink><font size=2 color=black face=verdana>&nbsp;|&nbsp;</font><font face=verdana size=1><a href=""javascript:AddApproverList();"">Add Product Approver List</a></font></font></span>"
		else
			strApprovalListLink = "<span style=""Display:none"" ID=spnAddListLink><font size=2 color=black face=verdana>&nbsp;|&nbsp;</font><font face=verdana size=1><a href=""javascript:AddApproverList();"">Add Product Approver List</a></font></font></span>"
		end if
	
		dim ApprovalCount
        dim isApprover
	
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

'		rs.Open "spListApprovals " & IssueID,cn,adOpenForwardOnly
		i=0
		ApprovalCount = 0
		do while not rs.EOF
            
            isApprover = false
            if trim(CurrentUserID) = trim(rs("ApproverID")) then
                isApprover = true 'for one approval item
                isDcrApprover = true 'for the current DCR
            end if
			strStatusText = rs("Status")
			select case strStatusText
			case "1"
				strStatusText = "Approval Requested"
				strStatusValue = "<SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=""1"">Approval Requested</OPTION></SELECT><input type=hidden name=""commentsonly"" value=""true"" />"
			case "2"
				strStatusText = "Approved"
				strStatusValue = "<SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=""2"">Approved</OPTION></SELECT><input type=hidden name=""commentsonly"" value=""true"" />"
			case "3"
				strStatusText = "Disapproved"
				strStatusValue = "<SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=""3"">Disapproved</OPTION></SELECT><input type=hidden name=""commentsonly"" value=""true"" />"
			case "4"
				strStatusText = "Cancelled"
				strStatusValue = "<SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=""4"">Cancelled</OPTION></SELECT><input type=hidden name=""commentsonly"" value=""true"" />"
			case "5"
			    strStatusText = "Not Applicable"
				strStatusValue = "<SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=""5"">Not Applicable</OPTION></SELECT><input type=hidden name=""commentsonly"" value=""true"" />"
			end select

			strApproverComments = rs("Comments")
			if isTopPM or isPC or isSM then
				if isApprover and  rs("Status") = "1"  and ApprovalCount = 0 then
					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><span style=""Display:""><font face=verdana size=1><INPUT type=""radio"" id=optApproval name=optApproval onclick=""return ApprovalChange(0)"">Approve<BR><INPUT type=""radio"" id=optApproval name=optApproval onclick=""return ApprovalChange(1)"">Disapprove<BR><INPUT type=""radio"" id=optApproval name=optApproval onclick=""return ApprovalChange(2)"">Not Applicable<BR></span><SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=1>Approval Requested</OPTION><OPTION value=2>Approved</OPTION><OPTION value=3>Disapproved</OPTION><OPTION value=5>Not Applicable</OPTION></SELECT></td><TD><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
					strSaveApprovals = trim(rs("ID"))
					ApprovalCount = ApprovalCount + 1
					ApproversPending = ApproversPending & rs("ApproverID") & ","
				else
					if  rs("Status") = "1" then
						ApproversPending = ApproversPending & rs("ApproverID") & ","
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & ">" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
    				elseif isApprover and  rs("Status") <> "1" then
    					strSaveApprovals = trim(rs("ID"))
    					'strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap>" & strStatusValue & "<font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & "><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & "><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
                    else
                        strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & ">" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
                    end if
				end if
			else
				if isApprover and rs("Status") = "1" and ApprovalCount = 0 then
					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><SELECT id=cboApproverStatus name=cboApproverStatus><OPTION selected value=1>Approval Requested</OPTION><OPTION value=2>Approved</OPTION><OPTION value=3>Disapproved</OPTION></SELECT></td><TD><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
					strSaveApprovals = trim(rs("ID"))
					ApproversPending = ApproversPending & rs("ApproverID") & ","
					ApprovalCount = ApprovalCount + 1
				else
					if  rs("Status") = "1" then
						ApproversPending = ApproversPending & rs("ApproverID") & ","
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1>" & strStatusText & "</font></td><TD><font face=verdana size=1>" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
    				elseif isApprover and  rs("Status") <> "1" then
    					strSaveApprovals = trim(rs("ID"))
	                	strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap>" & strStatusValue & "<font face=verdana size=1>" & strStatusText & "</font></td><TD><font face=verdana size=1><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
                    else
                        strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1>" & strStatusText & "</font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & ">" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
                    end if
				end if
			end if
			i=i+1
			rs.MoveNext
		loop			
		rs.Close
		
		ApproversLoaded = i
		
		JustificationTemplate = ""

		if len(strApprovals) > 0 then
			if (isTopPM or isPC or isSM or isDcrOwner or isDcrApprover) then
				strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=600 bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD nowrap><font face=verdana size=1><strong>Status</strong></font></TD><TD width=""100%""><font face=verdana size=1><strong>Comments</strong></font></TD></TR>" & strApprovals & "<TR><TD colspan=4><font size=1 face=verdana><a href=""javascript:AddApprover();"">Add Approver</a></font></td></tr></TABLE><BR>"
			else
				strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=600 bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD nowrap><font face=verdana size=1><strong>Status</strong></font></TD><TD width=""100%""><font face=verdana size=1><strong>Comments</strong></font></TD></TR>" & strApprovals & "</TABLE><BR>"
			end if
		end if

		if strDistribution = "Both" then
			BTOYes = "selected"
			CTOYes = "selected"
			DisplayBTODate = ""
			DisplayCTODate = ""
		elseif strDistribution = "BTO" then
			BTOYes = "selected"
			CTONo = "selected"
			DisplayBTODate = ""
			DisplayCTODate = "none"
		elseif strDistribution = "CTO" then
			CTOYes = "selected"
			BTONo = "selected"
			DisplayBTODate = "none"
			DisplayCTODate = ""
		else
			BTOYes = ""
			CTOYes = ""
			BTONo = "selected"
			CTONo = "selected"
			DisplayBTODate = "none"
			DisplayCTODate = "none"
		end if

	

	else
		if TypeID = "3" then
			strJustification = JustificationTemplate
            strDescription = DescriptionTemplate
		else
			strJustification = ""
			JustificationTemplate = ""	
            
            strDescription = ""
            DescriptionTemplate = ""	
		end if
		
		rs.Open "SELECT OptionConfig, Name, MIN(DisplayOrder) AS DisplayOrder FROM Regions with (NOLOCK) WHERE (Active = 1) GROUP BY OptionConfig, Name ORDER BY OptionConfig",cn,adOpenForwardOnly
		LanguageList = ""
		do while not rs.EOF
			LanguageList = LanguageList &  "<OPTION Value=""" & rs("OptionConfig") & """>" & rs("OptionConfig") & " - " & rs("Name") & "</OPTION>"
			rs.movenext
		loop
		rs.Close
		'LanguageList = LanguageList &  "<OPTION Value=""NA=(US,SP,FR)"">NA (US,SP,FR)</OPTION>"
		'LanguageList = LanguageList &  "<OPTION Value=""LA=(SP,BR,US)"">LA (SP,BR,US)</OPTION>"
		'LanguageList = LanguageList &  "<OPTION Value=""CKK=(JP,US)"">CKK (JP,US)</OPTION>"
		'LanguageList = LanguageList &  "<OPTION Value=""APD=(CH,US,TW,TZ,KR)"">APD (CH,US,TW,TZ,KR)</OPTION>"
		'LanguageList = LanguageList &  "<OPTION Value=""EMEA=(AR,NL,FR,CS,DK,US,FI,GR,GK,HU,IL,IT,NO,PL,PT,RU,SP,SE,TR)"">EMEA (AR,NL,FR,CS,DK,US,FI,GR,GK,HU,IL,IT,NO,PL,PT,RU,SP,SE,TR)</OPTION>"
		
		'"<a href=""javascript: SelectProducts('247,340');"">3C05</a><BR>"
        dim  strCycleProductLinks
		strCycleProductLinks = ""

		strCycleProductLinks = strCycleProductLinks & "<div id=""divGroup""> "
		dim strProductIDLinks
        dim strLastProgramGroup
		strProductIDLinks = ""
        strLastProgramGroup = ""
		Dim LastProgram
        Lastprogram=""
		rs.Open "spGetProgramTree",cn,adOpenForwardOnly
		do while not rs.EOF
            if strLastProgramGroup ="" then
                strCycleProductLinks = strCycleProductLinks & "<a href=""javascript:GroupHeaderClicked(" & rs("programGroupID") & ");"" class=""grp-header"">" & rs("ProgramGroup") & "</a><br><div style=""display:none"" id=""ProgramGroupList" & rs("ProgramGroupID") & """><ul style=""margin-top:0px;margin-bottom:0px;"">"
                strLastProgramGroup = rs("ProgramGroup")
            end if
			if LastProgram <> rs("Program") and LastProgram <> "" then
				if len(strProductIDLinks) > 1 then
					strProductIDLinks = mid(strProductIDLinks,2)
				end if
                    '--If Type 3, add SelectMultipleProductCheckbox to li -- Harris, Valerie (2/6/2016) - PBI 15660/ Task 16234: ---
                    If trim(TypeID) = "3" Then
					    strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectMultipleProductCheckbox('" & strProductIDLinks & "', 'Product Group');"">" & LastProgram & "</a></li>"
				    Else
					strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a></li>"
                    End If
            
				strProductIDLinks = ""
			end if
            if strLastProgramGroup <> rs("ProgramGroup") then
                strCycleProductLinks = strCycleProductLinks & "</ul></div><a href=""javascript:GroupHeaderClicked(" & rs("programGroupID") & ");"" class=""grp-header"">" & rs("ProgramGroup") & "</a><br><div style=""display:none"" id=""ProgramGroupList" & rs("ProgramGroupID") & """><ul style=""margin-top:0px;margin-bottom:0px;"">"
                strLastProgramGroup = rs("ProgramGroup")
            end if
			strProductIDLinks = strProductIDLinks & "," & rs("ID")
			LastProgram = rs("Program") & ""
			rs.MoveNext
		loop
		if strLastProgramGroup <> "" and LastProgram <> "" then
			if len(strProductIDLinks) > 1 then
				strProductIDLinks = mid(strProductIDLinks,2)
			end if
            '--If Type 3, add SelectMultipleProductCheckbox to li -- Harris, Valerie (2/6/2016) - PBI 15660/ Task 16234: ---
            If trim(TypeID) = "3" Then
			strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectMultipleProductCheckbox('" & strProductIDLinks & "', 'Product Group');"">" & LastProgram & "</a></li>"
			Else
			strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectProducts('" & strProductIDLinks & "'); SelectMultipleProductCheckbox('" & strProductIDLinks & "', 'Product Group');"">" & LastProgram & "</a></li>"
            End If
			strProductIDLinks = ""
            strCycleProductLinks = strCycleProductLinks & "</ul></div>"
        end if
		rs.Close		
		
		strCycleProductLinks = strCycleProductLinks & "</div> "
		
	end if

if strID="0" then
	Response.Write "<font size=2 face=verdana><BR><BR>Item not found.</font>"
else
		strPriorityOptions = "<option value=0></option>"
		if strPriority = "1" then
			strPriorityOptions = strPriorityOptions & "<option selected value=1>High</option>"
		else
			strPriorityOptions = strPriorityOptions & "<option value=1>High</option>"
		end if 
		if strPriority = "2" then
			strPriorityOptions = strPriorityOptions & "<option selected value=2>Medium</option>"
		else
			strPriorityOptions = strPriorityOptions & "<option value=2>Medium</option>"
		end if 
		if strPriority = "3" then
			strPriorityOptions = strPriorityOptions & "<option selected value=3>Low</option>"
		else
			strPriorityOptions = strPriorityOptions & "<option value=3>Low</option>"
		end if 


	strCommercial = replace(replace(strCommercial,"True","checked"),"False","")
	strConsumer = replace(replace(strConsumer,"True","checked"),"False","")
	strSMB = replace(replace(strSMB,"True","checked"),"False","")

    strNA = replace(replace(strNA,"True","checked"),"False","")
    strLA = replace(replace(strLA,"True","checked"),"False","")
	strEMEA = replace(replace(strEMEA,"True","checked"),"False","")
	strAPJ = replace(replace(strAPJ,"True","checked"),"False","")
	strNetAffect = strCustomers
	strCustomers = replace(strCustomers,"1","checked")
	
	strAddChange = replace(replace(strAddChange,"True","checked"),"False","")
	strModifyChange = replace(replace(strModifyChange,"True","checked"),"False","")
	strRemoveChange = replace(replace(strRemoveChange,"True","checked"),"False","")

	strStatuses = ""

	if TypeID = "3" then
		if strStatus <> "3" and strStatus <> "4" and strStatus <> "5" and strStatus <> "1" then
			strStatuses = "<Option value=0 selected></option>"
		end if
		if strStatus = "1" then
			strStatuses = strStatuses & "<OPTION value=1 selected>Open</OPTION>"
			 strStatusText  = "Open"
		else
			strStatuses = strStatuses & "<OPTION value=1 >Open</OPTION>"
		end if 
		if strStatus = "4" then
			strStatuses = strStatuses & "<OPTION value=4 selected>Approved</OPTION>"
			 strStatusText  = "Approved"
		else
			strStatuses = strStatuses & "<OPTION value=4>Approved</OPTION>"
		end if 
		if strStatus = "5" then
			strStatuses = strStatuses & "<OPTION value=5 selected>Disapproved</OPTION>"
			 strStatusText  = "Disapproved"
		else
			strStatuses = strStatuses & "<OPTION value=5>Disapproved</OPTION>"
		end if 
		if strStatus = "3" then
			strStatuses = strStatuses & "<OPTION value=3 selected>Need More Information</OPTION>"
			 strStatusText  = "Need More Information"
		else
			strStatuses = strStatuses & "<OPTION value=3>Need More Information</OPTION>"
		end if 
		if strStatus = "6" then
			strStatuses = strStatuses & "<OPTION value=6 selected>Investigating</OPTION>"
			 strStatusText  = "Investigating"
		else
			strStatuses = strStatuses & "<OPTION value=6>Investigating</OPTION>"
		end if 
	else
		if strStatus <> "1" and strStatus <> "2" then
			strStatuses = "<Option value=0 selected></option>"
		end if
		if strStatus = "1" then
			strStatuses = strStatuses & "<OPTION value=1 selected>Open</OPTION>"
		else
			strStatuses = strStatuses & "<OPTION value=1>Open</OPTION>"
		end if 
		if strStatus = "2" then
			strStatuses = strStatuses & "<OPTION value=2 selected>Closed</OPTION>"
		else
			strStatuses = strStatuses & "<OPTION value=2>Closed</OPTION>"
		end if 
	end if

		rs.Open "spgetproductsall -3",cn
		strproducts = ""
		strDevCenter = ""
		
		blnProdFound = false
		
		do while not rs.EOF
			if (trim(TypeID) = "3" and rs("AllowDCR") ) or ( trim(TypeID) <> "3" and rs("ProductStatusID") < 5) then
                if rs("DevCenterName") = "" then
                    strproducts = "<optgroup label=""" & rs("DevCenterName") & """>"
                    strDevCenter = rs("DevCenterName")
                elseif rs("DevCenterName") <> strDevCenter then
                    strDevCenter = rs("DevCenterName")
                    strproducts = strproducts & "</optgroup><optgroup label=""" & rs("DevCenterName") & """>"
                end if

				if strProgramID = rs("ID") & "" or ProdID = rs("ID") & "" then
                    strproducts = strproducts &  "<OPTION selected Value=""" & rs("ID") & """>" & rs("Product") & "</OPTION>"
                    strProductname = rs("Product")
					blnprodFound = true
				else
    				strproducts = strproducts &  "<OPTION Value=""" & rs("ID") & """>" & rs("Product") & "</OPTION>"
				end if

			end if
			rs.movenext
		loop
		rs.Close
		strproducts = strproducts & "</optgroup>"
		
		If Trim(strDeliverableRootId) <> "0" And Trim(strDeliverableRootID) <> "" Then
		    rs.Open "spGetDeliverableRootName " & strDeliverableRootID, cn, adOpenStatic
		    If Not rs.Eof And Not rs.Bof Then
		        strDeliverableRootName = rs("Name") & ""
		    Else
		        strDeliverableRootName = "&nbsp;"
		    End If
		    rs.Close
		End If
		
		if not blnProdFound and IssueID <> "" then
			rs.open "spGetProductVersionName " & strProgramID ,cn,adOpenForwardOnly
			if rs.EOF and rs.EOF then
				strproducts = strproducts &  "<OPTION selected Value=""" & strProgramID & """>Product Name Not Found</OPTION>"		
			else
				strproducts = strproducts &  "<OPTION selected Value=""" & strProgramID & """>" & rs("Name") & "</OPTION>"	
				strProductName = rs("Name") 
			end if
			rs.Close
		end if


'	if strProgramID = "170" then
'		strproducts = strproducts &  "<OPTION selected Value=""" & "170" & """>" & "Not Assigned" & "</OPTION>"
'	end if
	
	rs.Open "spListAllActionOwners",cn,adOpenForwardOnly
	'rs.Open "spgetEmployees",cn,adOpenForwardOnly
	
	if (trim(TypeID) = "1" or trim(TypeID) = "2" or trim(TypeID) = "5") and IssueID = "" then
		strOwners = "<option value=0 selected>[Product PM]</option>"	
	else
		strOwners = "<Option value=0 selected></option>"
	end if
	Dim strEmployee
	strEmployee = strOwners
	strEditSubmitter = ""
	blnSubmitterFound = false
	do while not rs.EOF
		if strOwnerID = rs("ID") & "" then
			strOwners = strOwners &  "<OPTION selected Value=""" & rs("ID") & """>" & rs("Name")&  "</OPTION>"
			strOwner = rs("Name")
		else
			strOwners = strOwners &  "<OPTION Value=""" & rs("ID") & """>" & rs("Name")&  "</OPTION>"
		end if

		strEmployee = strEmployee &  "<OPTION Value=""" & rs("ID") & """>" & rs("Name")&  "</OPTION>"

		rs.movenext
	loop
	rs.Close
	
	
	rs.Open "spgetOS",cn,adOpenForwardOnly
	strOSList = ""
	do while not rs.EOF
		if rs("ID") <> 16 then
			strOSList = strOSList &  "<OPTION Value=""" & rs("Name") & """>" &  rs("Name") &  "</OPTION>"
		end if
		rs.movenext
	loop
	rs.Close	
	
	if trim(PreinstallOwnerID) = "-1" then
		strPreinstallOwnerList = "<OPTION selected value=-1>{Database Team}</OPTION>" & strPreinstallOwnerList
	else
		strPreinstallOwnerList = "<OPTION value=-1>{Database Team}</OPTION>" & strPreinstallOwnerList
	end if

	
	
	if not blnsubmitterfound then
			strEditSubmitter = strEditSubmitter & "<Option selected value=""" & strSubmitter & """>" & strSubmitter  & "</Option>"			
	end if
	
	Dim DisplayOwner
	if strOwner = "" and trim(TypeID) = "4" then
		DisplayOwner = "none"
	else
		DisplayOwner = ""
	end if
	
	rs.Open "spgetCoreTeamReps",cn,adOpenForwardOnly
	strReps = "<Option value=0 selected></option>"
	do while not rs.EOF
		if strCoreTeamRep = rs("ID") & "" then
			strReps = strReps &  "<OPTION selected Value=""" & rs("ID") & """>" & replace(replace(rs("Name"),"<",""),">","") &  "</OPTION>"
		else
			strReps = strReps &  "<OPTION Value=""" & rs("ID") & """>" & replace(replace(rs("Name"),"<",""),">","") &  "</OPTION>"
		end if
		rs.movenext
	loop
	rs.Close
	
	dim strTypeDisplay
	DisplayDistribution = ""
	Select case TypeID
	case "1"
		strTypeDisplay = "Issue"
		DisplayDistribution = "none"
	case "3"
		strTypeDisplay = "Change Request"
		If IssueID <> "" Then
		    DisplayDistribution = "none"
		End If
	case "2"
		strTypeDisplay = "Action Item"
		DisplayDistribution = "none"
	case "4"
		strTypeDisplay = "Status Note"
		DisplayDistribution = "none"
	case "5"
		strTypeDisplay = "Improvement Opportunity"
		DisplayDistribution = "none"
	case "6"
		strTypeDisplay = "Test Request"
		DisplayDistribution = "none"
	case else
		strTypeDisplay = "&nbsp;"
		DisplayDistribution = "none"
	end select
	
	if IssueID <> "" then
		if isTopPM or isPC or isSM or CurrentUserSysAdmin  then
		
		    rs.Open "usp_DCRWorkflowCheck " & IssueID,cn,adOpenForwardOnly
	        Dim OpenWorkflowCount
	        if rs.eof and rs.bof then
			    OpenWorkflowCount = 0
                WorkflowID = 0
		    else
		        OpenWorkflowCount = rs("OpenWorkflowCount")
                WorkflowID = rs("WorkFlowID")
		    end if
	        rs.Close
	        
	        if strImpersonateID <> 0 then
	            if TypeID = 3 and OpenWorkflowCount = 0 and (isTopPM or isSM ) then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1><a href=""javascript:AddDCRWorkflow(" & IssueID &  "," & CurrentUserID &  "," & PVID & ");"">Add Workflow</a></TD></TR><TR><TD align=left><font color=red><H5>" & strImpersonateName & " </H5></font></TD></TR></TABLE>" 
		        elseif TypeID = 3 and OpenWorkflowCount > 0 then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1><a href=""javascript:ViewDCRWorkflowStatus(" & IssueID &  "," & CurrentUserID &  "," & PVID & ");"">View Workflow Status</a></TD></TR><TR><TD align=left><font color=red><H5>" & strImpersonateName & " </H5></font></TD></TR></TABLE>" 
                  elseif TypeID = 3 and OpenWorkflowCount = -1 then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1>Workflow&nbsp;Pending&nbsp;Approval</TD></TR><TR><TD align=left><font color=red><H5>" & strImpersonateName & " </H5></font></TD></TR></TABLE>" 
                else
			        Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1></font></TD></TR><TR><TD align=left><font color=red><H5>" & strImpersonateName & " </H5></font></TD></TR></TABLE>" 
		        end if
	        else
		        if TypeID = 3 and OpenWorkflowCount = 0 and (isTopPM or isSM ) then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1><a href=""javascript:AddDCRWorkflow(" & IssueID &  "," & CurrentUserID &  "," & PVID & ");"">Add Workflow</a></font></TD></TR></TABLE>" 
		        elseif TypeID = 3 and OpenWorkflowCount > 0 then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1><a href=""javascript:ViewDCRWorkflowStatus(" & IssueID &  "," & CurrentUserID &  "," & PVID & ");"">View Workflow Status</a></font></TD></TR></TABLE>" 
                  elseif TypeID = 3 and OpenWorkflowCount = -1 then
		            Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1>Workflow&nbsp;Pending&nbsp;Approval</font></TD></TR></TABLE>" 
                else
			        Response.write "<TABLE width=""100%""><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1></font></TD></TR></TABLE>" 
		        end if
		    end if
		else
			Response.write "<H3>" & strTypeDisplay & " Properties</H3>" 
		end if
	else
		if CategoryID="1" then	
			Response.write "<H3>Add New SKU Change Request</H3>" 
		else
			Response.write "<H3>Add New " & strTypeDisplay & "</H3>" 
		end if
	end if
	if (isTopPM or isPC or isSM or CurrentUserSysAdmin) and strOnlineReports = "1"  then
		strDisplayReport = ""
	else
		strDisplayReport = "none"
	end if		
	if trim(strReportValue) = "1" then
		strReportValue = "checked"
	else
		strReportValue = ""
	end if

strRecordLocked = "0"
if (TypeID = "3") and (trim(strStatus) = "4" or trim(strStatus) = "5") then
	if not (isTopPM or isPC or isSM) then '  or CurrentUserSysAdmin) then
		if ((not SustainingProduct) and (not isSustainingTeam) ) or (trim(strECNDate) <> "" or trim(strStatus) = "5") then
			Response.Write "<font size=2 face=verdana color=red><b>This DCR is closed and can only be edited by the PM.</b></font><BR><BR>"
			strRecordLocked = "1"
		end if
	end if
end if


	noApprovals = false
	if ((isTopPM or isPC or isSM or isDcrOwner or isDcrApprover ) and strApprovals = "") and strProgramID <> "170"  then
		strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=100% bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD width=160 nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD><font face=verdana size=1><strong>Status</strong></font></TD><TD><font face=verdana size=1><strong>Comments</strong></font></TD></TR><TR><TD colspan=4><font size=1 face=verdana><a href=""javascript:AddApprover();"">Add Approver</a></font>" & strApprovalListLink & "<font face=verdana size=1>&nbsp;</font></td></tr></TABLE><BR>" 
		
	elseif strApprovals = "" then
		Response.write "<Table style=""Display:none;WIDTH=100%"" ID=ApproverTable></TABLE>"	
	end if	

	if (not noapprovals) and trim(TypeID) <> "4" then
		Response.write strApprovals
	elseif trim(TypeID) = "4" and strID <> "" then
		Response.Write "<Table style=""Display:none"" ID=ApproverTable></TABLE>"
	end if
        %>
        <table border="0" cellpadding="2" cellspacing="0" width="100%" bgcolor="cornsilk"
            bordercolor="tan">

            <tr style="display: none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Preinstall Owner:</font></strong></td>
                <td>
                    <select style="width: 180px;" id="Select7" name="cboPreinstallApprover">
                        <option value="0"></option>
                        <%=strPreinstallOwnerList%>
                    </select>
                </td>
            </tr>

            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">ID:</font></strong></td>
                <td>
                    <font size="2">
                        <%=strID%>
                    </font>
                </td>
            </tr>
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Submitter:</font></strong></td>
                <td>

                    <font size="2" face="verdana">
                        <%=strSubmitter%>
                    </font>
         
                    <input id="txtSubmitter" name="txtSubmitter" type="hidden" value="<%=strSubmitter%>">
                    <input id="txtSubmitterID" name="txtSubmitterID" type="hidden" value="<%=strSubmitterID%>">
                </td>
            </tr>
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Date Submitted:</font></strong></td>
                <td>
                    <font size="2" face="verdana">
                        <%=strSubmitted%>
                    </font>
                </td>
            </tr>
            <%if strStatus = "2" or  strStatus = "4" or  strStatus = "5" then%>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Date&nbsp;Closed:</font></strong></td>
                <td colspan="2" valign="top">
                    <font size="2" face="verdana">
                        <%=strActual%>
                    </font>
                </td>
            </tr>
            <%end if%>
            <%if IssueID <> "" then%>
            <% If Trim(strDeliverableRootID) = "0" Then %>
            <tr>
                <td style="vertical-align: top">
                    <strong><font size="2">Product(s):</font></strong></td>
                <td valign="top">
                    <%
                        If Trim(TypeID) = "3" Then
                            Call BuildChangeRequestGroupList(IssueID, iChangeRequestID) 
                        Else
                    %>
                    <%=strProductname %>
                    <%End If %>
                    <select id="lstProducts" name="lstProducts" style="display: none; width: 180px;">
                        <option value="" selected></option>
                        <%=strproducts%>
                    </select>
                </td>
            </tr>
            <% Else %>
            <tr>
                <td style="vertical-align: top">
                    <strong><font size="2">Deliverable Root:</font></strong></td>
                <td>
                    <%=strDeliverableRootName %>
                </td>
            </tr>
            <% End If %>
            <%end if%>
            <%If IssueID = "" Then %>
            <tr>
                <td style="width: 160px; vertical-align: top; font-weight: bold; font-size: x-small">
                    <font size="2">Change Type:</font><span style="color: Red; font-size:xx-small;">&nbsp;*</span></td>
                <td colspan="2">
                <table border="0" style="font-size:14px">
                <tr><td><input id="rbTypeDCR" name="rbChangeType" type="radio" value="DCR" checked="checked" onclick="rbProductChange_onclick()" /></td><td>Product Change: Request a change to the PDD of one or more products.</td></tr>
                <!--<tr><td><input id="rbTypeBCR" name="rbChangeType" type="radio" value="BCR" onclick="rbBiosChange_onclick();" /></td><td>BIOS Change: Request a change to Core BIOS (BCR).</td></tr>-->
                <tr><td><input id="rbTypeSCR" name="rbChangeType" type="radio" value="SCR" onclick="rbSwChange_onclick();" /></td><td>Software Change: Request a change to a SW Root Deliverable (SCR).</td></tr>
                <tr><td><input id="rbTypeICR" name="rbChangeType" type="radio" value="ICR" onclick="rbIDChange_onclick();" /></td><td>Industrial Design Change: Request a change to the ID Spec (ICR).</td></tr>
                </table>
                </td>
            </tr>
            <%End If %>
            <tr id="RowChangeCategory" style="display:none">
                <td width="160" style="vertical-align: top; white-space:nowrap;">
                    <strong><font size="2">Change Category:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2" nowrap>
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <table border="0">
                            <tr>
                                <td width="114">
                                    <input type="checkbox" id="chkReqChange" name="chkReqChange" <%=strreqchange%>><font
                                        size="2" face="verdana">Requirement</font></td>
                                <td width="60">
                                    <input type="checkbox" id="chkSKUChange" name="chkSKUChange" <%=strskuchange%>><font
                                        size="2" face="verdana">SKU</font></td>
                                <td>
                                    <input type="checkbox" id="chkImageChange" name="chkImageChange" <%=strimagechange%>
                                        onclick="return chkImageChange_onclick()"><font size="2" face="verdana">Software</font></td>
                                <td>
                                    <input type="checkbox" id="chkCategoryBiosChange" name="chkCategoryBiosChange" <%=strCategoryBiosChange%> onclick="return chkCategoryBiosChange_onclick()"><font size="2" face="verdana">BIOS</font></td>
                                <td style="display:none" >
                                    <input type="checkbox" id="chkBiosChange" name="chkBiosChange" <%=strbioschange%>
                                        onclick="return chkBiosChange_onclick()"><font size="2" face="verdana">BCR</font></td>
                                <td style="display:none" >
                                    <input type="checkbox" id="chkIDChange" name="chkIDChange" <%=stridchange%>
                                        onclick="return chkIDChange_onclick()"><font size="2" face="verdana">ICR</font></td>
                            </tr>
                            <tr>
                                <td>
                                    <input type="checkbox" id="chkCommodityChange" name="chkCommodityChange" <%=strcommoditychange%>><font
                                        size="2" face="verdana">Commodity</font></td>
                                <td>
                                    <input type="checkbox" id="chkDocChange" name="chkDocChange" <%=strdocchange%>><font
                                        size="2" face="verdana">Docs</font></td>
                                <td width="60">
                                    <input type="checkbox" id="chkOtherChange" name="chkOtherChange" <%=strotherchange%>><font
                                        size="2" face="verdana">Other</font></td>
                                <td style="display:none">
                                    <input type="checkbox" id="chkSwChange" name="chkSwChange" <%=strswchange%>
                                        onclick="return chkSwChange_onclick()"><font size="2" face="verdana">SCR</font></td>
                            </tr>
                        </table>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <table border="0">
                        <tr>
                            <td width="114">
                                <input type="checkbox" disabled <%=strreqchange%>><font size="2" face="verdana">Requirement</font></td>
                            <td width="60">
                                <input type="checkbox" disabled <%=strskuchange%>><font size="2" face="verdana">SKU</font></td>
                            <td>
                                <input type="checkbox" disabled <%=strimagechange%>><font size="2" face="verdana">Software</font></td>
                            <td>
                                <input type="checkbox" disabled <%=strCategoryBiosChange%>><font size="2" face="verdana">Bios</font></td>
                            <td>
                                <input type="checkbox" disabled <%=strbioschange%>><font size="2" face="verdana">BIOS</font></td>
                            <td>
                                <input type="checkbox" disabled <%=stridchange%>><font size="2" face="verdana">ID</font></td>
                        </tr>
                        <tr>
                            <td>
                                <input type="checkbox" disabled <%=strcommoditychange%>><font size="2" face="verdana">Commodity</font></td>
                            <td>
                                <input type="checkbox" disabled <%=strdocchange%>><font size="2" face="verdana">Docs</font></td>
                            <td width="60">
                                <input type="checkbox" disabled <%=strotherchange%>><font size="2" face="verdana">Other</font></td>
                            <td>
                                <input type="checkbox" disabled <%=strswchange%>><font size="2" face="verdana">SCR</font></td>
                        </tr>
                    </table>
                    <%End If %>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Summary:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <input type="text" id="txtSummary" name="txtSummary" style="width: 100%;" maxlength="120"
                            value="<%=strsummary%>">
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= replace(server.HTMLEncode(strSummary), vbcrlf, "<BR>")%>
                    <%End If %>
                </td>
            </tr>
            <%if IssueID = "" then%>
            <tr id="RowProgram" style="display:none">
                <td style="vertical-align: top">
                    <strong><font size="2">Product(s):</font><font color="red" size="1"> *</font></strong></td>
                <td>
                    <% if trim(TypeID)="2"  or  trim(TypeID)="5"  then%>
                    <input style="display: none" type="checkbox" id="chkPreinstallDeliverable" name="chkPreinstallDeliverable"
                        onclick="return chkPreinstallDeliverable_onclick()"><font size="2"
                            face="verdana" id="PDtext" onclick="return PDtext_onclick()"
                            onmouseover="return PDtext_onmouseover()">
                            <!--Product Independent<br>-->
                        </font>
                    <%else%>
                    <input style="display: none" type="checkbox" id="Checkbox1" name="chkPreinstallDeliverable"
                        onclick="return chkPreinstallDeliverable_onclick()">
                    <%end if%>
                    <table>
                        <tr>
                            <td valign="top">
                                <%
                                    '---If TypeID is 3, show products table list, else show select drop-down list
                                    If Trim(TypeID) = "3" Then
                                        Call BuildProductsList()
                                    Else   
                                %>                                
                                    <select size="2" id="lstProducts" name="lstProducts" style="width: 180px; height: 220px;" multiple>
                                    <%=strproducts%>
                                </select>
                                <font color="green" size="1" face="verdana" id="multiselect">
                                    <br>Use CRTL or SHIFT to multi-select</font>
                                <%End If%>
                            </td>
                            <td valign="top">
                                <div class="select-group"><p class="select-title">Select Products by Product Group</p>
                                    <%=strCycleProductLinks%>
                                </div>
                            </td>
                            <td valign="top">
                                <div class="select-businesssegment"><p class="select-title">Select Products by Business Segment</p>
                                    <%
                                        Call BuildProductSelectionList(1)
                                    %>
                                </div>
                            </td>
                            <td valign="top">
                                <div class="select-productlines"><p class="select-title">Select Products by Product Lines</p>
                                    <%
                                        Call BuildProductSelectionList(2)
                                    %>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>


            <%end if%>
            <tr id="RowRootDeliverable" style="display:none">
                <td style="vertical-align: top; white-space:nowrap; font-weight:bold; width:160px; font-size:small;">Root Deliverable: <span style="font-size:xx-small; color:Red">*</span></td>
                <td style="white-space:nowrap; vertical-align:top;" colspan="2">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <input type="hidden" id="hidDeliverableRootId" name="hidDeliverableRootId" value="<%=strDeliverableRootId %>" /> </br>
                        <input id="txtDeliverableRootName" name="txtDeliverableRootName" style="width: 90%; height: 22px" size="28" readonly="readonly"
                            value="<%=strDeliverableRootName%>" maxlength="255" />&nbsp;<input id="btnFindDeliverableRoot" type="button" value="Find" onclick="cmdFindRoot_onclick()" />
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= server.HTMLEncode(strNotify) %>
                    <%End If %>
                </td>
            </tr>
            <%if ((isTopPM or isPC or isSM  or (isDcrOwner and ( isSustainingTeam or SustainingProduct) ) ) or (TypeID <> "3")) then%>
            <!-- and (NOT (strStatus = "4" Or strStatus = "5")) -->
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Status:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <select id="cboStatus" name="cboStatus" style="width: 180px;" onchange="return cboStatus_onchange()">
                        <%=strStatuses%>
                    </select>
                </td>
            </tr>
            <%else%>
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Status:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <font face="verdana" size="2">
                        <%=strStatusText%>
                    </font>
                    <select id="cboStatus" name="cboStatus" style="display: none; width: 180px;" onchange="return cboStatus_onchange()">
                        <%=strStatuses%>
                    </select>
                </td>
            </tr>
            <%end if%>
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Owner:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <%If strStatus = "4" Or strStatus = "5" or ( not (isTopPM or isSM or isDcrOwner))  Then %>
                        <div style="display: none">
                            <select id="cboOwner" name="cboOwner" style="width: 180px;" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                                onkeydown="return combo_onkeydown()">
                                <%=strowners%>
                            </select>&nbsp;<INPUT disabled type="button" value="Add" id=cmdOwnerAdd name=cmdOwnerAdd onclick="return cmdOwnerAdd_onclick()">
                        </div>
                        <font face="verdana" size="2"><%= strOwner %></font>

                    <%Else %>
                        <select id="cboOwner" name="cboOwner" style="width: 180px;" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                            onkeydown="return combo_onkeydown()">
                            <%=strowners%>
                        </select>&nbsp;<INPUT type="button" value="Add" id=cmdOwnerAdd name=cmdOwnerAdd onclick="return cmdOwnerAdd_onclick()">
                    <%End If %>


                </td>
            </tr>
            <tr id="rowBusiness" style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Business:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2" nowrap>
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <table>
                            <tr>
                                <td width="125">
                                    <input type="checkbox" id="chkConsumer" name="chkConsumer" <%=strconsumer%>><font
                                        size="2">Consumer</font></td>
                                <td width="100">
                                    <input type="checkbox" id="chkSMB" name="chkSMB" <%=strsmb%>><font size="2">SMB</font></td>
                                <td width="100">
                                    <input type="checkbox" id="chkCommercial" name="chkCommercial" <%=strcommercial%>><font
                                        size="2">Commercial</font></td>
                                <td>
                                    &nbsp;<input type="button" value=" All " id="cmdAllBus" name="cmdAllBus" onclick="return cmdAllBus_onclick()"></td>
                            </tr>
                        </table>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <table>
                        <tr>
                            <td width="125">
                                <input type="checkbox" id="Checkbox2" <%=strconsumer%> disabled="disabled"><font
                                    size="2">Consumer</font></td>
                            <td width="100">
                                <input type="checkbox" id="Checkbox3" <%=strsmb%> disabled="disabled"><font size="2">SMB</font></td>
                            <td width="100">
                                <input type="checkbox" id="Checkbox4" <%=strcommercial%> disabled="disabled"><font
                                    size="2">Commercial</font></td>
                            <td>
                                &nbsp;</td>
                        </tr>
                    </table>
                    <%End If %>
                </td>
            </tr>
            <%if stridchange <> "" then %>
                <tr id="rowRegions" style="display: none">
            <%else%>
                <tr id="rowRegions" style="display: <%=DisplayForChangeOnly%>">
            <%end if %>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Regions:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2" nowrap>
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <table>
                            <tr>
                                <td width="80">
                                    <input type="checkbox" id="chkNA" name="chkNA" <%=strNA%>><font
                                        size="2">NA</font></td>
                                <td width="80">
                                    <input type="checkbox" id="chkLA" name="chkLA" <%=strLA%>><font
                                        size="2">LA</font></td>
                                <td width="80">
                                    <input type="checkbox" id="chkEMEA" name="chkEMEA" <%=stremea%>><font size="2">EMEA</font></td>
                                <td width="80">
                                    <input type="checkbox" id="chkAPJ" name="chkAPJ" <%=strapj%>><font size="2">APJ</font></td>
                                <td>
                                    &nbsp;<input type="button" value=" All " id="cmdAllGeos" name="cmdAllGeos" onclick="return cmdAllGeos_onclick()"></td>
                            </tr>
                        </table>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <table>
                        <tr>
                            <td width="125">
                                <input type="checkbox" id="Checkbox5" disabled <%=strNA%>><font size="2">NA</font></td>
                            <td width="125">
                                <input type="checkbox" id="Checkbox8" disabled <%=strLA%>><font size="2">LA</font></td>
                            <td width="100">
                                <input type="checkbox" id="Checkbox6" disabled <%=stremea%>><font size="2">EMEA</font></td>
                            <td width="100">
                                <input type="checkbox" id="Checkbox7" disabled <%=strapj%>><font size="2">APJ</font></td>
                            <td>
                                &nbsp;</td>
                        </tr>
                    </table>
                    <%End If %>
                </td>
            </tr>
            <%if trim(TypeID) = "3" then%>
            <tr style="display: none">
                <td nowrap width="160" style="vertical-align: top">
                    <strong><font size="2">Sys Team Rep:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <select id="Select4" name="cboCoreTeam" style="width: 180px;">
                        <%=strReps%>
                    </select>
                </td>
            </tr>
            <%elseif trim(TypeID) <> "4" then%>
            <tr style="display: <%=DisplayForAdd%>">
                <td nowrap width="160" style="vertical-align: top">
                    <strong><font size="2">Sys Team Rep:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <select id="Select2" name="cboCoreTeam" style="width: 180px;">
                        <%=strReps%>
                    </select>
                </td>
            </tr>
            <%else%>
            <tr>
                <td nowrap width="160" style="vertical-align: top">
                    <strong><font size="2">Sys Team Rep:</font><font color="red" size="1"> *</font></strong></td>
                <td colspan="2">
                    <select id="Select5" name="cboCoreTeam" style="width: 180px;">
                        <%=strReps%>
                    </select>
                </td>
            </tr>
            <%end if%>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Description:</font></strong>
                    <% if trim(TypeID) = "5"  or trim(TypeID) = "3" then%>
                    <font color="red" size="1"><strong>*</strong></font>
                    <%end if%>
                </td>
                <td colspan="2">
                    <%if trim(TypeID) = "5" then%>
                    <%
				dim DescriptionArray
				DescriptionArray = split(strDescription,chr(1))
				dim strPositiveDesc
				dim strNegativeDesc
				if strDescription = "" then
					strPositiveDesc =""
					strNegativeDesc= ""
				else
					if DescriptionArray(0) <> "" then
						strPositiveDesc = DescriptionArray(0)
					else
						strPositiveDesc = ""
					end if
					if ubound(DescriptionArray) > 0 then
						if DescriptionArray(1) <> "" then
							strNegativeDesc = DescriptionArray(1)
						else
							strNegativeDesc = ""
						end if
					else
						strNegativeDesc = ""
					end if
				end if
                    %>
                    <font size="2" face="verdana"><b>Positive Impact:</b><br>
                    </font>
                    <textarea id="txtPositiveDescription" style="width: 100%; height: <%=DescriptionHeight%>px"
                        name="txtPositiveDescription" rows="5" cols="42"><%=strPositiveDesc%></textarea>
                    <font size="2" face="verdana"><b>Negative Impact:</b><br>
                    </font>
                    <textarea id="txtNegativeDescription" style="width: 100%; height: <%=DescriptionHeight%>px"
                        name="txtNegativeDescription" rows="5" cols="42"><%=strNegativeDesc%></textarea>
                    <textarea id="txtDescription" style="display: none; width: 100%; height: <%=DescriptionHeight%>px"
                        name="txtDescription" rows="5" cols="42"></textarea>
                    <%elseif trim(TypeID) <> "4" then%>
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                    <%End If %>
                        <%If trim(TypeID) = "3" And strProgramID = "0" Then%>
                            <textarea id="txtDescription" style="width: 100%; height: 120px; color: blue; font-style: italic"
                                        name="txtDescription" rows="5" cols="42" onfocus="return txtDescription_onfocus()"
                                        onblur="return txtDescription_onblur()"><%= server.HTMLEncode(strDescription)%></textarea>
                        <%Else%>
                            <textarea id="txtDescription" style="width: 100%; height: <%=DescriptionHeight%>px" name="txtDescription"
                                rows="5" cols="42"><%=server.HTMLEncode(strDescription)%></textarea>
                        <%End If %>
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= replace(server.HTMLEncode(strDescription), vbcrlf, "<BR>") %>
                    <%End If %>
                    <%else%>
                    <a href="javascript:formatSelection('Bold');">
                        <img style="border-right: gainsboro thin outset; border-top: white thin outset; border-left: white thin outset;
                            border-bottom: gainsboro  thin outset" alt="Bold" src="../../images/BLD.BMP"></a><a
                                href="javascript:formatSelection('Italic');"><img style="border-right: gainsboro thin outset;
                                    border-top: white thin outset; border-left: white thin outset; border-bottom: gainsboro  thin outset"
                                    alt="Italic" src="../../images/ITL.BMP"></a><a href="javascript:formatSelection('Underline');"><img
                                        style="border-right: gainsboro thin outset; border-top: white thin outset; border-left: white thin outset;
                                        border-bottom: gainsboro  thin outset" alt="Underline" src="../../images/UNDRLN.BMP"></a>
                    <a href="javascript:InsertFormat('insertunorderedlist');">
                        <img style="border-right: gainsboro thin outset; border-top: white thin outset; border-left: white thin outset;
                            border-bottom: gainsboro  thin outset" alt="Unordered List" src="../../images/uordList.BMP"></a><a
                                href="javascript:InsertFormat('insertorderedlist');"><img style="border-right: gainsboro thin outset;
                                    border-top: white thin outset; border-left: white thin outset; border-bottom: gainsboro  thin outset"
                                    alt="Ordered List" src="../../images/ordList.BMP"></a> <a href="javascript:InsertFormat('outdent');">
                                        <img style="border-right: gainsboro thin outset; border-top: white thin outset; border-left: white thin outset;
                                            border-bottom: gainsboro  thin outset" alt="Outdent" src="../../images/outdent.BMP"></a><a
                                                href="javascript:InsertFormat('indent');"><img style="border-right: gainsboro thin outset;
                                                    border-top: white thin outset; border-left: white thin outset; border-bottom: gainsboro  thin outset"
                                                    alt="Indent" src="../../images/indent.BMP"></a>
                    <a href="javascript:InsertFormat('createLink');">
                        <img style="border-right: gainsboro thin outset; border-top: white thin outset; border-left: white thin outset;
                            border-bottom: gainsboro  thin outset" alt="Hyperlink" src="../../images/Hlink.BMP"></a>
                    <br>
                    <iframe id="myEditor" style="width: 100%; height: 250px" noresize></iframe>

                    <script>
                        function formatSelection(strFormat) {
                            // Get a text range for the selection
                            frames.myEditor.focus();
                            var tr = frames.myEditor.document.selection.createRange();

                            // Execute command
                            tr.execCommand(strFormat);

                            // Reselect and give the focus back to the editor
                            tr.select()
                            frames.myEditor.focus();
                        }


                        function InsertFormat(strFormat) {
                            // Get a text range for the selection
                            frames.myEditor.focus();
                            var tr = frames.myEditor.document.selection.createRange();

                            // Execute  command
                            //	frames.myEditor.document.execCommand(strFormat);
                            tr.execCommand(strFormat);

                            if (strFormat == "createLink") {
                                var oAnchors = frames.myEditor.document.all.tags("A");
                                if (oAnchors != null) {
                                    for (var i = oAnchors.length - 1; i >= 0; i--)
                                        if (oAnchors[i].target != "_blank")
                                        oAnchors[i].target = "_blank"
                                }
                            }

                            // Reselect and give the focus back to the editor
                            frames.myEditor.focus();
                        }


                        function showData() {
                            alert(frames.myEditor.document.body.innerHTML);
                        }

                        //frames.myEditor.document.designMode = "On"
					
                    </script>
                         <%If trim(TypeID) = "3" And strProgramID = "0" Then%>
                            <textarea id="txtDescription" style="width: 100%; height: 120px; color: blue; font-style: italic"
                                        name="txtDescription" rows="5" cols="42" onfocus="return txtDescription_onfocus()"
                                        onblur="return txtDescription_onblur()"><%= server.HTMLEncode(strDescription)%></textarea>
                         <%Else%>
                            <textarea id="txtDescription" style="display: none; width: 100%; height: <%=DescriptionHeight%>px"
                                name="txtDescription" rows="5" cols="42"><%=strDescription%></textarea>
                         <%End If%>
                    <%end if%>
                </td>
            </tr>
            <% if strIDChange <> "" then %>
                <tr style="display:none">
            <%else%>
                <tr style="display: <%=DisplayForAdd%>">
            <%end if%>
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Details: </font></td>
                <td colspan="2">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <textarea id="txtDetails" style="width: 100%; height: 120px" name="txtDetails"><%= strDetails%></textarea>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= replace(server.HTMLEncode(strDetails), vbcrlf, "<BR>") %>
                    &nbsp;<%End If %>
                </td>
            </tr>
            <!--//PBI 5621 - Prevent ODM User from seeing Justification//-->
            <%if not isODM then%>
                <%if trim(TypeID) = "5" then%>
                <tr>
                    <td width="160" style="vertical-align: top">
                        <font size="2"><strong>Root Cause:</strong></font>
                        <% if strStatus = "2" or strStatus = "4" or strStatus = "5" then %>
                        <font id="RequireJustification" color="red" size="1"><strong>*</strong></font>
                        <%end if%>
                    </td>
                    <td colspan="2">
                        <textarea id="txtJustification" style="width: 100%; height: 120px" name="txtJustification"><%=strJustification%></textarea>
                    </td>
                </tr>
                <%else%>
                <tr style="display: <%=DisplayForChangeOnly%>">
                    <td width="160" style="vertical-align: top">
                        <strong><font size="2">Justification:</font><font color="red" size="1" id="RequireJustification" > *</font></strong></td>
                    <td colspan="2">
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                        <div style="display: none">
                            <%End If %>
                            <%if JustificationTemplate = "" then%>
                            <textarea id="txtJustification" style="width: 100%; height: 120px" name="txtJustification"
                                rows="5" cols="42" onfocus="return txtJustification_onfocus()"
                                onblur="return txtJustification_onblur()"><%= server.HTMLEncode(strJustification)%></textarea>
                            <%else%>
                            <textarea id="txtJustification" style="width: 100%; height: 120px; color: blue; font-style: italic"
                                name="txtJustification" rows="5" cols="42" onfocus="return txtJustification_onfocus()"
                                onblur="return txtJustification_onblur()"><%= server.HTMLEncode(strJustification)%></textarea>
                            <%end if%>
                            <% If strStatus = "4" Or strStatus = "5" Then %>
                        </div>
                        <%= replace(server.HTMLEncode(strJustification), vbcrlf, "<BR>")%>
                        <%End If %>
                    </td>
                </tr>
                <%end if%>
            <%end if%>

            <%if trim(TypeID) <> "4" then%>
            <tr style="display: <%=DisplayForAdd%>">
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Resolution:</strong></font>
                    <% if strStatus = "2" or strStatus = "4" or strStatus = "5" then %>
                    <font id="RequireResolution" color="red" size="1"><strong>*</strong></font>
                    <%else%>
                    <font id="RequireResolution" color="red" size="1" style="display: none"><strong>*</strong></font>
                    <%end if%>
                </td>
                <td colspan="2">
                <%' if strStatus = "4" or strStatus = "5" then %'><div style="display:none"><'% end if %>
                    <textarea id="txtResolution" style="width: 100%; height: 80px" name="txtResolution"
                        rows="5" cols="42"><%=server.HTMLEncode(strResolution)%></textarea>
                <%' if strStatus = "4" or strStatus = "5" then %'></div><'%= replace(server.HTMLEncode(strResolution), vbcrlf, "<br />") & "&nbsp;" %'><% end if %>
                </td>
            </tr>
            <%else%>
            <tr style="display: none">
                <td colspan="3">
                    <font id="RequireResolution" color="red" size="1" style="display: none"></font>
                    <textarea id="txtResolution" style="display: none; width: 100%; height: 80px" name="txtResolution"
                        rows="5" cols="42"></textarea></td>
            </tr>
            <%end if%>

            <tr id="rowBiosChange1" style="display: none">
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Change / Feature Request: <font color="red" size="1">*</font></strong></font></td>
                <td colspan="2">
                    <input type="radio" id="rbNewBiosFeature" name="rbBiosNewChange" value="New Bios Feature" />&nbsp;New
                    Feature &nbsp;&nbsp;&nbsp;
                    <input type="radio" id="rbChangeBiosFeature" name="rbBiosNewChange" value="Change of an existing feature" />&nbsp;Change
                    Existing Feature
                </td>
            </tr>
            <tr id="rowBiosChange2" style="display: none">
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Current Implementation: <font color="red" size="1">*</font></strong></font>
                </td>
                <td colspan="2">
                    How does this feature work in current products?<br />
                    <textarea id="txtBiosCurrentImp" style="width: 100%; height: 60px" name="txtBiosCurrentImp"></textarea>
                </td>
            </tr>
            <tr id="rowBiosChange3" style="display: none">
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Future Implementation: <font color="red" size="1">*</font></strong></font>
                </td>
                <td colspan="2">
                    How should this feature work for future products?<br />
                    <textarea id="txtBiosFutureImp" style="width: 100%; height: 60px" name="txtBiosFutureImp"></textarea>
                </td>
            </tr>

            <%if trim(TypeID) = "5" then
		dim ActionArray
				ActionArray = split(strAction,chr(1))
				dim strCorrectiveAction
				dim strPreventiveAction
				if strAction = "" then
					strCorrectiveAction =""
					strPreventiveAction = ""
				else
					if ActionArray(0) <> "" then
						strCorrectiveAction = ActionArray(0)
					else
						strCorrectiveAction = ""
					end if
					if ubound(ActionArray) > 0 then
						if ActionArray(1) <> "" then
							strPreventiveAction = ActionArray(1)
						else
							strPreventiveAction = ""
						end if
					else
						strPreventiveAction = ""
					end if
				end if

            %>
            <tr>
                <td width="160" style="vertical-align: top">
                    <font size="2"><strong>Actions Needed:</strong></font>
                    <% if strStatus = "2" or strStatus = "4" or strStatus = "5" then %>
                    <span id="RequireActions"><font color="red" size="1"><strong>*</strong></font><br>
                        <font size="1" color="green">(Both fields Required)</font></span>
                    <%end if%>
                </td>
                <td colspan="2">
                    <font size="2" face="verdana"><b>Corrective Actions:</b></font><br>
                    <textarea id="txtCorrectiveActions" style="width: 100%; height: 120px" name="txtCorrectiveActions"
                        rows="5" cols="42"><%=strCorrectiveAction%></textarea>
                    <font size="2" face="verdana"><b>Preventive Actions:</b></font><br>
                    <textarea id="txtPreventiveActions" style="width: 100%; height: 120px" name="txtPreventiveActions"
                        rows="5" cols="42"><%=strPreventiveAction%></textarea>
                    <textarea id="txtActions" style="display: none; width: 100%; height: 120px" name="txtActions"
                        rows="5" cols="42"></textarea>
                </td>
            </tr>
            <%elseif trim(TypeID) <> "4" and IssueID <> "" then%>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Actions Needed:</font></strong></td>
                <td colspan="2">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <textarea id="Textarea5" style="width: 100%; height: 120px" name="txtActions" rows="5"
                            cols="42"><%=strAction%></textarea>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                        <%if trim(strAction) = "" then %>
                            none
                        <%else%>
                            <%= replace(server.HTMLEncode(strAction), vbcrlf, "<BR>") & "&nbsp;"%>
                        <%end if%>
                    <%End If %>
                </td>
            </tr>
            <%else%>
            <tr style="display: none">
                <td colspan="3">
                    <textarea id="Textarea6" style="display: none; width: 100%; height: 120px" name="txtActions"
                        rows="5" cols="42"></textarea></td>
            </tr>
            <%end if%>

            <%if stridchange <> "" then %>
                <tr id="rowZsrp" style="display: none">
            <%else%>
                <tr id="rowZsrp" style="display: <%=DisplayForChangeOnly%>">
            <%end if%>
                <td style="vertical-align:top;white-space:nowrap">
                    <span style="font-size:small; font-weight:bold">Additional Options:</span><span id="divZsrpRequired" style="color:Red; font-weight:bold; font-size:xx-small; visibility:hidden;">*</span></td>
                <td colspan="2" style="white-space:nowrap">
                    <table>
                        <tr>
                            <td>
                                <% IF (isTopPM or isPC) and NOT(isODM) Then %>
                                    <input type="checkbox" id="chkZsrpRequired" name="chkZsrpRequired" onclick="chkZsrpRequired_onclick()" <%=strZsrpRequired%>> <font size="2">ZSRP Ready Date Required</font>
                                <% ELSE %>
                                    <input type="checkbox" id="chkZsrpRequired" name="chkZsrpRequired" disabled <%=strZsrpRequired%>> <font size="2">ZSRP Ready Date Required</font>
                                <% END IF %>
                                <input type="checkbox" id="chkAVRequired" name="chkAVRequired" <%=strAVRequired%>> <font size="2">AV Required</font>
                                <input type="checkbox" id="chkQualificationRequired" name="chkQualificationRequired" <%=strQualificationRequired%>> <font size="2">Qualification Required</font>
                            </td>
                        </tr>
                        <tr id="rowZsrpTarget" style="display:none">
                            <td>
                                <% IF NOT(isODM) Then 
                                      IF (isTopPM or isPC) Then %>
                                        Target:&nbsp;<input id="txtZsrpReadyTargetDt" name="txtZsrpReadyTargetDt" style="width: 180px; height: 22px" size="28" value="<%=strZsrpReadyTargetDt%>" />
                                        <a href="javascript: cmdDate_onclick('txtZsrpReadyTargetDt')">
                                        <img id="picZsrpReadyTargetDt" src="images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>
                                   <% ELSE %>
                                        Target:&nbsp;<%=strZsrpReadyTargetDt %>
                                   <% END IF 
                                   END IF %>
                            </td>
                        </tr>
                        <tr id="rowZsrpActual" style="display:none">
                            <td>
                                <% IF NOT(isODM) Then 
                                     IF (isTopPM or isPC) Then %>
                                        Actual:&nbsp;<input id="txtZsrpReadyActualDt" name="txtZsrpReadyActualDt" style="width: 180px; height: 22px" size="28" value="<%=strZsrpReadyActualDt%>" />
                                        <a href="javascript: cmdDate_onclick('txtZsrpReadyActualDt')">
                                        <img id="picZsrpReadyActualDt" src="images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>
                                    <% ELSE %>
                                        Actual:&nbsp;<%=strZsrpReadyActualDt %>
                                    <% END IF 
                                   END IF %>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <%if stridchange <> "" then %>
                <tr id="rowImpact" style="display: none">
            <%else%>    
                <tr id="rowImpact" style="display: <%=DisplayForChangeOnly%>">
            <%end if %>
                <td style="vertical-align: top;white-space:nowrap;">
                    <span style="font-size:small; font-weight:bold">Customer Impact:</span><span id="divBiosRequired1" style="color:Red; font-weight:bold; font-size:xx-small; visibility:hidden;">*</span></td>
                <td colspan="2" style="white-space:nowrap;">
                    <div id="divBiosImpact1">
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                        <div style="display: none">
                            <%End If %>
                            <table>
                                <tr>
                                    <td>
                                        <input type="checkbox" id="chkCustomers" name="chkCustomers" <%=strcustomers%>></input><font
                                            size="2">Affects images and/or BIOS on shipping products</font>
                                    </td>
                                </tr>
                            </table>
                            <% If strStatus = "4" Or strStatus = "5" Then %>
                        </div>
                        <table>
                            <tr>
                                <td>
                                    <input type="checkbox" id="Checkbox14" disabled="disabled" <%=strcustomers%>></input><font size="2">Affects
                                        images and/or BIOS on shipping products</font>
                                </td>
                            </tr>
                        </table>
                        <%End If %>
                    </div>
                    <div id="divBiosImpact2" style="display: none">
                        <textarea id="txtCustomerImpact" style="width: 100%; height: 60px" name="txtCustomerImpact" cols="60" rows="2"></textarea></div>
                </td>
            </tr>
            
            <tr <%if trim(TypeID) <> "5" then%>style="display:none;"<%End If %>>
                <td style="vertical-align: top">
                    <strong><font size="2">Impact:</font><font color="red" size="1"> *</font></strong></td>
                <td>
                    <select id="lstPriority" name="lstPriority" style="width: 180px;">
                        <%=strPriorityOptions%>
                    </select>
                </td>
            </tr>           
            
            <tr id="rowDistribution" style="display:none;"></tr>        
            
            <%if stridchange <> "" then %>
                <tr id="rowLocalization" style="display: none">
            <%ELSE %>
                <tr id="rowLocalization" style="display: <%=DisplayRestore%>">
            <%end if %>
                <td style="vertical-align: top">
                    <strong><font size="2">Localization(s):</font></strong></td>
                <td>
                    <select size="2" id="lstLanguages" name="lstLanguages" style="width: 180px; height: 121px" multiple="multiple" >
                        <%=LanguageList%>
                    </select>
                    <br />
                    <font color="green" size="1" face="verdana">&nbsp;Use CRTL or SHIFT to multi-select</font></td>
            </tr>
            <tr id="rowOperatingSystem" style="display: <%=DisplayRestore%>">
                <td style="vertical-align: top">
                    <strong><font size="2">Operating System(s):</font></strong></td>
                <td>
                    <select size="2" id="lstOS" name="lstOS" style="width: 180px; height: 121px" multiple>
                        <%=strOSList%>
                    </select>
                    <br>
                    <font color="green" size="1" face="verdana">&nbsp;Use CRTL or SHIFT to multi-select</font></td>
            </tr>
            <%if trim(TypeID) = "5" then%>
            <tr>
                <td style="vertical-align: top">
                    <strong><font size="2">Net Affect:</font></strong></td>
                <td>
                    <select id="cboNetAffect" name="cboNetAffect" style="width: 180px;">
                        <option value="0" selected></option>
                        <%if strNetAffect = "1" then%>
                        <option value="1" selected>Positive</option>
                        <%else%>
                        <option value="1">Positive</option>
                        <%end if%>
                        <%if strNetAffect = "2" then%>
                        <option value="2" selected>Negative</option>
                        <%else%>
                        <option value="2">Negative</option>
                        <%end if%>
                    </select>
                </td>
            </tr>
            <%else%>
            <tr style="display: none">
                <td colspan="3">
                    <select id="cboNetAffect" name="cboNetAffect" style="display: none; width: 180px;">
                        <option value="0" selected>
                    </select>
                </td>
            </tr>
            <%end if%>
            <%if trim(TypeID) = "5" then%>
            <tr>
                <td style="vertical-align: top">
                    <strong><font size="2">Metric Impacted:</font><font color="red" size="1"> *</font></strong></td>
                <td>
                    <%
    dim MetricArray
	MetricArray = split("Scope,Cost,Quality,Scope & Cost,Schedule & Scope,Schedule & Cost,Resources,Revenue,TCE,Other",",")
                    %>
                    <select id="cboMetricImpact" name="cboMetricImpact" style="width: 180px;">
                        <option value="" selected></option>
                        <%
			for i = lbound(MetricArray) to ubound(MetricArray)	
				if lcase(trim(strAvailableNotes)) = lcase(trim(MetricArray(i))) then
					Response.Write "<option selected>" & trim(MetricArray(i)) & "</option>"
				else
					Response.Write "<option>" & trim(MetricArray(i)) & "</option>"
				end if
			next
                        %>
                    </select>
                </td>
            </tr>
            <%else%>
            <tr style="display: none">
                <td colspan="3">
                    <select id="cboMetricImpact" name="cboMetricImpact" style="display: none; width: 180px;">
                        <option value="0" selected>
                    </select>
                </td>
            </tr>
            <%end if%>

            <tr>
                            <td valign="top" style="font-size:small"><b>Attachment&nbsp;1:</b></td>
                            <td valign="top">
                               
                                    <% if strAttachment1 = "" then %>
                                        <div id="UploadAddLinks1"><a href="javascript: UploadZip(1);">Upload</a></div>
                                        <div id="UploadRemoveLinks1" style="display:none"><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment1" style="display:none" target=_blank href=""></a><label id=UploadPath1></label></div>
                                    <% else  AttachmentArray = split(strAttachment1,"\") %>
                                        <div id="UploadAddLinks1" style="display:none"><a href="javascript: UploadZip(1);">Upload</a></div>
                                        <div id="UploadRemoveLinks1" style=""><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment1" target=_blank href="file://<%=strAttachment1%>"><%=AttachmentArray(ubound(AttachmentArray))%></a><label id=UploadPath1></label></div>
                                    <%end if %>
                           
                                <input id="txtAttachmentPath1" name="txtAttachmentPath1" type="hidden" value="<%=server.htmlencode(strAttachment1)%>" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" style="font-size:small"><b>Attachment&nbsp;2:</b></td>
                            <td valign="top">
                                
                                    <% if strAttachment2 = "" then %>
                                        <div id="UploadAddLinks2"><a href="javascript: UploadZip(2);">Upload</a></div>
                                        <div id="UploadRemoveLinks2" style="display:none"><a href="javascript: UploadZip(2);">Change</a> | <a href="javascript: RemoveUpload(2);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a style="display:none" id="hplAttachment2" style="display:none" target=_blank href=""></a><label id=UploadPath2></label></div>
                                    <% else  AttachmentArray = split(strAttachment2,"\") %>
                                        <div id="UploadAddLinks2" style="display:none"><a href="javascript: UploadZip(2);">Upload</a></div>
                                        <div id="UploadRemoveLinks2" style=""><a href="javascript: UploadZip(2);">Change</a> | <a href="javascript: RemoveUpload(2);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment2" target=_blank href="file://<%=strAttachment2%>"><%=AttachmentArray(ubound(AttachmentArray))%></a><label id=UploadPath2></label></div>
                                    <%end if %>
                         
                                <input id="txtAttachmentPath2" name="txtAttachmentPath2" type="hidden" value="<%=server.htmlencode(strAttachment2)%>" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" style="font-size:small"><b>Attachment&nbsp;3:</b></td>
                            <td valign="top">
                                
                                    <% if strAttachment3 = "" then %>
                                        <div id="UploadAddLinks3"><a href="javascript: UploadZip(3);">Upload</a></div>
                                        <div id="UploadRemoveLinks3" style="display:none"><a href="javascript: UploadZip(3);">Change</a> | <a href="javascript: RemoveUpload(3);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a style="display:none" id="hplAttachment3" style="display:none" target=_blank href=""></a><label id=UploadPath3></label></div>
                                    <% else  AttachmentArray = split(strAttachment3,"\") %>
                                        <div id="UploadAddLinks3" style="display:none"><a href="javascript: UploadZip(3);">Upload</a></div>
                                        <div id="UploadRemoveLinks3" style=""><a href="javascript: UploadZip(3);">Change</a> | <a href="javascript: RemoveUpload(3);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment3" target=_blank href="file://<%=strAttachment3%>"><%=AttachmentArray(ubound(AttachmentArray))%></a><label id=UploadPath3></label></div>
                                    <%end if %>

                                <input id="txtAttachmentPath3" name="txtAttachmentPath3" type="hidden" value="<%=server.htmlencode(strAttachment3)%>" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" style="font-size:small"><b>Attachment&nbsp;4:</b></td>
                            <td valign="top">
                                
                                    <% if strAttachment4 = "" then %>
                                    <div id="UploadAddLinks4"><a href="javascript: UploadZip(4);">Upload</a></div>
                                    <div id="UploadRemoveLinks4" style="display:none"><a href="javascript: UploadZip(4);">Change</a> | <a href="javascript: RemoveUpload(4);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a style="display:none" id="hplAttachment4" target=_blank href=""></a><label id=UploadPath4></label></div>
                                    <% else  AttachmentArray = split(strAttachment4,"\") %>
                                        <div id="UploadAddLinks4" style="display:none"><a href="javascript: UploadZip(4);">Upload</a></div>
                                        <div id="UploadRemoveLinks4" style=""><a href="javascript: UploadZip(4);">Change</a> | <a href="javascript: RemoveUpload(4);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment4" target=_blank href="file://<%=strAttachment4%>"><%=AttachmentArray(ubound(AttachmentArray))%></a><label id=UploadPath4></label></div>
                                    <%end if %>
                                
                                <input id="txtAttachmentPath4" name="txtAttachmentPath4" type="hidden" value="<%=server.htmlencode(strAttachment4)%>" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" style="font-size:small"><b>Attachment&nbsp;5:</b></td>
                            <td valign="top">
                                
                                    <% if strAttachment5 = "" then %>
                                    <div id="UploadAddLinks5"><a href="javascript: UploadZip(5);">Upload</a></div>
                                    <div id="UploadRemoveLinks5" style="display:none"><a href="javascript: UploadZip(5);">Change</a> | <a href="javascript: RemoveUpload(5);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a style="display:none" id="hplAttachment5" target=_blank href=""></a><label id=UploadPath5></label></div>
                                    <% else  AttachmentArray = split(strAttachment5,"\") %>
                                    <div id="UploadAddLinks5" style="display:none"><a href="javascript: UploadZip(5);">Upload</a></div>
                                    <div id="UploadRemoveLinks5" style=""><a href="javascript: UploadZip(5);">Change</a> | <a href="javascript: RemoveUpload(5);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><a id="hplAttachment5" target=_blank href="file://<%=strAttachment5%>"><%=AttachmentArray(ubound(AttachmentArray))%></a><label id=UploadPath5></label></div>
                                    <%end if %>
                                
                                <input id="txtAttachmentPath5" name="txtAttachmentPath5" type="hidden" value="<%= server.htmlencode(strAttachment5)%>" />
                            </td>
                        </tr>
            <tr style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: middle">
                    <strong><font size="2"><span id="SampleText">
                    <%if stridchange <> "" then %>
                        Deadline:
                    <%else %>
                        Samples Available:
                    <%end if%>
                    </span></font></strong></td>
                <td colspan="2">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <input id="txtAvailDate" name="txtAvailDate" style="width: 180px; height: 22px" size="28" class="dateselection"
                            value="<%=strAvailableForTest%>" onkeydown="return false;" autocomplete="off" >
                        <!--<a href="javascript: cmdAvailDate_onclick()">
                            <img id="picTarget" src="images/calendar.gif" alt="Choose Date" border="0" width="26"
                                height="21"></a><input type="button" value="Choose" id="cmdAvailDate" name="cmdAvailDate" onclick="return cmdAvailDate_onclick()">-->
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= Server.HTMLEncode(strAvailablefortest) %>
                    &nbsp;<%End If %>
                </td>
            </tr>
                      
            
            <%if trim(TypeID) <> "4" then%>
            <tr style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: middle">
                    <strong><font size="2"><span id="Span1">Target&nbsp;Approval&nbsp;Date:</span></font></strong>
                </td>
                <td colspan="2">
                 <% If strStatus = "4" Or strStatus = "5" Or (IsDcrOwner = false and isSubmitter = false and strID <> "") Then %>
                    <div style="display: none">
                 <%End If %>
                    <input id="txtTargetApprovalDate" name="txtTargetApprovalDate" style="width: 180px; height: 22px" size="28" class="dateselection"
                            value="<%=strTargetApprovalDate%>" onkeydown="return false;" autocomplete="off">                    
                 <% If strStatus = "4" Or strStatus = "5" Or (IsDcrOwner = false and isSubmitter = false and strID <> "") Then %>
                    </div>
                        <% if strTargetApprovalDate = "" then %>
                        N/A <% else %><%= Server.HTMLEncode(strTargetApprovalDate) %><% end if %>&nbsp;
                 <%End If %>
                 </td>
            </tr>
            <tr style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: middle">
                    <strong><font size="2"><span id="spanImportant">Important:</span></font></strong>
                </td>
                <td colspan="2"><input type="checkbox" id="chkImportant" name="chkImportant" <%=strImportant%>></td>
            </tr>
            <tr style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: middle">
                    <strong><font size="2"><span id="Span1">RTP&nbsp;Date:</span></font></strong>
                </td>
                <td colspan="2">
                 <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                 <%End If %>
                    <input id="txtRTPDate" name="txtRTPDate" style="width: 180px; height: 22px" size="28" class="dateselection"
                            value="<%=strRTPDate%>" onkeydown="return false;" autocomplete="off">                    
                 <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                        <%= Server.HTMLEncode(strRTPDate) %>&nbsp;
                 <%End If %>
                 </td>
            </tr>
            <%end if %>
            
            <%if trim(TypeID) <> "4" then%>
            <tr style="display: <%=DisplayForChangeOnly%>">
                <td width="160" style="vertical-align: middle">
                    <strong><font size="2"><span id="Span2">End&nbsp;of&nbsp;Manufacturing:</span></font></strong>
                </td>
                <td colspan="2">
                 <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                 <%End If %>
                    <input id="txtRASDiscoDate" name="txtRASDiscoDate" style="width: 180px; height: 22px" size="28" class="dateselection"
                            value="<%=strRASDiscoDate%>" onkeydown="return false;" autocomplete="off">                    
                 <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                        <%= Server.HTMLEncode(strRASDiscoDate) %>&nbsp;
                 <%End If %>
                 </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top"></td>
                <td nowrap colspan="2" valign="top">
                    <font size="1" face="verdata" color="blue">If changing AV dates for a Product Release is needed, please select only one Release for that Product.</font>
                </td>
            </tr>
            <%end if %>
            
            <%if trim(TypeID) <> "4" then%>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Notify&nbsp;on&nbsp;Approval:</font></strong></td>
                <td nowrap colspan="2" valign="top">
                    <% If strStatus = "4" Or strStatus = "5" Then %>
                    <div style="display: none">
                        <%End If %>
                        <input id="txtNotify" name="txtNotify" style="width: 90%; height: 22px" size="28"
                            value="<%=strNotify%>" maxlength="255">&nbsp;<input id="btnAddSumbitter" type="button" value="Add" onclick="cmdAdd_onclick()" />
                        <br>
                        <font size="1" face="verdata" color="blue">Email addresses other than Submitter or System
                            Team to notify on approval.<br>
                            - Use full SMTP Email Addresses (first.last@hp.com)<br>
                            - Seperate multiple addresses with a semicolon</font>
                        <% If strStatus = "4" Or strStatus = "5" Then %>
                    </div>
                    <%= server.HTMLEncode(strNotify) & "&nbsp;" %>
                    <%End If %>
                </td>
            </tr>
            <%else%>
            <tr>
                <td colspan="3">
                    <input id="txtNotify" name="txtNotify" style="display: none; width: 430; height: 22px"
                        size="28" value maxlength="255"></td>
            </tr>
            <%end if%>
            <%if trim(TypeID) <> "3"  and trim(TypeID) <> "4" then%>
            <tr>
                <%else%>
                <tr style="display: none">
                    <%end if%>
                    <td width="160" style="vertical-align: top">
                        <strong><font size="2">Sub-Owner&nbsp;Group(s):</font></strong></td>
                    <td colspan="2" nowrap>
                        <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
                            border-left: steelblue 1px solid; width: 100%; border-bottom: steelblue 1px solid;
                            height: 143px; background-color: white" id="DIV1">
                            <table id="TablePart" width="100%">
                                <thead>
                                    <tr style="position: relative; top: expression(document.getElementById('DIV1').scrollTop-2);">
                                        <td bgcolor="lightsteelblue" nowrap style="border-right: 1px outset; border-top: 1px outset;
                                            border-left: 1px outset; border-bottom: 1px outset">
                                            &nbsp;</td>
                                        <td bgcolor="LightSteelBlue" style="width: 100%; border-right: 1px outset; border-top: 1px outset;
                                            border-left: 1px outset; border-bottom: 1px outset">
                                            <font size="2" face="verdana">&nbsp;Functional&nbsp;Group&nbsp;</font></td>
                                    </tr>
                                </thead>
                                <%
				dim strSaveString
				dim strLoadedGroups
				if IssueID = "" then
					rs.open "spListGroups4Action 0",cn,adOpenForwardOnly
				else
					rs.open "spListGroups4Action " & clng(IssueID),cn,adOpenForwardOnly
				end if
				strSaveString = ""
				strLoadedGroups = ""
				do while not rs.eof
					if trim(rs("ID") & "") = "" then
						strSaveString = strSaveString & "<TR><TD><INPUT type=""checkbox"" value=""" & rs("GroupID") & """ id=lstFunctionalGroup name=lstFunctionalGroup></TD><TD>" & rs("GroupName") & "</TD></TR>"
					else
						Response.Write "<TR><TD><INPUT type=""checkbox"" checked value=""" & rs("GroupID") & """ id=lstFunctionalGroup name=lstFunctionalGroup></TD><TD>" & rs("GroupName") & "</TD></TR>"
						strLoadedGroups = strLoadedGroups & "," & rs("GroupID")
					end if
					rs.movenext	
				loop
				rs.close
				if strLoadedGroups <> "" then
					strLoadedGroups = mid(strLoadedGroups,2)
				end if

				if strSaveString <> "" then
					Response.Write strSaveString
				end if				
                                %>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr style="display: <%=strDisplayReport%>">
                    <td width="160" style="vertical-align: top">
                        <strong><font size="2">Status Report:</font></strong></td>
                    <td>
                        <font size="2">
                            <input type="checkbox" id="chkReports" <%=strreportvalue%> name="chkReports">
                            <label id="lblStatus" onclick="return lblStatus_onclick()"
                                onmouseover="return lblStatus_onmouseover()">
                                Remove from Online Status Reports</label></font>
                    </td>
                </tr>
        </table>
        <%
	if noapprovals then
		Response.write "<BR>" & strApprovals
	end if
        %>

<%end if%>
        <% if not isODM then %> 
            <textarea style="display: none" rows="2" cols="20" id="txtJustificationTemplate"
            name="txtJustificationTemplate"><%=JustificationTemplate%></textarea>

            <textarea style="display: none" rows="2" cols="20" id="txtDescriptionTemplate"
            name="txtDescriptionTemplate"><%=DescriptionTemplate%></textarea>
        <% else%>
            <textarea style="display: none" rows="2" cols="20" id="txtJustificationTemplate"
            name="txtJustificationTemplate"></textarea>

            <textarea style="display: none" rows="2" cols="20" id="txtDescriptionTemplate"
            name="txtDescriptionTemplate"></textarea>
        <% end if %>
        <input style="display: none" type="text" id="txtID" name="txtID" value="<%=IssueID%>">
        <input style="display: none" type="text" id="txtType" name="txtType" value="<%=TypeID%>">
        <input type="hidden" id="Approvers2Add" name="Approvers2Add" value="">
        <input style="display: none" type="text" id="txtSaveApproval" name="txtSaveApproval"
            value="<%=strSaveApprovals%>">
        <input style="display: none" type="text" id="txtCurrentUserID" name="txtCurrentUserID"
            value="<%=CurrentUSerID%>">
        <input style="display: none" type="text" id="txtApproversLoaded" name="txtApproversLoaded"
            value="<%=ApproversLoaded%>">
        <input style="display: none" type="text" id="txtApproversPending" name="txtApproversPending"
            value="<%=ApproversPending%>">
        <input style="display: none" type="text" id="txtCommodityManagerID" name="txtCommodityManagerID"
            value="<%=strCommodityManagerID%>">
        <input type="hidden" id="txtGroupsLoaded" name="txtGroupsLoaded" value="<%=strLoadedGroups%>">
        <input type="hidden" id="hidAddDCRNotificationList" name="hidAddDCRNotificationList" value="<%=AddDCRNotificationList%>" />
        <input type="hidden" id="hidInitialZsrpReadyTargetDt" name="hidInitialZsrpReadyTargetDt" value="<%=strZsrpReadyTargetDt%>" />
        <input type="hidden" id="hidInitialstrZsrpReadyActualDt" name="hidInitialstrZsrpReadyActualDt" value="<%=strZsrpReadyActualDt%>" />
        <input type="hidden" id="hidCurrentUserPartner" name="hidCurrentUserPartner" value="<%=CurrentUserPartner %>"/>
        <input type="hidden" id="inpChangeRequestID" name="inpChangeRequestID" value="<%=iChangeRequestID %><%=CurrentUserID %>" />
        <input type="hidden" id="hdnWorkflowID" name="hdnWorkflowID" value="<%=WorkflowID%>" />
        <input type="button" id="hdnClearProjectList" name="hdnClearProjectList" value="" hidden="hidden" onclick="javascript: hdnClearProjectList_onclick();" />
        <input type="hidden" id="hdnisDcrApprover" name="hdnisDcrApprover" value="<%=isDcrApprover%>" />
        <input type="hidden" id="hdOrgTargetApprovalDate" name="hdOrgTargetApprovalDate" value="<%=strTargetApprovalDate%>" />
        <input type="hidden" id="layout" name="layout" value="<%=Layout%>" />
    </form>
    <select id="cboEmployee" name="cboEmployee" style="display: none; width: 180px;">
        <%=strEmployee%>
    </select>
    <br>
    <div id="divApproverList">
        <select style="display: none" id="cboProdApprovalList" name="cboProdApprovalList">
            <%=strProdApprovalList%>
        </select>
    </div>
    <input type="hidden" id="txtProductName" name="txtProductName" value="<%=strProductname%>">
    <input type="hidden" id="txtProductID" name="txtProductID" value="<%=strProgramID%>">
    <input type="hidden" id="txtRecordLocked" name="txtRecordLocked" value="<%=strRecordLocked%>">
    <div id="dialog-charError" title="Characters Error">
      <p><b>Invalid characters detected(location marked with _):</b></p>
      <textarea style="width: 90%;" id="CharErrorMsg" cols="60" rows="15"></textarea>
    </div>
    <%
        if strID="0" then
            response.write "<input id=""rowZsrpActual"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""rowZsrpTarget"" style=""display:none"" type=""hidden"" />"
            response.write "<input id=""divZsrpRequired"" style=""display:none"" type=""hidden"" />"
        end if
    %>
</body>
</html>
<%
'---SUB PROCEDURES & FUNCTION: ---
'**************************************************************************************
'* Description	: Create random number; to be used as unique Change Request ID
'* Creator		: Harris, Valerie
'* Created		: 02/8/2016 - PBI 15660/ Task 16234
'**************************************************************************************/ 
Function GetRandomNumber(iMin, iMax)
    Dim Rand

    Randomize
    Rand = (Int((iMax-iMin+1)*Rnd+iMin))
    GetRandomNumber = Rand
End Function


'**************************************************************************************
'* Description	: If product was submitted as a group, get the change quest ID 
'* Creator		: Harris, Valerie
'* Created		: 02/8/2016 - PBI 15660/ Task 16234
'**************************************************************************************/ 
Function GetChangeRequestGroupID(iIssueID)
    Call OpenDBConnection(PULSARDB(), True)			'Open database connection, oConnect.
    Call GetChangeRequest(True, iIssueID)
	    If Not(oRSChangeRequest Is Nothing) Then
		    If Not oRSChangeRequest.EOF Then
				GetChangeRequestGroupID = oRSChangeRequest("ChangeRequestID")
		    Else
			    GetChangeRequestGroupID = 0
		    End If
	    End If
    Call GetChangeRequest(False, Empty)
    Call OpenDBConnection(PULSARDB(), False)	    'Close database connection, oConnect
End Function


'**************************************************************************************
'* Description	: If product was submitted as a group, get the change quest ID 
'* Creator		: Harris, Valerie
'* Created		: 02/8/2016 - PBI 15660/ Task 16234
'**************************************************************************************/ 
Sub BuildChangeRequestGroupList(iIssueID, iChangeRequestID)
    '--DECLARE LOCAL VARIABLES: ---
    Dim sHTML                       'STRING
    Dim iDelieverableIssueID        'INTEGER
    Dim sProductName                'STRING 
    Dim sProductReleases            'STRING
    Dim iProductID                  'INTEGER
    Dim sClassName                  'STRING

    Call OpenDBConnection(PULSARDB(), True)			'Open database connection, oConnect.
    Call ListChangeRequestGroup(True, iIssueID, iChangeRequestID)
	    If Not(oRSChangeRequestGroup Is Nothing) Then
		    If Not oRSChangeRequestGroup.EOF Then
                sHTML = "<div id=""divProductGroup"">"
                sHTML = sHTML & "<table cellpadding=""1"" cellspacing=""2"" class=""shade"" id=""TBLChangeRequestGroup""> "
                sHTML = sHTML & "<tbody class=""group""> "
				'---start recordset list: ---
                Do While Not oRSChangeRequestGroup.EOF
                    '--- Define Product Variables: ---
                    iDelieverableIssueID = oRSChangeRequestGroup("ID")
                    sProductName  =  oRSChangeRequestGroup("ProductName")
                    iProductID =  oRSChangeRequestGroup("ProductVersionID")
                    sProductReleases =  oRSChangeRequestGroup("ProductVersionRelease")
                     
                    '---Alternate background row color for Product rows
                    If sClassName = "" Then
                        sClassName = "alt"
                    Else
                        sClassName = ""
                    End If

                    sHTML = sHTML & "<tr class="""&sClassName&"""><td class=""w375, left""> "

                    '---Add link to Product Name when IssueID doent's match querystring value: ---
                    If CLng(iIssueID) <> CLng(iDelieverableIssueID) Then
                        sHTML = sHTML & "<a href=""action.asp?ProdID=&CAT=&ID="&iDelieverableIssueID&"&Type=3"" target=""_parent"" title=""View Change Request""> "
                        sHTML = sHTML & "<strong>"& sProductName & "</strong> "
                        sHTML = sHTML & "</a> "
                    End If
                    
                    '---Make font-larger, with no link when IssueID does match querystring value: ---
                    If CLng(iIssueID) = CLng(iDelieverableIssueID) Then 
                        sHTML = sHTML & "<span class=""font-medium""> "  
                        sHTML = sHTML & ""& sProductName & " "         
                        sHTML = sHTML & "</span> "
                    End If

                    '---If Product has Releases selected, add Release list: ---
                    If sProductReleases <> "" Then
                       sHTML = sHTML & "&nbsp;-&nbsp; " & sProductReleases & " " 
                    End If
                    
                    sHTML = sHTML & "</td></tr>"
                    oRSChangeRequestGroup.Movenext								
				Loop
                '---end recordset list: ---
                sHTML = sHTML & "</tbody> "
                sHTML = sHTML & "</table>"
                sHTML = sHTML & "</div>"
		    End If
	    End If
    Call ListChangeRequestGroup(False, Empty, Empty)
    Call OpenDBConnection(PULSARDB(), False)	    'Close database connection, oConnect
    Response.Write(sHTML)
End Sub


'**************************************************************************************
'* Description	: Build hierarchial product - release table 
'* Creator		: Harris, Valerie
'* Created		: 02/4/2016 - PBI 15660/ Task 16234
'**************************************************************************************/  
Sub BuildProductsList()
    '--DECLARE LOCAL VARIABLES: ---
    Dim sHTML                       'STRING
    Dim sDevCenterName              'STRING
    Dim iProductID                  'INTEGER
    Dim sProductName                'STRING 
    Dim iProductReleases            'INTEGER
    Dim sProductSelected            'STRING
    Dim sProductIcon                'STRING
    Dim iReleaseID                  'INTEGER
    Dim sReleaseName                'STRING
    Dim sReleaseDisplay             'STRING
    Dim sReleaseSelected            'STRING
    Dim sParentClassName            'STRING
    Dim sPCEmail                    'STRING
    Dim iProductLineID              'INTEGER
    Dim clsRelease                  'STRING    
    Call OpenDBConnection(PULSARDB(), True)			'Open database connection, oConnect.
        Call ListProducts(True)	                    'Open Recordset Connection
        Call ListProductRelease(True)        'Open recordset connection
	        If Not (oRSProducts Is Nothing) Then
		        If Not oRSProducts.EOF Then 
                    'sHTML = "<a href=""javascript: ClearProductList();"" class=""font-title"">Clear All</a><br/>"
                    sHTML = "<div id=""divProductRelease"" class=""scroll-box""> " _
                            & "<table cellpadding=""1"" cellspacing=""2"" class=""shade"" id=""TBLProducts""> " _
                            & "<tbody> "
                    '---start recordset list: ---
					Do While Not oRSProducts.EOF
                        If (trim(TypeID) = "3" and oRSProducts("AllowDCR") ) OR ( trim(TypeID) <> "3" and oRSProducts("ProductStatusID") < 5) then  '*Used same If statement that filters strproducts
                            '--- Define Product (Parent) Variables: ---
                            iProductID = oRSProducts("ID") 
                            sProductName = oRSProducts("Product")
                            iProductReleases = CLng(oRSProducts("ProductReleases")) 'Number of product releases associated with the product
                            sPCEmail = oRSProducts("PCEmail")
                            iProductLineID = oRSProducts("ProductLineID")
                            
                            '--- If Products' DevCenter variable is empty OR different than the Recordset value, add a Dev Center heading: ---
                            If oRSProducts("DevCenterName") = "" Then   '*Used same If statement that filters strproducts
                                    sDevCenterName = oRSProducts("DevCenterName")
                                    sHTML = sHTML & "<tr class=""heading-row, bg-gray"" data-devcenter="""&sDevCenterName&""">" _
                                                  & "<td class=""left"" colspan=""4""><strong>"&sDevCenterName&"</strong></td>" _
                                                  & "</tr>"
                            ElseIf oRSProducts("DevCenterName") <> sDevCenterName Then
                                    sDevCenterName = oRSProducts("DevCenterName")
                                    sHTML = sHTML & "<tr class=""heading-row, bg-gray"" data-devcenter="""&sDevCenterName&""">" _
                                                  & "<td class=""left"" colspan=""4""><strong>"&sDevCenterName&"</strong></td>" _
                                                  & "</tr>"
                            End If
                            
                            '---Alternate background row color for Product rows
                            If sParentClassName = "" Then
                                sParentClassName = "alt"
                            Else
                                sParentClassName = ""
                            End If

                            '---Selects checkbox for selected Product and, if they exists, changes +/- icon: ---
                            If strProgramID = iProductID & "" Or ProdID = iProductID & "" Then   '*Used same If statement that filters strproducts
                                sProductSelected = "checked"
                                sProductIcon = "fa fa-minus-square" 
                            Else
                                sProductSelected = ""
                                sProductIcon = "fa fa-plus-square"  
                            End If
                            
                            '--- Parent - Add Product Row to the Table: ---
                            sHTML = sHTML & "<tr class=""parent-row, "&sParentClassName&""" data-devcenter="""&sDevCenterName&""">" _
                                          & "<td class=""w10, center"">"
                            If iProductReleases > 0 Then
                                sHTML = sHTML & "<span class=""select-menu-release"" data-product="""&iProductID&""" data-productline=""" & iProductLineID & """>" _
                                              & "<em id=""option_" & iProductID & """ class=""row-option "&sProductIcon&" pointer font-darkgray""></em>" _
                                              & "</span>"
                            Else
                                sHTML = sHTML & "&nbsp;"
                            End If
                            sHTML = sHTML & "</td>" _
                                          & "<td class=""w10, center"">"
                            If iProductReleases <> 0 Then                                                        
                                sHTML = sHTML & "<input type=""checkbox"" id=""product_" & iProductID & """ data-productname=""" & sProductName & """ data-productline=""" & iProductLineID & """ name=""chkProducts"" value=""" & iProductID & """ class=""chk-product"" " & sProductSelected & " />"
                            Else
                                sHTML = sHTML & "<input type=""checkbox"" id=""product_" & iProductID & """ data-productname=""" & sProductName & """ data-productline=""" & iProductLineID & """ name=""chkProducts"" value=""" & iProductID & """ class=""chk-product"" " & sProductSelected & " />" _
                                              & "<input type=""hidden"" name=""chkRelease_"&iProductID&""" value="""" />"
                            End If
                            sHTML = sHTML & "</td>" _
                                          & "<td class=""left"" colspan=""2""> " & sProductName & "</td>" _
                                          & "</tr>"

                            '--- If Product has Release(s), add Release row(s) below the Product row: ---
                            If IsNumeric(iProductID) = True And iProductReleases > 0  Then 
                                    If Not (oRSProductRelease Is Nothing) Then
                                        oRSProductRelease.Filter = ""
                                        oRSProductRelease.Filter = "ProductVersionID=" & iProductID
		                                If Not oRSProductRelease.EOF Then 
		                                    Do While Not oRSProductRelease.EOF
                                                '--- Define Release (Child) Variables: ---
                                                iReleaseID = oRSProductRelease("ID") 
                                                sReleaseName = oRSProductRelease("Name")
                                                clsRelease = "release_" & iReleaseID & "_" & iProductID                                     
                                                '--- If Product is selected, show Release rows; else hide: ---
                                                If sProductSelected = "checked" Then
                                                    sReleaseDisplay = "show-row"
                                                    sReleaseSelected = ""
													If (iProductReleases = 1) Then
                                                       sReleaseSelected = "checked"
													End if							
                                                Else
                                                    sReleaseDisplay = "hide"
                                                    sReleaseSelected = ""
                                                End If
                                                
    
                                                '--- Child - Add Product's Release Rows to the Table: ---
                                                sHTML = sHTML & "<tr class=""child-row, "&sReleaseDisplay&""" data-product="""&iProductID&""" data-devcenter="""&sDevCenterName&""">" _
                                                              & "<td class=""w10"" colspan=""2"">&nbsp;</td>" _
                                                              & "<td class=""w10, center"">" _
                                                              & "<input type=""checkbox"" id=""release_"&iReleaseID&"_"&iProductID&""" name=""chkRelease_"&iProductID&""" value="""&sReleaseName&""" class=""chk-release,chk_group_"&iProductID&", "&clsRelease&""" " & sReleaseSelected & " />" _
                                                              & "</td>" _
                                                              & "<td class=""left"" width=""385px"">" & sReleaseName & "</td>" _            
                                                              & "</tr>"
                                                oRSProductRelease.MoveNext
		                                    Loop
                                        End If
                                    
                                    End If
                            End If
                        End If
                        oRSProducts.Movenext	
                        Response.Write(sHTML)
                        sHTML = ""
					Loop
                    sParentClassName = ""
					'---end recordset list: ---
                    sHTML = sHTML & "</tbody> " _
                                  & "</table>" _               
                                  & "</div>"
		        End If
	        End If
        Call ListProductRelease(False)              'Clost recordset connection
        Call ListProducts(False)                     'Close Recordset Connection
    Call OpenDBConnection(PULSARDB(), False)	    'Close database connection, oConnect
    Response.Write(sHTML)
End Sub

'**************************************************************************************
'* Description	: Create Product Selection List by Business Segement or Product Line
'* Creator		: Harris, Valerie
'* Created		: 10/31/2016 - PBI 27697/ Task 28523
'**************************************************************************************/ 
Sub BuildProductSelectionList(iSelectionType)
    '--DECLARE LOCAL VARIABLES: ---
    Dim sHTML                       'STRING
    Dim sSelectionTypeName          'STRING
    Dim sSelectionName              'STRING 
    Dim sProductIDs                 'STRING
    Dim sReleaseProductIDs          'STRING
    Dim sClassName                  'STRING
    Dim sRowClassName               'STRING
    Dim sFontClassName              'STRING
    Dim iSelectionID                'INTEGER
    Dim sProductLineIDs             'STRING
    Dim iReleaseID                  'INTEGER
    Dim sReleaseName                'STRING
    
    If iSelectionType = 1 Then
        sSelectionName = "BusinessSegment"
        sClassName = "list"
    Else 'if type equals 2
        sSelectionName = "ProductLine"
        sClassName = "scroll-box-small"
    End If

    Call OpenDBConnection(PULSARDB(), True)			    'Open database connection, oConnect.
        Call ListProductSelection(True, iSelectionType)	    'Open Recordset Connection		
	        If Not (oRSProductSelection Is Nothing) Then
		        If Not oRSProductSelection.EOF Then 

                    If iSelectionType = 1 Then
                        sHTML = "<div id=""div"&sSelectionName&"""> "
                        sHTML = sHTML & "<ul class=""" & sClassName & """> "
                    Else
                        sHTML = "<div id=""div"&sSelectionName&""" class=""" & sClassName & """> "
                        sHTML = sHTML & "<table cellpadding=""1"" cellspacing=""2"" class=""shade"" id=""TBLProductLines""> "
                        sHTML = sHTML & "<tbody> "
                    End If
                    
                    '---start recordset list: ---
			        Do While Not oRSProductSelection.EOF
                        sSelectionName = oRSProductSelection("SelectionName")
                        iSelectionID = oRSProductSelection("SelectionID")
                        sProductIDs = oRSProductSelection("ProductIDs")  

                        If iSelectionType = 1 Then
                            If sProductIDs <> "" Then                              
                                '--- If Business Segment has Release(s), add Release list(s) below the Business Segment list item: ---
                                Call ListProductOption(True, 3, iSelectionID)      'Open recordset connection
                                    If Not (oRSProductOption Is Nothing) Then
		                                If Not oRSProductOption.EOF Then
                                            sHTML = sHTML & "<li class=""nobullet""><a href=""javascript: ShowBusSegReleases(" & iSelectionID & ");"">" & sSelectionName & "</a> "  
                                            sHTML = sHTML & "<ul id=""busseg_" & iSelectionID & """ class=""child-list hide""> "
		                                        Do While Not oRSProductOption.EOF
                                                    '--- Define Release (Child) Variables: ---
                                                    iReleaseID = oRSProductOption("SelectionID") 
                                                    sReleaseName = oRSProductOption("SelectionName")
                                                    sReleaseProductIDs = oRSProductOption("ProductIDs") 
                                                    sHTML = sHTML & "<li class=""square""><a href=""javascript: SelectMultipleProductCheckbox('" & sReleaseProductIDs & "', 'Business Segment', '" & iReleaseID & "');"">" & sReleaseName & "</a></li> "
                                                    oRSProductOption.MoveNext
		                                        Loop
                                            sHTML = sHTML & "</ul> "
                                        Else
                                            sHTML = sHTML & "<li class=""nobullet font-darkgray"">" & sSelectionName & " "                
                                        End If                                   
                                    End If
                                Call ListProductOption(False, Empty, Empty)         'Close recordset connection
                                sHTML = sHTML & "</li> "
                            Else
                                sHTML = sHTML & "<li class=""nobullet font-darkgray"">" & sSelectionName & "</li> "                          
                            End If
                        Else
                            '---Alternate background row color for Product Line rows
                            If sRowClassName = "" Then
                                sRowClassName = "alt"
                            Else
                                sRowClassName = ""
                            End If
                            
                            sHTML = sHTML & "<tr class="""&sRowClassName&""">"
                            If sProductIDs <> "" Then
                                sFontClassName = ""
                                sHTML = sHTML & "<td class=""center"" width=""8%"">"
                                sHTML = sHTML & "<input type=""checkbox"" id=""productline_" & sSelectionName & """ data-productline=""" & iSelectionID & """ value=""" & sProductIDs & """ class=""chk-productline"" /> "
                                sHTML = sHTML & "</td>"
                            Else
                                sFontClassName = "font-darkgray"
                                sHTML = sHTML & "<td class=""w5, center"">"
                                sHTML = sHTML & "<input type=""checkbox"" id=""productline_" & sSelectionName & """ data-productline=""" & iSelectionID & """ value="""" disabled=""true"" /> "
                                sHTML = sHTML & "</td>"
                            End If
                            sHTML = sHTML & "<td class=""left, "&sFontClassName&""" width=""50%""> " & sSelectionName & "</td>"
                            sHTML = sHTML & "</tr>"                       
                            
                            'Create concatenated string of Product Line IDs
                            If sProductLineIDs = "" Then
                                sProductLineIDs = iSelectionID 
                            Else
                                sProductLineIDs = sProductLineIDs & "," & iSelectionID
                            End If
                        End If                    
                        oRSProductSelection.Movenext								
					Loop
					'---end recordset list: ---
                    If iSelectionType = 1 Then
                        sHTML = sHTML & "</ul> "
                        sHTML = sHTML & "</div>"
                    Else
                        sHTML = sHTML & "</tbody> "
                        sHTML = sHTML & "</table> "                 
                        sHTML = sHTML & "</div>"
                    End If

                    '--Add hidden field with prodline ids
                    If iSelectionType = 2 Then
                        sHTML = sHTML & "<input type=""hidden"" id=""inpProductLineIDs"" value=""" & sProductLineIDs & """>"
                    End If
		        End If
	        End If
        Call ListProductSelection(False, Empty)          'Close Recordset Connection
    Call OpenDBConnection(PULSARDB(), False)	        'Close database connection, oConnect
    Response.Write(sHTML)
End Sub
%>
