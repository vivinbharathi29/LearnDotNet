<%@  language="VBScript" %>
<% option explicit  %>
<html>
<head>
     <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <link rel="stylesheet" href="http://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
    <script src="http://code.jquery.com/jquery-2.2.4.js"></script>
    <script src="http://code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
	<title>Observation Reports</title>
	<script type="text/javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
	<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
	<script type="text/javascript" src="../../Includes/Client/Common.js"></script>
	<script type="text/javascript" src="http://cdnjs.cloudflare.com/ajax/libs/json3/3.3.2/json3.min.js"></script>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
    var MouseX, MouseY;

    function SaveMouseCoordinates() {
        MouseX = event.clientX;
        MouseY = event.clientY;
    }

    function ChangeColor() {
        if (frmMain.cboColor.selectedIndex > 0)
            document.body.bgColor = frmMain.cboColor.options[frmMain.cboColor.selectedIndex].text;
    }

    function cboDateOpenedCompare_onchange() {
        if (frmMain.cboDateOpenedCompare.selectedIndex == 0) {
            frmMain.txtDateOpenedDays.value = "";
        }

        if (frmMain.cboDateOpenedCompare.selectedIndex == 4) {
            spnDateOpenedCount.style.display = "none";
            spnDateOpenedRange.style.display = "";
            frmMain.txtDateOpenedRange1.focus();
        }
        else {
            spnDateOpenedCount.style.display = "";
            spnDateOpenedRange.style.display = "none";
            frmMain.txtDateOpenedDays.focus();
        }
    }

    function cboGraphScaleType_onchange() {
        if (frmMain.cboGraphScaleType.selectedIndex == 1)
            spnGraphScale.style.display = "";
        else
            spnGraphScale.style.display = "none";
    }

    function cboDaysOpenCompare_onchange() {
        if (frmMain.cboDaysOpenCompare.selectedIndex == 0) {
            frmMain.txtDaysOpenDays.value = "";
            frmMain.txtDaysOpenRange.value = "";
        }

        if (frmMain.cboDaysOpenCompare.selectedIndex == 4)
            spnDaysOpenRange.style.display = "";
        else
            spnDaysOpenRange.style.display = "none";
        frmMain.txtDaysOpenDays.focus();
    }

    function cboDateClosedCompare_onchange() {
        if (frmMain.cboDateClosedCompare.selectedIndex == 0) {
            frmMain.txtDateClosedDays.value = "";
        }

        if (frmMain.cboDateClosedCompare.selectedIndex == 4) {
            spnDateClosedCount.style.display = "none";
            spnDateClosedRange.style.display = "";
            frmMain.txtDateClosedRange1.focus();
        }
        else {
            spnDateClosedCount.style.display = "";
            spnDateClosedRange.style.display = "none";
            frmMain.txtDateClosedDays.focus();
        }
    }

    function cboDaysOwnerCompare_onchange() {
        if (frmMain.cboDaysOwnerCompare.selectedIndex == 0) {
            frmMain.txtDaysOwnerDays.value = "";
            frmMain.txtDaysOwnerRange.value = "";
        }

        if (frmMain.cboDaysOwnerCompare.selectedIndex == 4)
            spnDaysOwnerRange.style.display = "";
        else
            spnDaysOwnerRange.style.display = "none";
        frmMain.txtDaysOwnerDays.focus();
    }

    function cboDateModifiedCompare_onchange() {
        if (frmMain.cboDateModifiedCompare.selectedIndex == 0) {
            frmMain.txtDateModifiedDays.value = "";
        }

        if (frmMain.cboDateModifiedCompare.selectedIndex == 4) {
            spnDateModifiedCount.style.display = "none";
            spnDateModifiedRange.style.display = "";
            frmMain.txtDateModifiedRange1.focus();
        }
        else {
            spnDateModifiedCount.style.display = "";
            spnDateModifiedRange.style.display = "none";
            frmMain.txtDateModifiedDays.focus();
        }
    }

    function cboDaysStateCompare_onchange() {
        if (frmMain.cboDaysStateCompare.selectedIndex == 0) {
            frmMain.txtDaysStateDays.value = "";
            frmMain.txtDaysStateRange.value = "";
        }

        if (frmMain.cboDaysStateCompare.selectedIndex == 4)
            spnDaysStateRange.style.display = "";
        else
            spnDaysStateRange.style.display = "none";
        frmMain.txtDaysStateDays.focus();
    }

    function cboTargetDateCompare_onchange() {
        if (frmMain.cboTargetDateCompare.selectedIndex == 0) {
            frmMain.txtTargetDateDays.value = "";
        }

        if (frmMain.cboTargetDateCompare.selectedIndex == 4) {
            spnTargetDateCount.style.display = "none";
            spnTargetDateRange.style.display = "";
            frmMain.txtTargetDateRange1.focus();
        }
        else {
            spnTargetDateCount.style.display = "";
            spnTargetDateRange.style.display = "none";
            frmMain.txtTargetDateDays.focus();
        }
    }

    function cboProfile_onchange() {
        var strColumns;
        var strProducts;
        var strBuffer;
        var i;
        var strHeader;

        ProfileOptionsAdd.style.display = "none";
        ProfilePageLayout.style.display = "none";
        ProfileOptionsUpdate.style.display = "none";
        ProfileOptionsDelete.style.display = "none";
        ProfileOptionsRename.style.display = "none";
        ProfileOptionsOwner.style.display = "none";
        ProfileOptionsRemove.style.display = "none";
        ProfileOptionsShare.style.display = "none";

        FilterLoadingMessage.style.display = "";
        FilterLoadingMessage.style.width = FilterLoadingMessage.scrollWidth + 10;
        FilterLoadingMessage.style.height = FilterLoadingMessage.scrollHeight;
        FilterLoadingMessage.style.left = 200;
        FilterLoadingMessage.style.top = 76;

        if (cboProfile.selectedIndex > 0) {
            //			window.location.href = "../ots/default_mattH.asp?ProfileID=" + cboProfile.options[cboProfile.selectedIndex].value;
            window.location.href = "../ots/default.asp?ProfileID=" + cboProfile.options[cboProfile.selectedIndex].value;
        }
        else {
            //			window.location.href = "../ots/default_mattH.asp";
            window.location.href = "../ots/default.asp";
        }
    }

    function ActionCell_onmouseover(e, bkgColor) {
        if (bkgColor === undefined) {
            bkgColor = "gainsboro";
        }
        var MyEvent;
        if (document.all) //IE
            MyEvent = e.srcElement;
        else
            MyEvent = e.target;

        window_onmouseup();

        MyEvent.style.background = bkgColor;
        MyEvent.style.cursor = "pointer";
        MyEvent.style.color = "black";
    }

    function ActionCell_onmouseout(e, bkgColor) {
        if (bkgColor === undefined) {
            bkgColor = "#333333";
        }
        var MyEvent;
        if (document.all) //IE
            MyEvent = e.srcElement;
        else
            MyEvent = e.target;

        window_onmouseup();
        MyEvent.style.color = "white";
        MyEvent.style.background = bkgColor;
    }

    function MenuCell_onmouseover(e, MenuID) {
        if (document.all) //IE
            MyEvent = e.srcElement;
        else
            MyEvent = e.target;
        window_onmouseup();
        MyEvent.style.background = "gainsboro";
        MyEvent.style.color = "black";

        ShowMenu(MenuID, e);
    }

    function MenuCell_onmouseout(e) {
        if (document.all) //IE
            MyEvent = e.srcElement;
        else
            MyEvent = e.target;
        MyEvent.style.color = "white";
        MyEvent.style.background = "#333333";
    }

    function window_onmouseup() {
        if (typeof (mnuPopup) != "undefined")
            mnuPopup.style.display = "none";
    }

    function findPosY(obj) {
        var curtop = 0;

        if (obj.offsetParent) {
            do {
                curtop += obj.offsetTop;
            } while (obj = obj.offsetParent);
            return curtop;
        }
    }

    function findPosX(obj) {
        var curleft = 0;

        if (obj.offsetParent) {
            do {
                curleft += obj.offsetLeft;
            } while (obj = obj.offsetParent);
            return curleft
        }
    }

    function ShowMenu(MenuID, e) {
        var MyEvent;
        var SourceElement;
        var NewLeft;
        var NewTop;
        if (window.event) {
            MyEvent = window.event;
            SourceElement = window.event.srcElement;
            NewLeft = document.body.scrollLeft + ((MyEvent.clientX - MyEvent.offsetX) - 3);
            NewTop = (document.body.scrollTop - 1) + (MyEvent.clientY - MyEvent.offsetY) + SourceElement.offsetHeight;
        }
        else {
            MyEvent = e;
            SourceElement = e.target;
            NewLeft = findPosX(MyEvent.target);
            NewTop = findPosY(MyEvent.target) + SourceElement.offsetHeight;
        }

        if (typeof (mnuPopup) != "undefined") {
            mnuPopup.style.display = "";
            // mnuPopup.style.width = "140px" //mnuPopup.scrollWidth;
            //mnuPopup.style.height = mnuPopup.scrollHeight;
            mnuPopup.style.left = NewLeft;
            mnuPopup.style.top = NewTop;
        }
    }

    function PickDate(intControl) {
        var MyTextBox;
        if (intControl == 1)
            MyTextBox = frmMain.txtDateOpenedRange1;
        else if (intControl == 2)
            MyTextBox = frmMain.txtDateOpenedRange2;
        else if (intControl == 3)
            MyTextBox = frmMain.txtDateClosedRange1;
        else if (intControl == 4)
            MyTextBox = frmMain.txtDateClosedRange2;
        else if (intControl == 5)
            MyTextBox = frmMain.txtDateModifiedRange1;
        else if (intControl == 6)
            MyTextBox = frmMain.txtDateModifiedRange2;
        else if (intControl == 7)
            MyTextBox = frmMain.txtTargetDateRange1;
        else if (intControl == 8)
            MyTextBox = frmMain.txtTargetDateRange2;

        var strDate;
        strDate = window.showModalDialog("../../MobileSE/Today/calDraw1.asp", MyTextBox.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strDate) != "undefined") {
            MyTextBox.value = strDate;
        }
    }

    function ShareProfile() {
        var strResult;
        strResult = window.showModalDialog("../common/ProfileShare.asp?ID=" + cboProfile.options[cboProfile.selectedIndex].value, "", "dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    }

    function RemoveProfile() {
        if (window.confirm("Are you sure you want to stop receiving this shared profile?")) {
            frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].SharingID;
            frmMain.txtNewProfileName.value = "";
            frmMain.txtProfileUpdateType.value = "5";
            frmMain.target = "ProfileFrame";
            frmMain.action = "../common/UpdateProfile.asp"
            frmMain.submit();
        }
    }

    function RenameProfile() {
        var strNewName;
        strNewName = window.prompt("Enter new name for this profile.", cboProfile.options[cboProfile.selectedIndex].text);

        if (strNewName != null) {
            frmMain.txtNewProfileName.value = strNewName;
            frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
            frmMain.txtProfileUpdateType.value = "3";
            frmMain.target = "ProfileFrame";
            frmMain.action = "../common/UpdateProfile.asp"
            frmMain.submit();
        }
    }

    function AddProfile() {
        var strID = new Array();
        txtReturnValue.value = "";
        txtReturnValue2.value = "";
        txtReturnValue3.value = "";
        strID = window.showModalDialog("../Common/ProfileProperties.asp", "", "dialogWidth:655px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        if (typeof (strID) != "undefined" || txtReturnValue.value != "") {
            frmMain.txtProfileUpdateType.value = "1";
            frmMain.txtProfileType.value = "6";
            if (txtReturnValue.value != "") {
                frmMain.txtNewProfileName.value = txtReturnValue.value;
                frmMain.txtNewTodayLink.value = txtReturnValue2.value;
                frmMain.txtNewReportFormat.value = txtReturnValue3.value;
            }
            else {
                frmMain.txtNewProfileName.value = strID[0];
                frmMain.txtNewTodayLink.value = strID[1];
                frmMain.txtNewReportFormat.value = strID[2];
            }
            frmMain.txtProfileUpdateID.value = "0";
            frmMain.target = "ProfileFrame";
            frmMain.action = "../common/UpdateProfile.asp"
            frmMain.submit();
        }
    }

    function UpdateProfile() {
        var strID = new Array();
        txtReturnValue.value = "";
        txtReturnValue2.value = "";
        txtReturnValue3.value = "";
        strID = window.showModalDialog("../Common/ProfileProperties.asp?ID=" + cboProfile.options[cboProfile.selectedIndex].value, "", "dialogWidth:655px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
        if (typeof (strID) != "undefined" || txtReturnValue.value != "") {
            frmMain.txtProfileUpdateType.value = "2";
            if (txtReturnValue.value != "") {
                frmMain.txtNewProfileName.value = txtReturnValue.value;
                frmMain.txtNewTodayLink.value = txtReturnValue2.value;
                frmMain.txtNewReportFormat.value = txtReturnValue3.value;
            }
            else {
                frmMain.txtNewProfileName.value = strID[0];
                frmMain.txtNewTodayLink.value = strID[1];
                frmMain.txtNewReportFormat.value = strID[2];
            }
            frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
            frmMain.target = "ProfileFrame";
            frmMain.action = "../common/UpdateProfile.asp"
            frmMain.submit();
        }
    }

    function DeleteProfile() {
        if (window.confirm("Are you sure you want to delete this profile?")) {
            frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
            frmMain.txtNewProfileName.value = "";
            frmMain.txtProfileUpdateType.value = "4";
            frmMain.target = "ProfileFrame";
            frmMain.action = "../common/UpdateProfile.asp"
            frmMain.submit();
        }
    }

    function ProfileSaved(strType, strID, strResult, strError) {
        if (strType == "3") {
            if (strResult == "0")
                alert("Error Renaming Profile: " + strError);
            else {
                cboProfile.options[cboProfile.selectedIndex].text = frmMain.txtNewProfileName.value;
                //alert("Profile Renamed.");
            }
        }
        else if (strType == "4") {
            if (strResult == "0")
                alert("Error Deleting Profile: " + strError);
            else {
                cboProfile.options[cboProfile.selectedIndex] = null;
                cboProfile.selectedIndex = 0;
                //alert("Profile Deleted.");
                cboProfile_onchange();
            }
        }
        else if (strType == "5") {
            if (strResult == "0")
                alert("Error Removing Profile: " + strError);
            else {
                cboProfile.options[cboProfile.selectedIndex] = null;
                cboProfile.selectedIndex = 0;
                //alert("Profile Removed.");
                cboProfile_onchange();
            }
        }
        else if (strType == "1") {
            if (strResult == "0")
                alert("Error Adding Profile: " + strError);
            else {
                cboProfile.options[cboProfile.length] = new Option(frmMain.txtNewProfileName.value, strResult);
                // alert("Profile Added.");
                //				window.location.href = "../ots/default_mattH.asp?ProfileID=" + strResult;
                window.location.href = "../ots/default.asp?ProfileID=" + strResult;
                //cboProfile.options[cboProfile.length - 1].selected = true;
                //cboProfile.options[cboProfile.length - 1].CanEdit = "True";
                //cboProfile.options[cboProfile.length - 1].CanDelete = "True";
                //cboProfile.options[cboProfile.length - 1].PrimaryOwner = "";
                //ProfileOptionsUpdate.style.display = "";
                //ProfileOptionsDelete.style.display = "";
                //ProfileOptionsRename.style.display = "";
                //ProfileOptionsOwner.style.display = "none";
                //ProfileOptionsRemove.style.display = "none";
                //ProfileOptionsShare.style.display = "";
            }
        }
        else {
            if (strResult == "0")
                alert("Error Updating Profile: " + strError);
            else {
                cboProfile.options[cboProfile.selectedIndex].text = frmMain.txtNewProfileName.value
                alert("Profile Updated.");
            }
        }
        frmMain.txtNewProfileName.value = "";
        frmMain.txtProfileUpdateID.value = "";
        frmMain.txtProfileUpdateType.value = "";
        frmMain.txtNewTodayLink.value = "";
        frmMain.txtNewReportFormat.value = "";
    }

    function SortColumn_onchange(ControlID) {
        if (ControlID == 1 && frmMain.cboSortColumn1.options[frmMain.cboSortColumn1.selectedIndex].text == "Search Rank" && frmMain.cboSort1Direction.selectedIndex == 0)
            frmMain.cboSort1Direction.selectedIndex = 2;
        else if (ControlID == 2 && frmMain.cboSortColumn2.options[frmMain.cboSortColumn2.selectedIndex].text == "Search Rank" && frmMain.cboSort2Direction.selectedIndex == 0)
            frmMain.cboSort2Direction.selectedIndex = 2;
        else if (ControlID == 3 && frmMain.cboSortColumn3.options[frmMain.cboSortColumn3.selectedIndex].text == "Search Rank" && frmMain.cboSort3Direction.selectedIndex == 0)
            frmMain.cboSort3Direction.selectedIndex = 2;
    }

    function SummaryReport() {
        //	    var h = document.getElementById("lstProduct");
        //	    var selectedValues = "";
        //	    for (var count = 0; count < h.options.length; count++) {
        //	        if (h.options[count].selected) {
        //	            selectedValues += "|" + h.options[count].text + "|" + h.options[count].text.length + "|" + h.options[count].text.replace(/ /g, "+") + "\n";
        //	            selectedValues += "|" + h.options[count].value + "|" + h.options[count].value.length + "|" + h.options[count].value.replace(/ /g, "+");
        //            }
        //        }
        //	    alert(selectedValues);

        if (VerifyFields(1, false)) {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank";
            frmMain.txtReportSectionParameters.value = "";
            frmMain.txtReportSections.value = "0";
            frmMain.submit();
        }
    }

    function DetailedReport() {
        if (VerifyFields(1, false)) {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank"
            frmMain.txtReportSectionParameters.value = "";
            frmMain.txtReportSections.value = "1";
            frmMain.submit();
        }
    }

    function HistoryReport() {
        if (VerifyFields(1, false)) {
            frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            //			frmMain.target = "_blank"
            frmMain.txtReportSectionParameters.value = "";
            frmMain.txtReportSections.value = "-2";
            frmMain.submit();
        }
    }

    function StatusReport(ReportID) {
        if (VerifyFields(1, false)) {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank"
            frmMain.txtReportSectionParameters.value = "";
            if (ReportID == 0)
                frmMain.txtReportSections.value = "5,6,8,4"; //Update in the Report.asp page too
            else if (ReportID == 1)
                frmMain.txtReportSections.value = "5,9,11,12,10,13,4"; //Update in the Report.asp page too
            else
                frmMain.txtReportSections.value = "#" + ReportID;
            frmMain.submit();
        }
    }

    function CustomStatusReport() {
        var strResult = new Array();
        var strRunReportOK;
        if (VerifyFields(1, true))
            strRunReportOK = "1";
        else
            strRunReportOK = "0";
        //var MyWidth = screen.width;
        var MyWidth = 930//MyWidth - (MyWidth * .1);
        //		strResult = window.showModalDialog("../common/CustomReport_mattH.asp?RunReportOK=" + strRunReportOK, "", "dialogWidth:" + MyWidth + "px;dialogHeight:600px;edge: Raised;center:Yes; help: No;resizable: Yes;scroll:No;status: No");
        strResult = window.showModalDialog("../common/CustomReport.asp?RunReportOK=" + strRunReportOK, "", "dialogWidth:" + MyWidth + "px;dialogHeight:600px;edge: Raised;center:Yes; help: No;resizable: Yes;scroll:No;status: No");
        if (typeof (strResult) != "undefined") {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank"
            frmMain.txtReportSections.value = strResult[0];
            frmMain.txtReportSectionParameters.value = strResult[1];
            frmMain.submit();
        }
        document.all.ReportMenuFrame.src = "../Common/BuildReportMenu.asp?CurrentUserID=" + frmMain.txtUserID.value;
    }

    function SiMacroReport() {
        if (VerifyFields(1, false)) {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank"
            frmMain.txtReportSectionParameters.value = "macro";
            frmMain.txtReportSections.value = "0";
            frmMain.submit();
        }
    }

    function EmailOwners() {
        if (VerifyFields(3, false)) {
            //			frmMain.action = "Report_mattH.asp"
            frmMain.action = "Report.asp"
            frmMain.target = "_blank"
            frmMain.txtReportSectionParameters.value = "";
            frmMain.txtReportSections.value = "3";
            frmMain.submit();
        }
    }

    function SelectAllFilterMessage(strFilterName) {
        return "Please select only the " + strFilterName + " that should be used to filter the observations included in this report.\r\rDo not select any " + strFilterName + " in this box if you want to see results for all " + strFilterName + "."
    }
    function VerifyFields(strReport, blnQuiet) {
        var i;
        var blnSelected = false;
        var blnAllSelected;

        //Check for minimum filters
        if (typeof (frmMain.txtObservationID) != "undefined" && frmMain.txtObservationID.value.replace(/ /g, "") != "")
            blnSelected = true;
        else {
            if (typeof (frmMain.lstProduct) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProduct.length; i++)
                    if (frmMain.lstProduct.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;
                if (blnAllSelected && frmMain.lstProduct.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Products"));
                    frmMain.lstProduct.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductAndVersion) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductAndVersion.length; i++)
                    if (frmMain.lstProductAndVersion.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;
                if (blnAllSelected && frmMain.lstProductAndVersion.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product and Versions"));
                    frmMain.lstProductAndVersion.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductFamily) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductFamily.length; i++)
                    if (frmMain.lstProductFamily.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductFamily.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product Families"));
                    frmMain.lstProductFamily.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstOwner) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstOwner.length; i++)
                    if (frmMain.lstOwner.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstOwner.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Owners"));
                    frmMain.lstOwner.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstComponent) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstComponent.length; i++)
                    if (frmMain.lstComponent.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstComponent.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Components"));
                    frmMain.lstComponent.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstCoreTeam) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstCoreTeam.length; i++)
                    if (frmMain.lstCoreTeam.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstCoreTeam.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Core Teams"));
                    frmMain.lstCoreTeam.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstOwnerGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstOwnerGroup.length; i++)
                    if (frmMain.lstOwnerGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstOwnerGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Owner Groups"));
                    frmMain.lstOwnerGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstOriginator) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstOriginator.length; i++)
                    if (frmMain.lstOriginator.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstOriginator.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Originators"));
                    frmMain.lstOriginator.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstOriginatorGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstOriginatorGroup.length; i++)
                    if (frmMain.lstOriginatorGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstOriginatorGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Originator Groups"));
                    frmMain.lstOriginatorGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstDeveloper) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstDeveloper.length; i++)
                    if (frmMain.lstDeveloper.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstDeveloper.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Developers"));
                    frmMain.lstDeveloper.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstAssigned) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstAssigned.length; i++)
                    if (frmMain.lstAssigned.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstAssigned.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Assigned Users"));
                    frmMain.lstAssigned.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstDeveloperGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstDeveloperGroup.length; i++)
                    if (frmMain.lstDeveloperGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstDeveloperGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Developer Groups"));
                    frmMain.lstDeveloperGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstTester) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstTester.length; i++)
                    if (frmMain.lstTester.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstTester.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Testers"));
                    frmMain.lstTester.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstTesterGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstTesterGroup.length; i++)
                    if (frmMain.lstTesterGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstTesterGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Tester Groups"));
                    frmMain.lstTesterGroup.focus();
                    return false;
                }
            }
            if (typeof (frmMain.lstProductGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductGroup.length; i++)
                    if (frmMain.lstProductGroup.options[i].selected && frmMain.lstProductGroup.options[i].value.substring(0, 2) == "2:")
                        blnSelected = true;
                    else if (frmMain.lstProductGroup.options[i].value.substring(0, 2) == "2:")
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Cycles"));
                    frmMain.lstProductGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstComponentPM) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstComponentPM.length; i++)
                    if (frmMain.lstComponentPM.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstComponentPM.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Component PMs"));
                    frmMain.lstComponentPM.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstComponentPMGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstComponentPMGroup.length; i++)
                    if (frmMain.lstComponentPMGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstComponentPMGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Component PM Groups"));
                    frmMain.lstComponentPMGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstComponentTestLead) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstComponentTestLead.length; i++)
                    if (frmMain.lstComponentTestLead.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstComponentTestLead.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Component Test Leads"));
                    frmMain.lstComponentTestLead.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstComponentTestLeadGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstComponentTestLeadGroup.length; i++)
                    if (frmMain.lstComponentTestLeadGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstComponentTestLeadGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Component Test Lead Groups"));
                    frmMain.lstComponentTestLeadGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductTestLead) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductTestLead.length; i++)
                    if (frmMain.lstProductTestLead.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductTestLead.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product Test Leads"));
                    frmMain.lstProductTestLead.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductTestLeadGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductTestLeadGroup.length; i++)
                    if (frmMain.lstProductTestLeadGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductTestLeadGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product Test Lead Groups"));
                    frmMain.lstProductTestLeadGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstApprover) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstApprover.length; i++)
                    if (frmMain.lstApprover.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstApprover.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Approvers"));
                    frmMain.lstApprover.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstApproverGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstApproverGroup.length; i++)
                    if (frmMain.lstApproverGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstApproverGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Approver Groups"));
                    frmMain.lstApproverGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductPM) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductPM.length; i++)
                    if (frmMain.lstProductPM.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductPM.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product PMs"));
                    frmMain.lstProductPM.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstProductPMGroup) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstProductPMGroup.length; i++)
                    if (frmMain.lstProductPMGroup.options[i].selected)
                        blnSelected = true;
                    else
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstProductPMGroup.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Product PM Groups"));
                    frmMain.lstProductPMGroup.focus();
                    return false;
                }
            }

            if (typeof (frmMain.lstAffectedProduct) != "undefined") {
                blnAllSelected = true;
                for (i = 0; i < frmMain.lstAffectedProduct.length; i++)
                    if (!frmMain.lstAffectedProduct.options[i].selected)
                        blnAllSelected = false;

                if (blnAllSelected && frmMain.lstAffectedProduct.length > 0) {
                    if (!blnQuiet)
                        window.alert(SelectAllFilterMessage("Affected Products"));
                    frmMain.lstAffectedProduct.focus();
                    return false;
                }
            }
        }
        if (typeof (frmMain.txtAdvanced) != "undefined" && frmMain.txtAdvanced.value != "")
            blnSelected = true;

        if (!blnSelected) {
            if (!blnQuiet)
                window.alert("You must select at least one Product, Product Family,Actor, Actor Group, Component, Cycle, Other Criteria, or Core Team.");
            return false;
        }

        //make sure not all items are selected for Remaining lists
        if (typeof (frmMain.lstFrequency) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstFrequency.length; i++)
                if (!frmMain.lstFrequency.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstFrequency.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Frequencies"));
                frmMain.lstFrequency.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstGatingMilestone) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstGatingMilestone.length; i++)
                if (!frmMain.lstGatingMilestone.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstGatingMilestone.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Gating Milestones"));
                frmMain.lstGatingMilestone.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstFeature) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstFeature.length; i++)
                if (!frmMain.lstFeature.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstFeature.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Features"));
                frmMain.lstFeature.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstSubsystem) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstSubsystem.length; i++)
                if (!frmMain.lstSubsystem.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstSubsystem.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Sub Systems"));
                frmMain.lstSubsystem.focus();
                return false;
            }
        }
        if (typeof (frmMain.lstState) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstState.length; i++)
                if (!frmMain.lstState.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstState.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("States"));
                frmMain.lstState.focus();
                return false;
            }
        }
        if (typeof (frmMain.lstType) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstType.length; i++)
                if (!frmMain.lstType.options[i].selected)
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstType.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Types"));
                frmMain.lstType.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstProductGroup) != "undefined") {
            blnAllSelected = true;
            blnSomeODMsInList = false;
            for (i = 0; i < frmMain.lstProductGroup.length; i++)
                if (!frmMain.lstProductGroup.options[i].selected && frmMain.lstProductGroup.options[i].value.substring(0, 2) == "1:")
                    blnAllSelected = false;
                else if (frmMain.lstProductGroup.options[i].value.substring(0, 2) == "1:")
                    blnSomeODMsInList = true;

            if (blnAllSelected && blnSomeODMsInList && frmMain.lstProductGroup.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("ODM's"));
                frmMain.lstProductGroup.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstProductGroup) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstProductGroup.length; i++)
                if (!frmMain.lstProductGroup.options[i].selected && frmMain.lstProductGroup.options[i].value.substring(0, 2) == "3:")
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstProductGroup.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Dev Centers"));
                frmMain.lstProductGroup.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstProductGroup) != "undefined") {
            blnAllSelected = true;
            for (i = 0; i < frmMain.lstProductGroup.length; i++)
                if (!frmMain.lstProductGroup.options[i].selected && frmMain.lstProductGroup.options[i].value.substring(0, 2) == "4:")
                    blnAllSelected = false;

            if (blnAllSelected && frmMain.lstProductGroup.length > 0) {
                if (!blnQuiet)
                    window.alert(SelectAllFilterMessage("Product Phases"));
                frmMain.lstProductGroup.focus();
                return false;
            }
        }

        //Notes and Subject are required for email reports.
        if (strReport == 3) {
            if (typeof (frmMain.txtNotes) == "undefined") {
                if (!blnQuiet)
                    window.alert("Email Notes field is required to use this function.");
                return false;
            }
            if (typeof (frmMain.txtSubject) == "undefined") {
                if (!blnQuiet)
                    window.alert("Email Subject field is required to use this function.");
                return false;
            }
            if (frmMain.txtSubject.value == "") {
                if (!blnQuiet)
                    window.alert("Email Subject is required to use this function.");
                frmMain.txtSubject.focus();
                return false;
            }
            if (frmMain.txtNotes.value == "") {
                if (!blnQuiet)
                    window.alert("Email notes are required.");
                frmMain.txtNotes.focus();
                return false;
            }
        }
        //Observation ID - No Letters
        //if(/[[a-zA-Z]/.test(frmMain.txtObservationID.value))
        if (frmMain.txtObservationID.value.replace(/ /g, "") != "" && !(/^ *([sS][iI][oO])*[0-9]{5,8} *( *,{1} *([sS][iI][oO])*[0-9]{5,8} *)*$/.test(frmMain.txtObservationID.value))) {
            if (!blnQuiet)
                window.alert("The Observation Numbers field can only contain a comma-separated list of observation numbers.");
            frmMain.txtObservationID.focus();
            return false;
        }
        else {
            frmMain.txtObservationID.value = frmMain.txtObservationID.value.replace(/[sS][iI][oO]/g, "");
        }

        //Sort field Validations

        if (typeof (frmMain.cboSortColumn1) != "undefined") {
            if ((frmMain.cboSortColumn1.selectedIndex != 0 && frmMain.cboSortColumn2.selectedIndex != 0 && frmMain.cboSortColumn1.selectedIndex == frmMain.cboSortColumn2.selectedIndex) || (frmMain.cboSortColumn1.selectedIndex != 0 && frmMain.cboSortColumn3.selectedIndex != 0 && frmMain.cboSortColumn1.selectedIndex == frmMain.cboSortColumn3.selectedIndex) || (frmMain.cboSortColumn2.selectedIndex != 0 && frmMain.cboSortColumn3.selectedIndex != 0 && frmMain.cboSortColumn2.selectedIndex == frmMain.cboSortColumn3.selectedIndex)) {
                if (!blnQuiet)
                    window.alert("You may not specify a column more than once in the Sort Order.");
                return false;
            }
        }

        //Checked types of entered values
        if (typeof (frmMain.cboDateOpenedCompare) != "undefined") {
            if (frmMain.cboDateOpenedCompare.selectedIndex == 4 && (frmMain.txtDateOpenedRange1.value == "" && frmMain.txtDateOpenedRange2.value == "")) {
                if (!blnQuiet)
                    window.alert("At least one date is required for Date Opened range.")
                frmMain.cboDateOpenedCompare.focus();
                return false;
            }
            if (frmMain.cboDateOpenedCompare.selectedIndex == 4 && (frmMain.txtDateOpenedRange1.value != "" && isNaN(Date.parse(frmMain.txtDateOpenedRange1.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Opened range.")
                frmMain.txDateOpenedRange1.focus();
                return false;
            }
            if (frmMain.cboDateOpenedCompare.selectedIndex == 4 && (frmMain.txtDateOpenedRange2.value != "" && isNaN(Date.parse(frmMain.txtDateOpenedRange2.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Opened range.")
                frmMain.txtDateOpenedRange2.focus();
                return false;
            }
            if (frmMain.cboDateOpenedCompare.selectedIndex >= 1 && frmMain.cboDateOpenedCompare.selectedIndex <= 3 && (frmMain.txtDateOpenedDays.value == "" || isNaN(frmMain.txtDateOpenedDays.value))) {
                if (!blnQuiet)
                    window.alert("Date Opened must be a number.")
                frmMain.txtDateOpenedDays.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboDateClosedCompare) != "undefined") {
            if (frmMain.cboDateClosedCompare.selectedIndex == 4 && (frmMain.txtDateClosedRange1.value == "" && frmMain.txtDateClosedRange2.value == "")) {
                if (!blnQuiet)
                    window.alert("At least one date is required for Date Closed range.")
                frmMain.cboDateClosedCompare.focus();
                return false;
            }
            if (frmMain.cboDateClosedCompare.selectedIndex == 4 && (frmMain.txtDateClosedRange1.value != "" && isNaN(Date.parse(frmMain.txtDateClosedRange1.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Closed range.")
                frmMain.txDateClosedRange1.focus();
                return false;
            }
            if (frmMain.cboDateClosedCompare.selectedIndex == 4 && (frmMain.txtDateClosedRange2.value != "" && isNaN(Date.parse(frmMain.txtDateClosedRange2.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Closed range.")
                frmMain.txtDateClosedRange2.focus();
                return false;
            }
            if (frmMain.cboDateClosedCompare.selectedIndex >= 1 && frmMain.cboDateClosedCompare.selectedIndex <= 3 && (frmMain.txtDateClosedDays.value == "" || isNaN(frmMain.txtDateClosedDays.value))) {
                if (!blnQuiet)
                    window.alert("Date Closed must be a number.")
                frmMain.txtDateClosedDays.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboDateModifiedCompare) != "undefined") {
            if (frmMain.cboDateModifiedCompare.selectedIndex == 4 && (frmMain.txtDateModifiedRange1.value == "" && frmMain.txtDateModifiedRange2.value == "")) {
                if (!blnQuiet)
                    window.alert("At least one date is required for Date Modified range.")
                frmMain.cboDateModifiedCompare.focus();
                return false;
            }
            if (frmMain.cboDateModifiedCompare.selectedIndex == 4 && (frmMain.txtDateModifiedRange1.value != "" && isNaN(Date.parse(frmMain.txtDateModifiedRange1.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Modified range.")
                frmMain.txDateModifiedRange1.focus();
                return false;
            }
            if (frmMain.cboDateModifiedCompare.selectedIndex == 4 && (frmMain.txtDateModifiedRange2.value != "" && isNaN(Date.parse(frmMain.txtDateModifiedRange2.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Date Modified range.")
                frmMain.txtDateModifiedRange2.focus();
                return false;
            }
            if (frmMain.cboDateModifiedCompare.selectedIndex >= 1 && frmMain.cboDateModifiedCompare.selectedIndex <= 3 && (frmMain.txtDateModifiedDays.value == "" || isNaN(frmMain.txtDateModifiedDays.value))) {
                if (!blnQuiet)
                    window.alert("Date Modified must be a number.")
                frmMain.txtDateModifiedDays.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboTargetDateCompare) != "undefined") {
            if (frmMain.cboTargetDateCompare.selectedIndex == 4 && (frmMain.txtTargetDateRange1.value == "" && frmMain.txtTargetDateRange2.value == "")) {
                if (!blnQuiet)
                    window.alert("At least one date is required for Target Date range.")
                frmMain.cboTargetDateCompare.focus();
                return false;
            }
            if (frmMain.cboTargetDateCompare.selectedIndex == 4 && (frmMain.txtTargetDateRange1.value != "" && isNaN(Date.parse(frmMain.txtTargetDateRange1.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Target Date range.")
                frmMain.txTargetDateRange1.focus();
                return false;
            }
            if (frmMain.cboTargetDateCompare.selectedIndex == 4 && (frmMain.txtTargetDateRange2.value != "" && isNaN(Date.parse(frmMain.txtTargetDateRange2.value)))) {
                if (!blnQuiet)
                    window.alert("A date is required for Target Date range.")
                frmMain.txtTargetDateRange2.focus();
                return false;
            }
            if (frmMain.cboTargetDateCompare.selectedIndex >= 1 && frmMain.cboTargetDateCompare.selectedIndex <= 3 && (frmMain.txtTargetDateDays.value == "" || isNaN(frmMain.txtTargetDateDays.value))) {
                if (!blnQuiet)
                    window.alert("Target Date must be a number.")
                frmMain.txtTargetDateDays.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboDaysOpenCompare) != "undefined") {
            if (frmMain.cboDaysOpenCompare.selectedIndex != 0 && (frmMain.txtDaysOpenDays.value == "" || isNaN(frmMain.txtDaysOpenDays.value))) {
                if (!blnQuiet)
                    window.alert("Days Open must be a number.")
                frmMain.txtDaysOpenDays.focus();
                return false;
            }
            if (frmMain.cboDaysOpenCompare.selectedIndex == 4 && (frmMain.txtDaysOpenRange.value == "" || isNaN(frmMain.txtDaysOpenRange.value))) {
                if (!blnQuiet)
                    window.alert("Days Open must be a number.")
                frmMain.txtDaysOpenRange.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboDaysStateCompare) != "undefined") {
            if (frmMain.cboDaysStateCompare.selectedIndex != 0 && (frmMain.txtDaysStateDays.value == "" || isNaN(frmMain.txtDaysStateDays.value))) {
                if (!blnQuiet)
                    window.alert("Days In State must be a number.")
                frmMain.txtDaysStateDays.focus();
                return false;
            }
            if (frmMain.cboDaysStateCompare.selectedIndex == 4 && (frmMain.txtDaysStateRange.value == "" || isNaN(frmMain.txtDaysStateRange.value))) {
                if (!blnQuiet)
                    window.alert("Days In State must be a number.")
                frmMain.txtDaysStateRange.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboDaysOwnerCompare) != "undefined") {
            if (frmMain.cboDaysOwnerCompare.selectedIndex != 0 && (frmMain.txtDaysOwnerDays.value == "" || isNaN(frmMain.txtDaysOwnerDays.value))) {
                if (!blnQuiet)
                    window.alert("Days Owner must be a number.")
                frmMain.txtDaysOwnerDays.focus();
                return false;
            }
            if (frmMain.cboDaysOwnerCompare.selectedIndex == 4 && (frmMain.txtDaysOwnerRange.value == "" || isNaN(frmMain.txtDaysOwnerRange.value))) {
                if (!blnQuiet)
                    window.alert("Days Owner must be a number.")
                frmMain.txtDaysOwnerRange.focus();
                return false;
            }
        }

        if (typeof (frmMain.cboGraphScaleType) != "undefined") {
            if (frmMain.cboGraphScaleType.selectedIndex != 0 && (frmMain.txtGraphScale.value == "" || isNaN(frmMain.txtGraphScale.value))) {
                if (!blnQuiet)
                    window.alert("Graph Scale must be a number.")
                frmMain.txtGraphScale.focus();
                return false;
            }
        }

        if (typeof (frmMain.txtSearch) != "undefined") {
            if (frmMain.txtSearch.value != "" && !frmMain.chkSearchSummary.checked && !frmMain.chkSearchDetails.checked && !frmMain.chkSearchImpact.checked && !frmMain.chkSearchReproduce.checked && ((!frmMain.chkSearchHistory.checked) || (frmMain.cboSearchType.options[frmMain.cboSearchType.selectedIndex].value == "2"))) {
                if (!blnQuiet)
                    window.alert("You must specify which fields to search when searching text.")
                frmMain.txtSearch.focus();
                return false;
            }
        }

        if (typeof (frmMain.lstAffectedProduct) != "undefined") {
            blnStateSelected = false;
            blnProductSelected = false;

            if (typeof (frmMain.lstAffectedState) != "undefined") {
                for (i = 0; i < frmMain.lstAffectedState.length; i++)
                    if (frmMain.lstAffectedState.options[i].selected)
                        blnStateSelected = true;
            }

            for (i = 0; i < frmMain.lstAffectedProduct.length; i++)
                if (frmMain.lstAffectedProduct.options[i].selected)
                    blnProductSelected = true;

            if ((!blnProductSelected) && blnStateSelected) {
                if (!blnQuiet)
                    window.alert("You must specify at least one Affected Product if you select Affected States.")
                frmMain.lstAffectedProduct.focus();
                return false;
            }
        }

        return true;
    }

    function txtAdvanced_onclick() {
        var strResult;

        strResult = window.showModalDialog("../common/BuildSQL.asp", self, "dialogWidth:800px;dialogHeight:500px;edge: Raised;center:Yes; help: No;resizable: No;status: No");

        if (typeof (strResult) != "undefined") {
            frmMain.txtAdvanced.value = strResult;
        }
    }

    function cboSearchType_onchange() {
        if (frmMain.cboSearchType.options[frmMain.cboSearchType.selectedIndex].value == 2) {
            frmMain.chkSearchHistory.disabled = true;
            spnSearchHistoryText.style.color = "gray";
        }
        else {
            frmMain.chkSearchHistory.disabled = false;
            spnSearchHistoryText.style.color = "black";
        }
    }

    function ReorderColumns() {
        var strResult;
        var strColumns = "";
        var ResultArray;
        var i;
        var strSelected = "";

        for (i = 0; i < frmMain.lstColumns.length; i++) {
            if (frmMain.lstColumns.options[i].selected)
                strSelected = strSelected + "," + frmMain.lstColumns.options[i].text;
            strSelected = strSelected + ",";

            if (strColumns == "")
                strColumns = frmMain.lstColumns.options[i].text;
            else
                strColumns = strColumns + "," + frmMain.lstColumns.options[i].text;
        }
        strResult = window.showModalDialog("../common/ReorderColumns.asp?UserSettingsID=3&lstColumns=" + strColumns, "", "dialogWidth:800px;dialogHeight:500px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strResult) != "undefined") {
            if (strResult.length > 0) {
                frmMain.lstColumns.options.length = 0;

                ResultArray = strResult.split(",");
                for (i = 0; i < ResultArray.length; i++)
                    if (strSelected.indexOf("," + ResultArray[i] + ",") > -1) {
                        frmMain.lstColumns.options[frmMain.lstColumns.length] = new Option(ResultArray[i], ResultArray[i]);
                        frmMain.lstColumns.options[i].selected = true;
                    }
                    else
                        frmMain.lstColumns.options[frmMain.lstColumns.length] = new Option(ResultArray[i], ResultArray[i]);
            }
        }
    }

    function EditPageLayout() {
        var strResult;
        var MyWidth = screen.width - (screen.width * .1);

        var url = "../common/Pagelayout.asp";
        //OpenPopUp(url, "600", MyWidth, "Page Layout", true, false, true, "dialog", "iFrameID");
        strResult = ShowPropertiesDialog(url, "Page Layout", MyWidth, 600);
        //strResult = window.showModalDialog("../common/Pagelayout.asp", self, "dialogWidth:" + MyWidth + "px;dialogHeight:600px;edge: Raised;center:Yes; help: No;resizable: Yes;scroll:No;status: No");
        //if (typeof (strResult) != "undefined") {
        //    cboProfile_onchange();
        //}
    }

    function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
        $("#modalDialog").attr("width", "98%");
        $("#modalDialog").attr("height", "98%");
        $("#modalDialog").attr("src", QueryString);
        $("#iframeDialog").dialog("option", "title", Title);
        $("#iframeDialog").dialog("open").prev(".ui-dialog-titlebar").css("background", "#85B5D9");
    }

    function ClosePropertiesDialog(strID) {

        $("#iframeDialog").dialog("close");
        //if (typeof (strID) != "undefined") window.location.reload(true);
    }

    function LodgeSelectedFiltersInNebula() {
        document.getElementById("tdNebulaLink").style.display = "none";
        document.getElementById("tdSpinner").style.display = "block";
        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/Nebula/Profile/LodgeLegacyFilters_reportSetting', true);
        xhr.setRequestHeader('Content-Type', 'application/json;charset=utf-8');
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4 && xhr.status === 200) {
                document.getElementById("tdSpinner").style.display = "none";
                document.getElementById("tdNebulaLink").style.display = "block";
                var response = JSON.parse(xhr.responseText);
                document.getElementById("NebulaLink").href = "/Nebula/Home/Index?guid=" + response.GuidOfLegacyFilters;
                /*if (response.UnsupportedLegacyFilterList.length > 0) {
                    var unsupportedFilters = response.UnsupportedLegacyFilterList.join('\r\n');									
                    alert ("Please be kindly informed that following filters:\r\n" + unsupportedFilters + " \r\nhave not been supported yet in Nebula and will be well-supported in the future. Thanks for your understanding.");
                }*/
            }
        };

        var selectedFilters = {
            reportSetting: {
                CriterionSetting: {
                    ListFilters: {},
                    DateFilters: {},
                    ChkFilters: {},
                    ObservationIdFilter: {}
                },
                OutputSetting: {
                    Settings: {}
                }
            }
        };
        var listFilterIds = [
			'lstProductAndVersion', 'lstProductFamily', 'lstProduct', 'lstAffectedProduct', 'lstAffectedState', 'lstCoreTeam', 'lstType', 'lstSubsystem', 'lstComponent', 'lstFeature', 'lstFrequency',
			'lstGatingMilestone', 'lstState', 'lstReferenceId', 'lstReviewed', 'lstODM', 'lstSupplier', 'lstOriginator', 'lstLastModifiedBy',
			'lstOwner', 'lstAssignedTester', 'lstComponentTestLead', 'lstProductTestLead', 'lstApprover', 'lstDeveloper', 'lstComponentPM', 'lstProductPM', 'lstOriginatorGroup', 'lstOwnerGroup', 'lstTesterGroup',
			'lstComponentTestLeadGroup', 'lstProductTestLeadGroup', 'lstApproverGroup', 'lstDeveloperGroup', 'lstComponentPMGroup', 'lstProductPMGroup', 'lstOriginatorManager', 'lstOwnerManager', 'lstTesterManager',
			'lstComponentTestLeadManager', 'lstProductTestLeadManager', 'lstApproverManager', 'lstDeveloperManager', 'lstComponentPMManager', 'lstProductPMManager'
        ];
        var productGroupListFilterIds = [
            'lstProductOdmGroup', 'lstProductCycleGroup', 'lstProductDevCenterGroup', 'lstProductPhaseGroup'
        ];
        var detailFilterIds = [
            'lstSeverity', 'lstStatus', 'lstTestEscape', 'lstImpact'
        ];
        var dateFilterIds = [
			'cboDateOpened', 'cboDateModified', 'cboDateClosed', 'cboTargetDate'
        ];
        var daysFilterIds = [
			'cboDaysOpen', 'cboDaysInState', 'cboDaysOwner'
        ];

        BuildProductGroupListFilter(selectedFilters.reportSetting.CriterionSetting.ListFilters, productGroupListFilterIds);

        for (i = 0; i < listFilterIds.length; ++i) {
            BuildListFilter(selectedFilters.reportSetting.CriterionSetting.ListFilters, listFilterIds[i]);
        }
        for (i = 0; i < detailFilterIds.length; ++i) {
            BuildDetailFilter(selectedFilters.reportSetting.CriterionSetting.ListFilters, detailFilterIds[i]);
        }
        for (i = 0; i < dateFilterIds.length; ++i) {
            BuildDateFilter(selectedFilters.reportSetting.CriterionSetting.DateFilters, dateFilterIds[i]);
        }
        for (i = 0; i < daysFilterIds.length; ++i) {
            BuildDaysFilter(selectedFilters.reportSetting.CriterionSetting.DateFilters, daysFilterIds[i]);
        }        
        BuildCheckboxFilter(selectedFilters.reportSetting.CriterionSetting.ChkFilters, 'chkPriority');
        BuildObservationIdFilter(selectedFilters.reportSetting.CriterionSetting.ObservationIdFilter);
        BuildCustomOutputColumn(selectedFilters.reportSetting.OutputSetting);
        xhr.send(JSON.stringify(selectedFilters));
    }

    function BuildCustomOutputColumn(outputSetting) {
        var selectedOutputCols = [];
        var select = document.getElementById('lstColumns');
        var options = select && select.options;
        if (options != null) {
            var affectedProductAndStateInserted = false;
            for (var i = 0; i < options.length; i++) {
                if (options[i].selected) {
                    var selectedOptionValue = Trim(options[i].value);                    
                    if (selectedOptionValue === 'Affected Product' || selectedOptionValue === 'Affected State') {
                        if (affectedProductAndStateInserted === false) {
                            selectedOutputCols.push('Affected Product and State');
                            affectedProductAndStateInserted = true;
                        }
                    }
                    else
                        selectedOutputCols.push(selectedOptionValue);
                }
            }
        }
        outputSetting.Settings['OutputColumns'] = selectedOutputCols;
    }

    function BuildProductGroupListFilter(filters, listFilterIds) {
        var selectedOptions = GetSelectedListFilterValues('lstProductGroup');
        var selectedOdms = [];
        var selectedCycles = [];
        var selectedDevCenters = [];
        var selectedPhases = [];

        if (selectedOptions.length > 0) {
            for (i = 0; i < selectedOptions.length; ++i) {
                if (selectedOptions[i].Value.indexOf("1:") >= 0) {
                    selectedOptions[i].Value = Trim(selectedOptions[i].Value.replace("1:", ""));
                    selectedOdms.push(selectedOptions[i]);
                }
                else if (selectedOptions[i].Value.indexOf("2:") >= 0) {
                    selectedOptions[i].Value = Trim(selectedOptions[i].Value.replace("2:", ""));
                    selectedCycles.push(selectedOptions[i]);
                }
                else if (selectedOptions[i].Value.indexOf("3:") >= 0) {
                    selectedOptions[i].Value = Trim(selectedOptions[i].Value.replace("3:", ""));
                    selectedDevCenters.push(selectedOptions[i]);
                }
                else if (selectedOptions[i].Value.indexOf("4:") >= 0) {
                    selectedOptions[i].Value = Trim(selectedOptions[i].Value.replace("4:", ""));
                    selectedPhases.push(selectedOptions[i]);
                }
            }
            for (i = 0; i < listFilterIds.length; ++i) {
                if (listFilterIds[i] == 'lstProductOdmGroup' && selectedOdms.length > 0)
                    filters[listFilterIds[i]] = selectedOdms;
                else if (listFilterIds[i] == 'lstProductCycleGroup' && selectedCycles.length > 0)
                    filters[listFilterIds[i]] = selectedCycles;
                else if (listFilterIds[i] == 'lstProductDevCenterGroup' && selectedDevCenters.length > 0)
                    filters[listFilterIds[i]] = selectedDevCenters;
                else if (listFilterIds[i] == 'lstProductPhaseGroup' && selectedPhases.length > 0)
                    filters[listFilterIds[i]] = selectedPhases;
            }
        }
    }

    function BuildListFilter(filters, listFilterId) {
        var selectedOptions = GetSelectedListFilterValues(listFilterId);
        if (selectedOptions.length > 0)
            filters[listFilterId] = selectedOptions;
    }
    function BuildDetailFilter(filters, detailFilterId) {        
        var selectedOptions = GetSelectedDetailFilterValues(detailFilterId);
        if (selectedOptions.length > 0)
            filters[detailFilterId] = selectedOptions;
    }
    function BuildDateFilter(filters, dateFilterId) {
        var selectedOptions = GetSelectedDateFilterValues(dateFilterId)
        if (selectedOptions.length > 0)
            filters[dateFilterId] = selectedOptions;
    }
    function BuildDaysFilter(filters, daysFilterId) {
        var selectedOptions = GetSelectedDaysFilterValues(daysFilterId)
        if (selectedOptions.length > 0)
            filters[daysFilterId] = selectedOptions;
    }
    function BuildCheckboxFilter(filters, chkFilterId) {
        var selectedOptions = GetPriorityValues();
        if (selectedOptions.length > 0)
            filters[chkFilterId] = selectedOptions;
    }

    function BuildObservationIdFilter(observationIdFilter) {
        var obsIds = [];
        var obsIdFilterElement = document.getElementById('txtObservationID');
        if (obsIdFilterElement != null && obsIdFilterElement.value != "") {
            var ids = obsIdFilterElement.value.split(",");
            for (i = 0; i < ids.length; i++)
                obsIds.push(Trim(ids[i]));
            observationIdFilter['txtObservationId'] = obsIds;
        }
    }

    function GetSelectedListFilterValues(filterListName) {
        var selectedOptions = [];
        var select = document.getElementById(filterListName);
        var options = select && select.options;
        if (options != null) {
            for (var i = 0; i < options.length; i++) {
                if (options[i].selected) {
                    if (Trim(options[i].value) === Trim(options[i].text))
                        selectedOptions.push({ Value: Trim(options[i].value) });
                    else
                        selectedOptions.push({ Text: Trim(options[i].text), Value: Trim(options[i].value) });
                }
            }
        }

        return selectedOptions;
    }

    function GetPriorityValues() {
        var checkedOptions = [];
        if (document.getElementById('chkP0').checked)
            checkedOptions.push({ Text: 'P0', Value: '0' });
        if (document.getElementById('chkP1').checked)
            checkedOptions.push({ Text: 'P1', Value: '1' });
        if (document.getElementById('chkP2').checked)
            checkedOptions.push({ Text: 'P2', Value: '2' });
        if (document.getElementById('chkP3').checked)
            checkedOptions.push({ Text: 'P3', Value: '3' });
        if (document.getElementById('chkP4').checked)
            checkedOptions.push({ Text: 'P4', Value: '4' });
        return checkedOptions;
    }

    function Trim(string) {
        return string.replace(/^\s+|\s+$/gm, '');
    }

    function GetSelectedDetailFilterValues(detailFilterId) {
        var detailOptions = [];
        var detailFilter = detailFilterId.replace('lst', 'cbo');
        if (detailFilter == 'cboTestEscape') detailFilter = 'cboEscape';
        if (document.getElementById(detailFilter).value != "") {
            detailOptions.push({ Value: document.getElementById(detailFilter).value });
        }
        return detailOptions;
    }
    function GetSelectedDateFilterValues(dateFilterId) {
        var dateOptions = [];
        var dateFilterItem;
        var dateFilter = dateFilterId.replace('div', 'cbo');
        if (document.getElementById(dateFilter + 'Compare').value > 0) {
            dateFilterItem = {};
            dateFilterItem.DateCompare = document.getElementById(dateFilter + 'Compare').value;
            dateFilterItem.DateFilterType = 0;
            dateFilter = dateFilter.replace('cbo', 'txt');
            if (dateFilterItem.DateCompare == 4) {
                dateFilterItem.JsonFrom = document.getElementById(dateFilter + 'Range1').value;
                dateFilterItem.JsonTo = document.getElementById(dateFilter + 'Range2').value;
            }
            else {
                var currentDate = new Date();
                var subtractDays = parseInt(document.getElementById(dateFilter + 'Days').value);
                var actualDate = new Date(currentDate.setDate(currentDate.getDate() - subtractDays));
                dateFilterItem.JsonFrom = actualDate.getFullYear() + '/' + (actualDate.getMonth() + 1) + '/' + actualDate.getDate();
            }
        }
        if (typeof (dateFilterItem) != 'undefined')
            dateOptions.push(dateFilterItem);
        return dateOptions;
    }


    function GetSelectedDaysFilterValues(daysFilterId) {
        var dateOptions = [];
        var dateFilterItem;
        var dateFilter = daysFilterId.replace('div', 'cbo');
        if (dateFilter == 'cboDaysInState') dateFilter = 'cboDaysState'; //workaround
        if (document.getElementById(dateFilter + 'Compare').value > 0) {
            dateFilterItem = {};
            dateFilterItem.DateCompare = document.getElementById(dateFilter + 'Compare').value;
            dateFilterItem.DateFilterType = 1;
            dateFilter = dateFilter.replace('cbo', 'txt');
            if (dateFilterItem.DateCompare == 4) {
                dateFilterItem.JsonFrom = document.getElementById(dateFilter + 'Days').value;
                dateFilterItem.JsonTo = document.getElementById(dateFilter + 'Range').value;
            }
            else {
                dateFilterItem.JsonFrom = document.getElementById(dateFilter + 'Days').value;
            }
        }
        if (typeof (dateFilterItem) != 'undefined')
            dateOptions.push(dateFilterItem);
        return dateOptions;
    }



    //-->
	</script>
	<style type="text/css">
		TEXTAREA
		{
			font-weight: normal;
			font-size: x-small;
			font-family: Verdana;
		}
		A:link
		{
			color: blue;
		}
		A:visited
		{
			color: blue;
		}
		A:hover
		{
			color: red;
		}
		TD.HeaderButton
		{
			font-size: 8pt;
			font-family: Verdana;
			font-weight: bold;
			color: White;
			padding: 3px;
		}
		body
		{
			background-color: lightsteelblue;
			font-size: 10pt;
			font-family: Verdana;
		}
		td
		{
			font-size: 10pt;
			font-family: Verdana;
		}
	</style>
</head>
<body onmouseup="window_onmouseup()">
	<%
	dim strMyBrowser

	strMyBrowser = Request.ServerVariables("HTTP_User_Agent")

	if instr(strMyBrowser,"MSIE") = 0 then
		if instr(strMyBrowser,"Chrome") = 0 then
			strMyBrowser= "<p style=""font-family:verdana;font-size:1;color:red"">Your browser is not supported.  Please switch to Internet Explorer or Chome.<br/></B></font>"
		else
			strMyBrowser= "<p style=""font-family:verdana;font-size:1;color:green"">Please switch to Internet Explorer if you encounter any issues.<br/></B></font>"
		end if
	else
		strMyBrowser=""
	end if

	%>
	<table style="width: 100%; padding: 0; border-spacing: 0; border-collapse: collapse">
		<tr>
			<td style="font-family: verdana; font-size: small; font-weight: bold; padding: 0">
				Observation Reports
			</td>
			<td style="float: right">
				<%=strMyBrowser%>
			</td>
		</tr>
	</table>
	<br />
	<%

	dim strColors, strcolor, strSearch,ColorArray
	strColors="Cornsilk,BlanchedAlmond,Bisque,NavajoWhite,Wheat,BurlyWood,Tan,RosyBrown,SandyBrown,Goldenrod,DarkGoldenrod,Peru,Chocolate,SaddleBrown,Sienna,Brown,Maroon,White,Snow,Honeydew,MintCream,Azure,AliceBlue,GhostWhite,WhiteSmoke,Seashell,Beige,OldLace,FloralWhite,Ivory,AntiqueWhite,Linen,LavenderBlush,MistyRose,Gainsboro,LightGrey,Silver,DarkGray,Gray,DimGray,LightSlateGray,SlateGray,DarkSlateGray,Black,IndianRed,LightCoral,Salmon,DarkSalmon,LightSalmon,Crimson,Red,FireBrick,DarkRed,Pink,LightPink,HotPink,DeepPink,MediumVioletRed,PaleVioletRed,LightSalmon,Coral,Tomato,OrangeRed,DarkOrange,Orange,Gold,Yellow,LightYellow,LemonChiffon,LightGoldenrodYellow,PapayaWhip,Moccasin,PeachPuff,PaleGoldenrod,Khaki,DarkKhaki,Lavender,Thistle,Plum,Violet,Orchid,Fuchsia,Magenta,MediumOrchid,MediumPurple,Amethyst,BlueViolet,DarkViolet,DarkOrchid,DarkMagenta,Purple,Indigo,SlateBlue,DarkSlateBlue,MediumSlateBlue,GreenYellow,Chartreuse,LawnGreen,Lime,LimeGreen,PaleGreen,LightGreen,MediumSpringGreen,SpringGreen,MediumSeaGreen,SeaGreen,ForestGreen,Green,DarkGreen,YellowGreen,OliveDrab,Olive,DarkOliveGreen,MediumAquamarine,DarkSeaGreen,LightSeaGreen,DarkCyan,Teal,Aqua,Cyan,LightCyan,PaleTurquoise,AquamarineTurquoise,MediumTurquoise,DarkTurquoise,CadetBlue,SteelBlue,LightSteelBlue,PowderBlue,LightBlue,SkyBlue,LightSkyBlue,DeepSkyBlue,DodgerBlue,CornflowerBlue,MediumSlateBlue,RoyalBlue,Blue,MediumBlue,DarkBlue,Navy,MidnightBlue"
'	strColors="Wheat,BurlyWood,Tan,Beige,Ivory,Gainsboro,LightYellow,DarkKhaki,Lavender,Thistle,DarkSeaGreen,CadetBlue,SteelBlue,LightSteelBlue,LightSkyBlue,CornflowerBlue"

	ColorArray = split (strColors,",")

	dim cnExcalibur, cnSIO, rs, strSQL, cm, p, j
	dim CurrentDomain, CurrentUser, CurrentUserID, CurrentUserDivision, CurrentUserPartner

	set cnExcalibur = server.CreateObject("ADODB.Connection")
	set cnSIO = server.CreateObject("ADODB.Connection")
	cnExcalibur.ConnectionString = Session("PDPIMS_ConnectionString")
	cnSIO.ConnectionString = "Provider=SQLOLEDB.1;Data Source=housireport01.auth.hpicorp.net;Initial Catalog=sio;User ID=Excalibur_RO;Password=sQ8be9AyqPQKEcqsa3mE;"
	on error resume next
	cnSIO.Open
	on error goto 0
	if cnSIO.errors.count > 0 then
		Response.Redirect "offline.asp"
	end if
	cnExcalibur.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	CurrentUser = lcase(Session("LoggedInUser"))
	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cnExcalibur
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

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = rs("PartnerID") & ""
	else
		Response.Redirect "../Excalibur.asp"
	end if
	rs.Close

	dim strNoticeTable,BulletinHeaderColors,BulletinBodyColors

	strNoticeTable = ""
	BulletinHeaderColors = split("SeaGreen,Firebrick,Gold,SeaGreen",",")
	BulletinBodyColors = split("Honeydew,MistyRose,LightYellow,Honeydew",",")
	rs.Open "Select * FROM Bulletins with (NOLOCK) where active=1 and OTS=1 Order By id;",cnExcalibur,adOpenForwardOnly
	do while not rs.eof
		strNoticeTable = strNoticeTable & "<table cellSpacing=0 cellPadding=2 width=""100%"" border=0 bordercolor=black>"
			if rs("Severity") > -1 and rs("Severity") < 4 then
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinHeaderColors(rs("Severity")) & ";font-family:verdana;font-size:x-small;font-weight:bold;color:white"">" &  rs("Subject") & "</TD></TR>"
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinBodyColors(rs("Severity"))& ";font-family:verdana;font-size:xx-small;color:black"">" & rs("Body") & "<br/><br/></TD></TR>"
			else
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinHeaderColors(0) & ";font-family:verdana;font-size:x-small;font-weight:bold;color:white"">" &  rs("Subject") & "</TD></TR>"
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinBodyColors(0)& ";font-family:verdana;font-size:xx-small;color:black"">" & rs("Body") & "<br/><br/></TD></TR>"
			end if
		strNoticeTable = strNoticeTable &  "</TABLE><br/>"
		rs.movenext
	loop
	rs.close

	'Load Page Layout
	dim strProfileOptions
	dim strProfilePageLayout
	dim strProfile
	dim LayoutRows
	dim strRow
	dim ColumnsArray
	dim strColumn
	dim FieldProperties
	dim i
	dim blnShowBorders
	dim blnProfileFound
	dim blnProfileCanEdit
	dim blnProfileCanDelete
	dim blnProfileCanRemove
	dim strProfilePrimaryOwner
	dim ListArray
	dim strValue
	dim ProfileData
	dim blnFound
	dim FieldArray
	dim ShowSpanCount
	dim ShowSpanRange
	dim valuepair
	dim ValueArray
	dim ShowEmailButtons
	dim ShowlayoutButton
	dim strDivisionFilter
	dim strComponentFilter
	dim DivisionFilterArray
	dim blnDefaultLayoutFound
	dim CustomStatusReports
	dim ProductPhases
'	strDivisionFilter = ""
'	strDivisionFilter = " and (MobileConsumer=1 or MobileCommercial=1) "
'	strDivisionFilter = " and (MobileConsumer=1 or MobileCommercial=1 or MobileFunctional=1) "
'	strDivisionFilter = " and (DTO=1) "
'	strComponentFilter = " and coreteamid=7 "

	strDivisionFilter = ""
	blnDefaultLayoutFound = false
	rs.open "spGetEmployeeUserSettings " & clng(currentuserid) & ",9", cnExcalibur
	if not (rs.eof and rs.bof) then
		strDivisionFilter = trim(rs("Setting") & "")
	end if
	rs.Close
	if trim(strDivisionFilter) = "" then
		if trim(CurrentUserDivision) = "1" then
			strDivisionFilter = " and (MobileConsumer=1 or MobileCommercial=1 or MobileFunctional=1) "
		else
			strDivisionFilter = " and (DTO=1) "
		end if
	else
		DivisionFilterArray = split(strDivisionFilter,"|")
		strDivisionFilter = DivisionFilterArray(0)
		if ubound(DivisionFilterArray)> 0 then
			strComponentFilter = DivisionFilterArray(1)
		end if
		blnDefaultLayoutFound = true
	end if

	blnShowBorders = 0

	blnProfileFound = false
	blnProfileCanEdit = false
	blnProfileCanDelete = false
	blnProfileCanRemove = true
	strProfilePrimaryOwner = ""

	strProfilePageLayout = ""
	strProfile = ""
	if trim(request("ProfileID")) <> "" then
		rs.open "spGetReportProfile " & clng(request("ProfileID")),cnExcalibur,adOpenStatic
		if not(rs.eof and rs.bof) then
			strProfile=trim(rs("ID"))
			strProfilePageLayout=rs("PageLayout")
		end if
		rs.Close
		ShowlayoutButton = "none"
	else
		ShowlayoutButton = ""
	end if

	rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",6",cnExcalibur,adOpenForwardOnly
	strProfileOptions = ""
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<option selected=""selected"" SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = true
			blnProfileCanDelete = true
		else
			strProfileOptions = strProfileOptions & "<Option SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",6",cnExcalibur,adOpenForwardOnly
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<Option selected=""selected"" SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = cbool(rs("CanEdit"))
			blnProfileCanDelete= cbool(rs("CanDelete"))
			strProfilePrimaryOwner = shortname(rs("PrimaryOwner"))
		else
			strProfileOptions = strProfileOptions & "<Option SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",6",cnExcalibur,adOpenForwardOnly
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<Option selected=""selected"" CanRemove=0 SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = cbool(rs("CanEdit"))
			blnProfileCanDelete= cbool(rs("CanDelete"))
			strProfilePrimaryOwner = shortname(rs("PrimaryOwner"))
			blnProfileCanRemove = false
		else
			strProfileOptions = strProfileOptions & "<Option CanRemove=0 SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	if strProfileOptions <> "" then
		strProfileOptions = "<option selected=""selected""/>" & strProfileOptions
	end if

	if strProfilePageLayout = "" or not blnProfileFound then
		rs.open "spGetEmployeeUserSettings " & clng(currentuserid) & ",8", cnExcalibur
		if not (rs.eof and rs.bof) then
			strProfilePageLayout = trim(rs("Setting") & "")
		end if
		rs.Close
		if strProfilePageLayout = "" then
			strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:150:134,54:180:134,27:180:134,20:170:134,22:170:134,31:150:134|28:150:134,55:180:134,14:180:134,7:380:134,12:150:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|42:100%:3|40:100%:0|3:100%:3"
		end if
'		response.write "<br/>strProfilePageLayout=" & strProfilePageLayout & "<br/>"
		LayoutRows = split(strProfilePageLayout,"|")
		strProfile = ""

		'Full
'		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:170:134,27:170:134,11:170:134,14:170:134,26:170:134,8:170:134,7:380:134|1:170:90,50:170:134,51:170:134,18:170:134,45:170:134,4:170:134,29:170:134,20:170:134|28:170:134,5:170:134,52:170:134,9:170:134,6:170:134,24:170:134,13:170:134,22:170:134|12:170:134,46:170:134,47:170:134,48:170:134,49:170:134,10:170:134,15:170:134,31:170:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,53:220:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
		'Duplicates
'		strProfilePageLayout = "2:100%:0|-|0|16:150:134,1:200:90,27:200:134,7:380:134|28:150:134,20:200:134,22:200:134,8:170:134,31:170:134|-|0|44:120:0,30:120:0|-|0|41:100%:0|42:100%:3"
'		LayoutRows = split("21:100%:0|0|2:100%:0|-|0|16:150:134,1:180:90,20:180:134,22:170:134,10:170:134,31:150:134|28:150:134,27:180:134,14:180:134,7:380:134,12:150:134|-|0|17:240:0,39:300:0|44:120:0,32:200:0,33:200:0|30:120:0,36:200:0,34:200:0|-|0|41:100%:0|42:100%:3|40:100%:0|3:100%:3","|")
'		LayoutRows = split("21:100%:0|0|2:100%:0|-|0|16:140:134,1:170:134,14:170:134,26:170:134,4:170:134,10:140:134,12:220:134,20:280:134|28:140:134,5:170:134,9:170:134,24:170:134,11:170:134,8:140:134,22:220:134,7:280:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:200:0,33:200:0|25:120:0,37:200:0,34:200:0|19:120:0,36:200:0,35:200:0|23:120:0,38:200:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3","|")
'		LayoutRows = split("40:100%:0|3:100%:4|-|0|16:170:134,1:170:134,12:220:134,20:270:134,20:270:134,31:120:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:200:0,33:200:0|25:120:0,37:200:0,34:200:0|19:120:0,36:200:0,35:200:0|23:120:0,38:200:0|-|0|41:100%:0|43:100%:0|42:100%:3|","|")
'		LayoutRows = split("43:100%:0|0|21:100%:0|0|2:100%:0|0|41:100%:0|0|-|0|14:170:134,26:170:134,10:140:134,20:280:134,11:220:134|5:170:134,9:170:134,8:140:134,7:280:134,22:220:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:200:0,33:200:0|25:120:0,37:200:0,34:200:0|19:120:0,36:200:0,35:200:0|23:120:0,38:200:0","|")
'		LayoutRows = split("21:100%:0|0|2:100%:0|-|0|16:170:134,27:170:134,15:170:134,14:170:134,26:170:134,7:280:134,8:170:134|1:170:134,12:170:134,18:170:134,4:170:134,29:170:134,20:280:134,10:170:134|28:170:134,5:170:134,9:170:134,24:170:134,11:170:134,22:280:134,31:170:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:200:0,33:200:0|25:120:0,37:200:0,34:200:0|19:120:0,36:200:0,35:200:0|23:120:0,38:200:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3","|")
	else
		LayoutRows = split(strProfilePageLayout,"|")
	end if

	if (left(strProfilePageLayout,3)="40:" or left(strProfilePageLayout,2)="3:" or instr(strProfilePageLayout,"|40:") > 0 or instr(strProfilePageLayout,"|3:") > 0 or instr(strProfilePageLayout,",40:") > 0 or instr(strProfilePageLayout,",3:") > 0) then
		ShowEmailButtons = ""
	else
		ShowEmailButtons = "none"
	end if

	ProfileData = ""
	if blnProfileFound then
		rs.open "Select SelectedFilters from dbo.reportprofiles  with (NOLOCK) where id = " & clng(strProfile),cnExcalibur
		if not (rs.eof and rs.bof) then
			ProfileData = rs("SelectedFilters") & ""
		end if
		rs.Close
'		response.write "<hr>" & replace(ProfileData,"&","<br/>") & "<hr>"

	end if

	dim strSavedAffectedProduct,strSavedAffectedState,strSavedProduct,strSavedProductAndVersion,strSavedProductFamily,strSavedOwner,strSavedOwnerGroup,strSavedOriginatorGroup, strSavedAssigned, strSavedDeveloperGroup
	dim strSavedProductPMGroup,strSavedTesterGroup,strSavedComponentTestLeadGroup,strSavedProductTestLeadGroup,strSavedApproverGroup,strSavedComponentPMGroup,strSavedCoreTeam
	dim strSavedTitle,strSavedComponent,strSavedState,strSavedSortColumn1,strSavedSortColumn2,strSavedSortColumn3,strSavedSubsystem,strSavedGatingMilestone,strSavedDeveloper
	dim strSavedComponentPM,strSavedProductPM,strSavedOriginator,strSavedTester,strSavedComponentTestLead,strSavedProductTestLead,strSavedApprover,strSavedObservationID
	dim strSavedTargetDateCompare,strSavedDaysOpenCompare,strSavedDaysStateCompare,strSavedDaysOwnerCompare,strSavedPriority,strSavedLargeFields,strSavedType,strSavedStatus
	dim strSavedSubject,strSavedNotes,strSavedAdvanced,strSaveColumns,strSavedDateOpenedCompare,strSavedDateClosedCompare,strSavedDateModifiedCompare
	dim strSavedFormat,strSavedSeverity,strSavedDivision, strSavedEscape, strSavedImpact,strSavedTextSearch, strSaveFrequency, strSavedFeature,strSavedProductGroup
	dim strProductGroupName
	dim strSaveSearchDetails, strSaveSearchSummary,strSaveSearchImpact,strSaveSearchReproduce,strSaveSearchHistory,strSaveSearchType, strSavedDaysOpenDays, strSavedDaysOpenRange
	dim strSavedDaysStateDays, strSavedDaysStateRange, strSavedDaysOwnerDays, strSavedDaysOwnerRange, strSavedDateOpenedDays
	dim strSavedDateOpenedRange1, strSavedDateOpenedRange2,strSavedDateClosedDays,strSavedDateClosedRange1,strSavedDateClosedRange2,strSavedDateModifiedDays,strSavedDateModifiedRange1
	dim strSavedDateModifiedRange2,strSavedTargetDateDays,strSavedTargetDateRange1,strSavedTargetDateRange2, strSavedSort1Direction, strSavedSort2Direction, strSavedSort3Direction, strSavedGraphScaleType, strSavedGraphScale

	dim strLastOwner
	'*********SCRUB SQL************
	strSavedAffectedProduct = GetValue("lstAffectedProduct")
	strSavedAffectedState = GetValue("lstAffectedState")
	strSavedProduct = GetValue("lstProduct")
	strSavedProductAndVersion = GetValue("lstProductAndVersion")
	strSavedProductFamily = getvalue("lstProductFamily")
	strSavedOwner = getvalue("lstOwner")
	strSavedOwnerGroup = getvalue("lstOwnerGroup")
	strSavedOriginatorGroup = getvalue("lstOriginatorGroup")
	strSavedDeveloperGroup = getvalue("lstDeveloperGroup")
	strSavedProductPMGroup = getvalue("lstProductPMGroup")
	strSavedTesterGroup = getvalue("lstTesterGroup")
	strSavedComponentTestLeadGroup = getvalue("lstComponentTestLeadGroup")
	strSavedProductTestLeadGroup = getvalue("lstProductTestLeadGroup")
	strSavedApproverGroup= getvalue("lstApproverGroup")
	strSavedComponentPMGroup = getvalue("lstComponentPMGroup")
	strSavedCoreTeam = getvalue("lstCoreTeam")
	strSavedProductGroup = getvalue("lstProductGroup")
	strSavedTitle = getvalue("txtTitle")
	strSavedComponent = getvalue("lstComponent")
	strSavedState = getvalue("lstState")
	strSavedSubsystem = getvalue("lstSubsystem")
	strSavedGatingMilestone = getvalue("lstGatingMilestone")
	strSavedDeveloper = getvalue("lstDeveloper")
	strSavedAssigned = getvalue("lstAssigned")
	strSavedComponentPM = getvalue("lstComponentPM")
	strSavedProductPM = getvalue("lstProductPM")
	strSavedOriginator = getvalue("lstOriginator")
	strSavedTester = getvalue("lstTester")
	strSavedComponentTestLead = getvalue("lstComponentTestLead")
	strSavedProductTestLead = getvalue("lstProductTestLead")
	strSavedApprover = getvalue("lstApprover")
	strSavedObservationID = getvalue("txtObservationID")
	strSavedSubject = getvalue("txtSubject")
	strSavedNotes = getvalue("txtNotes")
	strSavedAdvanced = getvalue("txtAdvanced")
	strSavedDateOpenedCompare = getvalue("cboDateOpenedCompare")
	strSavedDateOpenedDays = getvalue("txtDateOpenedDays")
	strSavedDateOpenedRange1 = getvalue("txtDateOpenedRange1")
	strSavedDateOpenedRange2 = getvalue("txtDateOpenedRange2")
	strSavedDateClosedCompare = getvalue("cboDateClosedCompare")
	strSavedDateClosedDays = getvalue("txtDateClosedDays")
	strSavedDateClosedRange1 = getvalue("txtDateClosedRange1")
	strSavedDateClosedRange2 = getvalue("txtDateClosedRange2")
	strSavedDateModifiedCompare = getvalue("cboDateModifiedCompare")
	strSavedDateModifiedDays = getvalue("txtDateModifiedDays")
	strSavedDateModifiedRange1 = getvalue("txtDateModifiedRange1")
	strSavedDateModifiedRange2 = getvalue("txtDateModifiedRange2")
	strSavedTargetDateCompare = getvalue("cboTargetDateCompare")
	strSavedTargetDateDays = getvalue("txtTargetDateDays")
	strSavedTargetDateRange1 = getvalue("txtTargetDateRange1")
	strSavedTargetDateRange2 = getvalue("txtTargetDateRange2")
	strSavedGraphScaleType = getvalue("cboGraphScaleType")
	strSavedGraphScale = getvalue("txtGraphScale")

	strSavedSortColumn1 = getvalue("cboSortColumn1")
	strSavedSort1Direction = getvalue("cboSort1Direction")
	strSavedSortColumn2 = getvalue("cboSortColumn2")
	strSavedSort2Direction = getvalue("cboSort2Direction")
	strSavedSortColumn3 = getvalue("cboSortColumn3")
	strSavedSort3Direction = getvalue("cboSort3Direction")

	strSavedDaysOpenCompare = getvalue("cboDaysOpenCompare")
	strSavedDaysOpenDays = getvalue("txtDaysOpenDays")
	strSavedDaysOpenRange = getvalue("txtDaysOpenRange")
	strSavedDaysStateCompare = getvalue("cboDaysStateCompare")
	strSavedDaysStateDays = getvalue("txtDaysStateDays")
	strSavedDaysStateRange = getvalue("txtDaysStateRange")
	strSavedDaysOwnerCompare = getvalue("cboDaysOwnerCompare")
	strSavedDaysOwnerDays = getvalue("txtDaysOwnerDays")
	strSavedDaysOwnerRange = getvalue("txtDaysOwnerRange")
	strSavedPriority = getvalue("chkPriority")
	strSavedLargeFields = getvalue("txtLargeFieldLimit")
	strSavedType = getvalue("lstType")
	strSavedStatus = getValue("cboStatus")
	strSavedFormat = getvalue("cboFormat")
	strSavedSeverity = getValue("cboSeverity")
	strSavedDivision = getvalue("cboDivision")
	strSaveColumns = getvalue("lstColumns")
	strSavedEscape = getValue("cboEscape")
	strSavedImpact = getValue("cboImpact")
	strSavedTextSearch = getvalue("txtSearch")
	strSaveSearchType = getvalue("cboSearchType")
	strSaveSearchSummary = getvalue("chkSearchSummary")
	strSaveSearchDetails = getvalue("chkSearchDetails")
	strSaveSearchImpact = getvalue("chkSearchImpact")
	strSaveSearchReproduce = getvalue("chkSearchReproduce")
	strSaveSearchHistory = getvalue("chkSearchHistory")
	strSaveFrequency = getvalue("lstFrequency")
	strSavedFeature = getvalue("lstFeature")

	dim strColumnNames
	dim MasterColumnList
	dim MasterColumnArray
	dim strSortColumnNames
	dim strColumnName

	MasterColumnList = "Affected Product,Affected State,Approval Check,Approver,Approver Email,Approver Group,Approver Location,Approver Manager,Closed In Version,Component,Component PartNo,Component PM,Component PM Email,Component PM Group,Component PM Location,Component PM Manager,Component Test Lead,Component Test Lead Email,Component Test Lead Group,Component Test Lead Location,Component Test Lead Manager,Component Type,Component Version,Core Team,Count Component Assignments,Count Owner Assignments,Customer Impact,Date Opened,Date Closed,Date Modified,Days Current Owner,Days In State,Days Open,Developer,Developer Email,Developer Group,Developer Location,Developer Manager,Division,EA Date,EA Number,EA Status,Earliest Product Milestone,Failed Fixes,Feature,Frequency,Gating Milestone,Impacts,Implementation Check,Last Modified By,Last Release Tested,Localization,Long Description,Observation ID,ODMs,On Board,Originator,Originator Email,Originator Group,Originator Location,Originator Manager,Owner,Owner Email,Owner Group,Owner Location,Owner Manager,Priority,Primary Product,Product Family,Product PM,Product PM Email,Product PM Group,Product PM Location,Product PM Manager,Product Segment,Product Test Lead,Product Test Lead Email,Product Test Lead Group,Product Test Lead Location,Product Test Lead Manager,Reference Number,Release Fix Implemented,Reviewed,SA Part Number,Severity,Short Description,State,Status,Steps To Reproduce,Sub System,Suppliers,Supplier Version,Target Date,Test Escape,Test Procedure,Tester,Tester Email,Tester Group,Tester Location,Tester Manager,Updates"
	rs.open "spGetEmployeeUserSettings " & clng(CurrentUserID) & ",7",cnExcalibur,adOpenKeyset
	if rs.eof and rs.bof then
		strColumnNames = MasterColumnList
	elseif trim(rs("Setting")) & "" = "" then
		strColumnNames = MasterColumnList
	else
		strColumnNames = rs("Setting") & ""
		MasterColumnArray = split(MasterColumnList,",")
		for each strColumnName in MasterColumnArray
			if instr("," & lcase(trim(strColumnNames)) & ",","," & lcase(trim(strColumnName)) & "," ) = 0 then
				strColumnNames = strColumnNames & "," & trim(strColumnName)
			end if
		next
	end if
	rs.Close

	strSortColumnNames = "Affected Product,Affected State,Approval Check,Approver,Approver Group,Closed In Version,Component,Component PM,Component PM Group,Component Test Lead,Component Test Lead Group,Component Type,Component Version,Core Team,Customer Impact,Date Opened,Date Closed,Date Modified,Days Current Owner,Days In State,Days Open,Developer,Developer Group,Division,EA Date,EA Number,EA Status,Earliest Product Milestone,Failed Fixes,Feature,Frequency,Gating Milestone,Impacts,Implementation Check,Last Modified By,Last Release Tested,Localization,Observation ID,On Board,Originator,Originator Group,Owner,Owner Group,Priority,Primary Product,Product Family,Product PM,Product PM Group,Product Test Lead,Product Test Lead Group,Reference Number,Release Fix Implemented,SA Part Number,Search Rank,Severity,Short Description,State,Status,Sub System,Supplier Version,Target Date,Test Escape,Test Procedure,Tester,Tester Group"

	dim ProfileDisplayUpdateLink, ProfileDisplayDeleteLink , ProfileDisplayRenameLink,ProfileDisplayRemoveLink, ProfileDisplayOwnerLink, ProfileDisplayShareLink
	if strProfile = "" then
		ProfileDisplayUpdateLink = "none"
		ProfileDisplayDeleteLink = "none"
		ProfileDisplayRenameLink = "none"
		ProfileDisplayRemoveLink = "none"
		ProfileDisplayOwnerLink = "none"
		ProfileDisplayShareLink  = "none"
	else
		if blnProfileCanEdit  then
			ProfileDisplayUpdateLink = ""
			ProfileDisplayRenameLink = ""
		else
			ProfileDisplayUpdateLink = "none"
			ProfileDisplayRenameLink = "none"
		end if

		if blnProfileCanDelete  then
			ProfileDisplayDeleteLink = ""
		else
			ProfileDisplayDeleteLink = "none"
		end if

		if strProfilePrimaryOwner = "" then
			ProfileDisplayRemoveLink = "none"
			ProfileDisplayOwnerLink = "none"
			ProfileDisplayShareLink  = ""
		else
			if blnProfileCanRemove then
				ProfileDisplayRemoveLink = ""
			else
				ProfileDisplayRemoveLink = "none"
			end if
			ProfileDisplayOwnerLink = ""
			ProfileDisplayShareLink  = "none"

		end if
	end if

	CustomStatusReports = ""
	rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",7",cnExcalibur,adOpenForwardOnly
	do while not rs.EOF
		CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'"" style=""font-family:arial;font-size:x-small""><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></DIV>"
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",7",cnExcalibur,adOpenForwardOnly
	do while not rs.EOF
		CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'"" style=""font-family:arial;font-size:x-small""><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></DIV>"
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",7",cnExcalibur,adOpenForwardOnly
	do while not rs.EOF
		CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'"" style=""font-family:arial;font-size:x-small""><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></DIV>"
		rs.MoveNext
	loop
	rs.Close
if CustomStatusReports <> "" then
	CustomStatusReports = "<hr width=""95%"">" & CustomStatusReports & "<HR width=""95%"">"
else
	CustomStatusReports = "<hr width=""95%"">"
end if

	%>
	<%=strNoticeTable%>
	<table style="border: 0; width: 100%">
		<tr>
			<td style="white-space: nowrap">
				<b>Saved&nbsp;Profiles:&nbsp;</b>
				<select id="cboProfile" name="cboProfile" style="width: 400px" onchange="return cboProfile_onchange()">
					<%=strProfileOptions%>
				</select>
				<span id="ProfileOptionsAdd" style="font-family: verdana; font-size: xx-small"><a href="javascript:AddProfile();">Add</a></span> <span style="display: <%=ShowlayoutButton%>; white-space: nowrap; font-family: verdana; font-size: xx-small" id="ProfilePageLayout"><a href="javascript:EditPageLayout();">Page&nbsp;Layout</a>
					<%if not blnDefaultLayoutFound then %>
					<span style="color: red">
						<img style="margin-bottom: -3px; height: 15px" src="../../images/arrowleft.gif" alt="arrow" />
						Define your default layout for this page.</span>
					<%end if%>
				</span><span style="display: <%=ProfileDisplayUpdateLink%>; font-family: verdana; font-size: xx-small" id="ProfileOptionsUpdate"><a href="javascript:UpdateProfile();">Update</a> </span><span style="display: <%=ProfileDisplayDeleteLink%>; font-family: verdana; font-size: xx-small" id="ProfileOptionsDelete"><a href="javascript:DeleteProfile();">Delete</a> </span><span style="display: <%=ProfileDisplayRenameLink%>; font-family: verdana; font-size: xx-small" id="ProfileOptionsRename"><a href="javascript:RenameProfile();">Rename</a> </span><span style="display: <%=ProfileDisplayRemoveLink%>; font-family: verdana; font-size: xx-small" id="ProfileOptionsRemove"><a href="javascript:RemoveProfile();">Remove</a> </span><span style="display: <%=ProfileDisplayShareLink%>; font-family: verdana; font-size: xx-small" id="ProfileOptionsShare"><a href="javascript:ShareProfile();">Share</a> </span><span style="display: <%=ProfileDisplayOwnerLink%>; font-family: verdana; font-size: xx-small; font-weight: bold; color: black" id="ProfileOptionsOwner">Profile Owner:
					<%=strProfilePrimaryOwner%>
				</span><span style="display: none">&nbsp;&nbsp;Select Background Color:&nbsp;<select style="display: none" id="cboColor" onchange="ChangeColor();"><option selected="selected" />
					<%
'					for each strColor in ColorArray
'						response.write "<option>" & strColor & "</option>"
'					next
					%>
				</select></span>
			</td>
		</tr>
		<tr>
			<td colspan="8">
				<hr />
			</td>
		</tr>
	</table>
	<!--<form id="frmMain" action="Report_mattH.asp" method="post" target="_blank" style="margin-top: 0; margin-bottom: 0">-->
	<form id="frmMain" action="Report.asp" method="post" target="_blank" style="margin-top: 0; margin-bottom: 0">
	<table style="border: 0; border-spacing: 2px; margin-bottom: 4px">
		<tr style="background-color: #333333" id="HeaderRow">
			<td class="HeaderButton" onmouseover="ActionCell_onmouseover(event)" onmouseout="ActionCell_onmouseout(event)" onclick="SummaryReport();">
				&nbsp;&nbsp;Summary&nbsp;Report&nbsp;&nbsp;
			</td>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="DetailedReport();">
				&nbsp;&nbsp;Detailed&nbsp;Report&nbsp;&nbsp;
			</td>
			<td class="HeaderButton" onmouseover="return MenuCell_onmouseover(event,1)" onmouseout="return MenuCell_onmouseout(event)">
				&nbsp;&nbsp;Status&nbsp;Report&nbsp;&nbsp;
			</td>
			<% if currentuserpartner=1 then %>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="HistoryReport();">
				&nbsp;&nbsp;History&nbsp;Report&nbsp;(beta)&nbsp;&nbsp;
			</td>
			<% end if %>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="SiMacroReport();">
				&nbsp;&nbsp;SI&nbsp;Macro-compatible&nbsp;Report&nbsp;(beta)&nbsp;&nbsp;
			</td>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="EmailOwners();" style="display: <%=ShowEmailButtons%>">
				&nbsp;&nbsp;Email&nbsp;Observation&nbsp;Owners...&nbsp;&nbsp;
			</td>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="frmMain.reset();">
				&nbsp;&nbsp;Reset&nbsp;Form&nbsp;&nbsp;
			</td>
			<td class="HeaderButton" onmouseover="return ActionCell_onmouseover(event)" onmouseout="return ActionCell_onmouseout(event)" onclick="LodgeSelectedFiltersInNebula();">
				&nbsp;&nbsp;Copy&nbsp;Selection&nbsp;to&nbsp;Nebula&nbsp;&nbsp;
			</td>
			<td id="tdSpinner" style="background-color:rgb(176,196,222);display:none">
				<img id="spinnerIcon" alt="" src="../../images/spinner.gif" style="width:16px;height:16px;"/>
			</td>
			<td id="tdNebulaLink" class="HeaderButton" onmouseover="return ActionCell_onmouseover(event, 'rgb(255,0,0)')" onmouseout="return ActionCell_onmouseout(event, 'rgb(255,0,0)')" style="background-color: rgb(255,0,0); display: none">								
				<a id="NebulaLink" href="/Nebula/Home/Index" style="color: rgb(255,255,255)" target="_blank">Go to Nebula</a>
			</td>
		</tr>
	</table>
	<%

	'Draw Layout
	response.write "<table bordercolor=red border=" & blnShowBorders & " cellpadding=0 cellspacing=0><tr><td>"

	response.write "<table bordercolor=green border=" & blnShowBorders & " width=""100%"" cellpadding=2 cellspacing=0>"
	for each strRow in LayoutRows
		if trim(strRow) = "-" then
			response.write "</table><table style=""width:100%"" border=" & blnShowBorders & " bordercolor=orange cellspacing=0><tr><td><hr style=""padding:0;margin-top:0;margin-bottom:0""></td></tr>"
		elseif trim(strRow) = "0" then
			response.write "</table><table width=""100%"" border=" & blnShowBorders & " bordercolor=blue cellpadding=2 cellspacing=0>"
		else
			response.write "<TR>"
			ColumnsArray = split(strRow,",")
			for i = 0 to ubound(ColumnsArray)
				FieldProperties = split(ColumnsArray(i),":")
				DrawField FieldProperties(0),FieldProperties(1),FieldProperties(2), ubound(ColumnsArray) - i
			next
			response.write "<td colspan=10>&nbsp;</td></TR>"
		end if
	next
	response.write "</table></td></tr></table>"

	sub DrawField(FieldID, FieldWidth,FieldHeight, RemainingColumnCount)
		dim SelectedArray
		Select case FieldID
		case 16: 'Primary Product
			SelectedArray = split(strSavedProduct,",")
			response.write "<td style=""width:" & FieldWidth & """>"
'			response.write "<b>Product&nbsp;Version:</b><br/>"
			response.write "<b>Primary&nbsp;Product:</b><br/>"
			response.write "<select  id=lstProduct name=lstProduct multiple style=""height:" & FieldHeight & "px;width:100%"">"

			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
			next

			strSQL = "SELECT Name " & _
					 "FROM [dbo].[List_Product] with (NOLOCK) " & _
					 "Where active=1 " & _
					 " and  MobileFunctional = 0 " & _
					 strDivisionFilter

			strSql = strSql & " order by Name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
'			response.write "<optgroup style="""" label=""-- Platforms -------------------------"">"
			do while not rs.EOF
				if not inlist(SelectedArray,rs("Name")) then
					response.write "<Option value=""" & server.HTMLEncode(rs("Name")) & """>" & server.HTMLEncode(rs("Name")) & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
'			response.write "</optgroup>"

			if instr(strDivisionFilter,"MobileFunctional=1") > 0 then
				strSQL = "SELECT Name " & _
						 "FROM [dbo].[List_Product] with (NOLOCK) " & _
						 "Where active=1 " & _
						 " and  MobileFunctional = 1 " & _
						 " order by name;"

				rs.Open strSQL,cnSIO,adOpenForwardOnly
				response.write "<optgroup style="""" label=""Functional Test -------------------------""></optgroup>"
				if not inlist(SelectedArray,"Any Functional Test") then
					response.write "<Option>Any Functional Test</OPTION>"
				end if
				do while not rs.EOF
					if not inlist(SelectedArray,rs("Name")) then
						response.write "<Option value=""" & server.HTMLEncode(rs("Name")) & """>" & server.HTMLEncode(rs("Name")) & "</OPTION>"
					end if
					rs.MoveNext
				loop
				rs.Close
'				response.write "</optgroup>"
			end if
			response.write "</select></td>"
		case 1: 'Affected Product - Product and state combo - NOT USED ANYMORE
			SelectedArray = split(strSavedAffectedState,",")
			response.write "<td style=""width:" & FieldWidth & """><b>Affected&nbsp;Product:</b><br/>"
			response.write "<SELECT style=""WIDTH: 100%"" id=cboAffectedProduct name=cboAffectedProduct>"
			response.write "<Option selected=""selected""/>"
			if strSavedAffectedProduct <> "" then
				response.write "<Option selected=""selected"" value=""" & strSavedAffectedProduct & """>" & strSavedAffectedProduct & "</OPTION>"
			end if

			strSQL = "SELECT ID, DOTSName " & _
					 "FROM ProductVersion v with (NOLOCK) " & _
					 "Where DOTSName <> '' and DOTSName is not null and v.id <> 100 " & _
					 "AND typeid in (1,3) " & _
					 " and (v.productstatusid < 5) order by dotsname"

			rs.Open strSQL,cnExcalibur,adOpenForwardOnly
			do while not rs.EOF
				if trim(lcase(strSavedAffectedProduct)) <> trim(lcase(rs("DotsName"))) then
					response.write "<Option value=""" & rs("DOTSName") & """>" & rs("DOTSName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</SELECT>"
			response.write "<div style=""margin-top: 7px"">"
			response.write "<b>Affected&nbsp;States:</b></font>"
			response.write "<SELECT style=""WIDTH: 100%; HEIGHT: " & FieldHeight & "px"" multiple  size=2 id=lstAffectedState2 name=lstAffectedState2>"
			for each strvalue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			ListArray = split("Affected,Test Required,Waiver Requested,Untested,Deferred,Fix Implemented,Fix Verified,Module/Feature Constraint,Module/Feature Dropped,Not Affected,Waiver Approved,Will Not Fix",",")

			for each strValue in ListArray
				if not inlist(SelectedArray,strValue) then
					response.write "<Option value=""" & strValue & """>" & strValue & "</Option>"
				end if
			next
			response.write "</SELECT>"
			response.write "</div>"

			response.write "</TD>"
		case 54: 'Affected Product
			SelectedArray = split(strSavedAffectedProduct,",")
			response.write "<td style=""width:" & FieldWidth & """><b>Affected&nbsp;Product:</b><br/>"
			response.write "<select  id=lstAffectedProduct name=lstAffectedProduct multiple style=""height:" & FieldHeight & "px;width:100%"">"

			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
			next

			strSQL = "SELECT Name " & _
					 "FROM [dbo].[List_Product] with (NOLOCK) " & _
					 "Where active=1 " & _
					 " and  MobileFunctional = 0 " & _
					 strDivisionFilter

			strSql = strSql & " order by Name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not inlist(SelectedArray,rs("Name")) then
					response.write "<Option value=""" & server.HTMLEncode(rs("Name")) & """>" & server.HTMLEncode(rs("Name")) & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			if instr(strDivisionFilter,"MobileFunctional=1") > 0 then
				strSQL = "SELECT Name " & _
						 "FROM [dbo].[List_Product] with (NOLOCK) " & _
						 "Where active=1 " & _
						 " and  MobileFunctional = 1 " & _
						 " order by name;"

				rs.Open strSQL,cnSIO,adOpenForwardOnly
				response.write "<optgroup style="""" label=""Functional Test -------------------------""></optgroup>"
				do while not rs.EOF
					if not inlist(SelectedArray,rs("Name")) then
						response.write "<Option value=""" & server.HTMLEncode(rs("Name")) & """>" & server.HTMLEncode(rs("Name")) & "</OPTION>"
					end if
					rs.MoveNext
				loop
				rs.Close
'				response.write "</optgroup>"
			end if

			response.write "</SELECT>"

			response.write "</TD>"
		case 55: 'Affected State
			SelectedArray = split(strSavedAffectedState,",")
			response.write "<td style=""width:" & FieldWidth & """><b>Affected&nbsp;State:</b><br/>"
			response.write "<SELECT style=""WIDTH: 100%; HEIGHT: " & FieldHeight & "px"" multiple  size=2 id=lstAffectedState name=lstAffectedState>"
			for each strvalue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			ListArray = split("Affected,Test Required,Waiver Requested,Untested,Deferred,Fix Implemented,Fix Verified,Module/Feature Constraint,Module/Feature Dropped,Not Affected,Waiver Approved,Will Not Fix",",")

			for each strValue in ListArray
				if not inlist(SelectedArray,strValue) then
					response.write "<Option value=""" & strValue & """>" & strValue & "</Option>"
				end if
			next
			response.write "</SELECT>"

			response.write "</TD>"
		case 26: 'Owner Group
			SelectedArray = split(strSavedOwnerGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Owner&nbsp;Group:</b><br/>"
			response.write "<select id=lstOwnerGroup name=lstOwnerGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedOwnerGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & scrubsql(strSavedOwnerGroup) & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if
			strSQL = "Select GroupID as ID, name as OwnerGroup " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and owner=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("OwnerGroup") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"

		case 29: 'Originator Group
			SelectedArray = split(strSavedOriginatorGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Originator&nbsp;Group:</b><br/>"
			response.write "<select id=lstOriginatorGroup name=lstOriginatorGroup  multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedOriginatorGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & scrubsql(strSavedOriginatorGroup) & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and Originator=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 27: 'Core team
			SelectedArray = split(strSavedCoreTeam,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Core&nbsp;Team:</b><br/>"
			response.write "<select id=lstCoreTeam name=lstCoreTeam multiple style=""height:" & FieldHeight & "px;width:100%"">"

			if trim(strSavedCoreTeam) <> "" then
				rs.open "Select ID, Name from dbo.DeliverableCoreTeam with (NOLOCK) where ID in (" & scrubsql(strSavedCoreTeam) & ") order by Name",cnExcalibur
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "spListDeliverableCoreTeams"
			rs.Open strSQL,cnExcalibur,adOpenForwardOnly
			do while not rs.EOF
				if rs("ID") <> 0 then
					if not inlist(SelectedArray,rs("ID")) then
						response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
					end if
				end if
				rs.MoveNext
			loop
			rs.Close
			if not inlist(SelectedArray,"0") then
				response.write "<Option value= ""0"">No Core Team Assigned</OPTION>"
			end if
			response.write "</select></td>"
		case 15: 'Product Family
			SelectedArray = split(strSavedProductFamily,",")
			response.write "<td style=""width:" & FieldWidth & """>"
'			response.write "<b>Product:</b><br/>"
			response.write "<b>Product Family:</b><br/>"
			response.write "<select id=lstProductFamily name=lstProductFamily multiple style=""height:" & FieldHeight & "px;width:100%"">"

			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
			next
			strSQL = "SELECT distinct FamilyName as productFamily " & _
					 "FROM [dbo].[List_Product] with (NOLOCK) " & _
					 "Where active=1 " & _
					 strDivisionFilter

			strSql = strSql & " order by FamilyName;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not inlist(SelectedArray,rs("ProductFamily")) then
					response.write "<Option value=""" & server.HTMLEncode(rs("ProductFamily")) & """>" & server.HTMLEncode(rs("ProductFamily")) & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 58: 'Product and Version
			SelectedArray = split(strSavedProductAndVersion,",")
			response.write "<td colspan=5 style=""width:" & FieldWidth & """>"
			response.write "<b>Product and Version:</b><br/>"
			response.write "<select id=lstProductAndVersion name=lstProductAndVersion multiple style=""height:" & FieldHeight & "px;width:100%"">"

			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
			next

			strSQL = "SELECT distinct coalesce(FamilyName, '') + '||' + coalesce(Name, '') as productAndVersion " & _
					 "FROM [dbo].[List_Product] with (NOLOCK) " & _
					 "Where active=1 " & _
					 strDivisionFilter

			strSql = strSql & " order by productAndVersion;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not inlist(SelectedArray,rs("productAndVersion")) then
					response.write "<Option value=""" & server.HTMLEncode(rs("productAndVersion")) & """>" & server.HTMLEncode(rs("productAndVersion")) & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 22: 'Subsystem
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Sub&nbsp;System:</b><br/>"
			response.write "<select id=lstSubsystem name=lstSubsystem multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedSubsystem,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			strSQL = " Select name " & _
					 "from dbo.List_SubSystem with (NOLOCK) " & _
					 "where active=1 " & _
					 strDivisionFilter

			strSql = strSql & " order by Name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("Name")) then
					response.write "<Option value=""" & rs("Name") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 12: 'Gating Milestone
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Gating&nbsp;Milestone:</b><br/>"
			response.write "<select id=lstGatingMilestone name=lstGatingMilestone multiple style=""height:" & FieldHeight & "px;width:100%"">"
			SelectedArray = split(strSavedGatingMilestone,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next
			strSQL = " Select Name " & _
					 "from dbo.List_GatingMilestone with (NOLOCK) " & _
					 "where active=1 " & _
					 strDivisionFilter

			strSql = strSql & " order by Name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("Name")) then
					if rs("Name") & "" = "" then
						if not InList(selectedarray, "Not Specified") then
							response.write "<Option>Not Specified</OPTION>"
						end if 
					else
						response.write "<Option value=""" & rs("Name") & """>" & rs("Name") & "</OPTION>"
					end if
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 57: 'Assigned To
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Assigned&nbsp;To:</b><br/>"
			response.write "<select id=lstAssigned name=lstAssigned multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedAssigned,",")
			if strSavedAssigned <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedAssigned & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 strDivisionFilter

					 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 9: 'Developer
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Developer:</b><br/>"
			response.write "<select id=lstDeveloper name=lstDeveloper multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedDeveloper,",")
			if strSavedDeveloper <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedDeveloper & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and developer=1 " & _
					 strDivisionFilter

					 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"

		case 5: 'Component PM
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Component PM:</b><br/>"
			response.write "<select id=lstComponentPM name=lstComponentPM multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedComponentPM,",")
			if strSavedComponentPM <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedComponentPM & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and componentpm=1 " & _
					 strDivisionFilter

					 strSQL = strSQL & " order by DisplayName; "

			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 18: 'ProductPM
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Product PM:</b><br/>"
			response.write "<select id=lstProductPM name=lstProductPM multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedProductPM,",")

			if strSavedProductPM <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedProductPM & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and productpm=1 " & _
					 strDivisionFilter

					 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 4: 'Originator
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Originator:</b><br/>"
			response.write "<select id=lstOriginator name=lstOriginator multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedOriginator,",")
			if strSavedOriginator <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedOriginator & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and originator=1 " & _
					 strDivisionFilter

			strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("UserID")) then
					response.write "<Option value=""" & rs("UserID") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.movenext
			loop
			rs.Close
			response.write "</select></td>"
		case 11: 'Frequency
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Frequency:</b><br/>"
			response.write "<SELECT multiple style=""height:" & FieldHeight & "px;width:100%"" id=lstFrequency name=lstFrequency>"

			ListArray = split("621|Always: 100%,622|Intermittent: <1%,623|Seen Once,780|Intermittent: 1-5%,781|Intermittent: 5-25%,782|Intermittent: 25-99%,783|Single Unit Failure,10247|Related Case",",")

			SelectedArray = split(strSaveFrequency,",")
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if inlist(SelectedArray,ValuePair(0)) then
					response.write "<Option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</OPTION>"
				end if
			next

			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if not InList(selectedarray,ValuePair(0) & "|" & ValuePair(1)) then
					response.write "<Option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 24: 'Tester
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Tester:</b><br/>"
			response.write "<select id=lstTester name=lstTester multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedTester,",")
			if strSavedTester <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedTester & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and tester=1 " & _
					 strDivisionFilter

			 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 20: 'State
			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<b>State:</b><br/>"
			response.write "<select id=lstState name=lstState multiple style=""height:" & FieldHeight & "px;width:100%"">"
			SelectedArray = split(strSavedState,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			strSQl = "Select Name " & _
					 "from dbo.list_state with (NOLOCK) " & _
					 "where active=1 " & _
					 "order by Name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("Name")) then
					response.write "<Option value=""" & rs("Name") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 7: 'Component
			response.write "<td colspan=3 style=""width:" & FieldWidth & """>"
			response.write "<b>Component:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b><br/>"
			response.write "<select id=lstComponent name=lstComponent multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedComponent,",")
			for each strValue in SelectedArray
				if instr(strValue,",")>0 then
					response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(replace(strValue,",","|")) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
				else
					response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
				end if
			next

			if instr(lcase(strComponentFilter),"linktype")= 0 then
				strSQL = " Select name " & _
						"from dbo.List_Component with (NOLOCK) " & _
						"where active=1 " & _
						"and archive=0 " & _
						 strDivisionFilter & " " & _
						strComponentFilter

			else
				strSQL = " Select distinct name " & _
						"from dbo.List_Component c with (NOLOCK), dbo.List_Component_Links l with (NOLOCK) " & _
						"where l.componentid = c.id " & _
						"and active=1 " & _
						"and archive=0 " & _
						 strDivisionFilter & " " & _
						strComponentFilter
			end if
			strSql = strSql & " order by Name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("Name")) then
					if instr(rs("name"),",")> 0 then
						response.write "<Option value=""" & server.HTMLEncode(replace(rs("Name"),",","|")) & """>" & server.HTMLEncode(rs("name")) & "</OPTION>"
					else
						response.write "<Option value=""" & server.HTMLEncode(rs("name")) & """>" & server.HTMLEncode(rs("name")) & "</OPTION>"
					end if
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 10: 'Feature
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Feature:</b><br/>"
			response.write "<select id=lstFeature name=lstFeature multiple style=""height:" & FieldHeight & "px;width:100%"">"
			SelectedArray = split(strSavedFeature,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strValue) & "</OPTION>"
			next

			strSQL = "Select Name as Feature " & _
					 "from dbo.DeliverableFeatures with (NOLOCK) " & _
					 "where Active=1 " & _
					 "order by Name"
			rs.Open strSQL,cnExcalibur,adOpenForwardOnly
			do while not rs.EOF
				if not InList(selectedarray,rs("Feature")) then
					response.write "<Option value=""" & server.HTMLEncode(rs("Feature")) & """>" & server.HTMLEncode(rs("Feature")) & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 8: 'Type
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Type:</b><br/>"
			response.write "<select id=lstType name=lstType multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedType,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			ListArray = split("CD/Docs,Certification,Factory,FW,HW,Image,Softpaq,SW",",")

			for each strValue in ListArray
				if not InList(selectedarray,strValue) then
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</select></td>"
		case 31 'Columns
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<table cellpadding=0 cellspacing=0 style=""width:100%""><tr><td><b>Columns:</b></td><td align=right><a href=""javascript: ReorderColumns();""><img src=""../../images/edit2.gif"" alt=""Reorder"" border=""0""></a></td></tr></table>"
			response.write "<SELECT style=""WIDTH: 100%; HEIGHT:" & FieldHeight & "px"" multiple  size=2 id=lstColumns name=lstColumns>"
			SelectedArray = split(strSaveColumns,",")
			for each strValue in SelectedArray
				response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
			next

			ListArray = split(strColumnNames,",")

			for each strValue in ListArray
				if not inlist(SelectedArray,strValue) then
					response.write "<Option value=""" & strValue & """>" & strValue & "</Option>"
				end if
			next

			response.write "</SELECT>"
			response.write "</td>"
		case 44: 'Report Format
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Report&nbsp;Format:</b>&nbsp;&nbsp;</td>"

			ListArray = split("HTML,Excel,Word",",")

			if strSavedFormat = "" then
				strSavedFormat = "HTML"
			end if

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboFormat name=cboFormat>"

			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedFormat)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 30: 'Status
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Status:</b>&nbsp;&nbsp;</td>"

			ListArray = split("Open,Pending EA,Closed,Not Closed",",")

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboStatus name=cboStatus>"

			response.write "<Option selected=""selected""/>"
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedStatus)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 25: 'Impact
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Impacts:</b>&nbsp;&nbsp;</td>"

			ListArray = split("Not Specified,Customer,Factory",",")

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboImpact name=cboImpact>"

			response.write "<Option selected=""selected""/>"
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedImpact)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 19: 'Severity
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Severity:</b>&nbsp;&nbsp;</td>"

			ListArray = split("1 - Critical,2 - Serious,3 - Medium,4 - Low",",")

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboSeverity name=cboSeverity>"

			response.write "<Option selected=""selected""/>"
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSeverity)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 56: 'Division
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Division:</b>&nbsp;&nbsp;</td>"

			ListArray = split("Mobile,DTO",",")

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboDivision name=cboDivision>"

			response.write "<Option selected=""selected""/>"
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedDivision)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 23: 'Test Escape
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Test&nbsp;Escape:</b>&nbsp;&nbsp;</td>"

			ListArray = split("Not Specified,Functional,Integration,Unit",",")

			response.write "<td style=""width:" & FieldWidth & "px"">"
			response.write "<SELECT style=""width:" & FieldWidth & "px"" id=cboEscape name=cboEscape>"

			response.write "<Option selected=""selected""/>"
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedEscape)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next

			response.write "</SELECT></td>"
		case 21: 'Observation Numbers
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Observation Numbers:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<INPUT id=txtObservationID name=txtObservationID style=""WIDTH: 100%; HEIGHT: 22px"" value=""" & server.HTMLEncode(strSavedObservationID) & """>"
			response.write "</td>"
		case 42: 'Other Criteria
			response.write "<td nowrap=""nowrap"" style=""width:120px"" valign=top><b>Other Criteria:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
'			if currentuserid = 31 then
				response.write "<textarea readonly id=""txtAdvanced"" name=""txtAdvanced"" onkeydown=""txtAdvanced_onclick();"" onmousedown=""txtAdvanced_onclick();"" style=""WIDTH: 100%;"" rows=" & FieldHeight & """>" & server.HTMLEncode(strSavedAdvanced) & "</textarea>"
'			else
'				response.write "<textarea id=""txtAdvanced"" name=""txtAdvanced"" style=""WIDTH: 100%;"" rows=" & FieldHeight & """>" & server.HTMLEncode(strSavedAdvanced) & "</textarea>"
'			end if
			response.write "</td>"
		case 43: 'Report Title
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Report Title:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<INPUT id=txtTitle name=txtTitle style=""WIDTH: 100%; HEIGHT: 22px"" value=""" & server.htmlencode(strSavedTitle) & """>"
			response.write "</td>"
		case 40: 'Email Subject
			response.write "<td valign=top nowrap=""nowrap"" style=""width:120px""><b>Email&nbsp;Subject:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<INPUT id=txtSubject name=txtSubject style=""WIDTH: 100%; HEIGHT: 22px"" value=""" & server.HTMLEncode(strSavedSubject) & """>"
			response.write "</td>"
		case 3: 'Email  Notes
			response.write "<td  valign=top nowrap=""nowrap"" style=""width:120px""><b>Email&nbsp;Notes:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<textarea id=""txtNotes"" name=""txtNotes"" style=""WIDTH: 100%;"" rows=" & FieldHeight & ">"  & server.HTMLEncode(strSavedNotes) & "</textarea>"
			response.write "</td>"
		case 39: 'Large Field Limit
			if trim(request("ProfileID")) = ""  then
				strSavedLargeFields = "500"
			else
				strSavedLargeFields = server.htmlencode(strSavedLargeFields)
			end if
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Large Fields:</b>&nbsp;&nbsp;</td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """>"
			response.write "Show first <input id=""txtLargeFieldLimit"" name=""txtLargeFieldLimit"" type=""text"" style=""width:40px"" value=""" & strSavedLargeFields & """> characters."
			response.write "</td>"
		case 17: 'Priority
			if trim(request("ProfileID")) = "" then
				strSavedPriority = "checked,checked,checked,checked,checked"
			else
				ValueArray = split(strSavedPriority,",")
				strSavedPriority = ""
				for j = 0 to 4
					if inlist(ValueArray,j) then
						strSavedPriority = strSavedPriority & "," & "checked"
					else
						strSavedPriority = strSavedPriority & "," & ""
					end if
				next
				strSavedPriority = mid(strSavedPriority,2)
			end if
			SelectedArray = split(strSavedPriority,",")
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Priority:</b>&nbsp;&nbsp;</td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & ";font-family:verdana;font-size:x-small"">"
			response.write "P0<INPUT " & SelectedArray(0) & " type=""checkbox"" id=chkP0 name=chkPriority value=0>&nbsp;&nbsp;"
			response.write "P1<INPUT " & SelectedArray(1) & " type=""checkbox"" id=chkP1 name=chkPriority value=1>&nbsp;&nbsp;"
			response.write "P2<INPUT " & SelectedArray(2) & " type=""checkbox"" id=chkP2 name=chkPriority value=2>&nbsp;&nbsp;"
			response.write "P3<INPUT " & SelectedArray(3) & " type=""checkbox"" id=chkP3 name=chkPriority value=3>&nbsp;&nbsp;"
			response.write "P4<INPUT " & SelectedArray(3) & " type=""checkbox"" id=chkP4 name=chkPriority value=4>&nbsp;&nbsp;"
			response.write "</td>"
		case 2: 'Text Search
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Text Search:</b>&nbsp;&nbsp;</td>"
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<input style=""width:100%"" id=""txtSearch"" name=""txtSearch"" type=""text"" value=""" & server.htmlencode(strSavedTextSearch) & """></td>"
			response.write "<td>&nbsp;<select id=""cboSearchType"" name=""cboSearchType"" onchange=""cboSearchType_onchange();"">"
			if trim(strSaveSearchType) = "1" or trim(strSaveSearchType) = "" then
				response.write "<option selected=""selected"" value=""1"">Keywords - All</option>"
				response.write "<option value=""3"">Keywords - Any</option>"
				response.write "<option value=""0"">Natural Language</option>"
				response.write "<option value=""2"">Text String</option>"
			elseif trim(strSaveSearchType) = "0" then
				response.write "<option value=""1"">Keywords - All</option>"
				response.write "<option value=""3"">Keywords - Any</option>"
				response.write "<option selected=""selected"" value=""0"">Natural Language</option>"
				response.write "<option value=""2"">Text String</option>"
			elseif trim(strSaveSearchType) = "3" then
				response.write "<option value=""1"">Keywords - All</option>"
				response.write "<option selected=""selected"" value=""3"">Keywords - Any</option>"
				response.write "<option value=""0"">Natural Language</option>"
				response.write "<option value=""2"">Text String</option>"
			else
				response.write "<option value=""1"">Keywords - All</option>"
				response.write "<option value=""3"">Keywords - Any</option>"
				response.write "<option value=""0"">Natural Language</option>"
				response.write "<option selected=""selected"" value=""2"">Text String</option>"
			end if
			response.write "</select>&nbsp;&nbsp;</td>"
			response.write "<td align=left style=""white-space:nowrap"">Look In:"
			if trim(strSaveSearchSummary) = "1" or request("ProfileID")="" then
				strSaveSearchSummary = "checked"
			else
				strSaveSearchSummary = ""
			end if
			if trim(strSaveSearchDetails) = "1" or request("ProfileID")=""then
				strSaveSearchDetails = "checked"
			else
				strSaveSearchDetails = ""
			end if
			if trim(strSaveSearchImpact) = "1" or request("ProfileID")="" then
				strSaveSearchImpact = "checked"
			else
				strSaveSearchImpact = ""
			end if
			if trim(strSaveSearchReproduce) = "1" or request("ProfileID")="" then
				strSaveSearchReproduce = "checked"
			else
				strSaveSearchReproduce = ""
			end if
			if trim(strSaveSearchHistory) = "1" then
				strSaveSearchHistory = "checked"
			else
				strSaveSearchHistory = ""
			end if
			response.write "<input " & strSaveSearchSummary & " type=""checkbox"" id=chkSearchSummary name=chkSearchSummary value=1>Short&nbsp;Desc.&nbsp;"
			response.write "<input " & strSaveSearchDetails & " type=""checkbox"" id=chkSearchDetails name=chkSearchDetails value=1>Long&nbsp;Desc.&nbsp;"
			response.write "<input " & strSaveSearchImpact & " type=""checkbox"" id=chkSearchImpact name=chkSearchImpact value=1>Impact&nbsp;"
			response.write "<input " & strSaveSearchReproduce & " type=""checkbox"" id=chkSearchReproduce name=chkSearchReproduce value=1>Reproduce&nbsp;"
			if trim(strSaveSearchType) = "2" then
				response.write "<input " & strSaveSearchHistory & " disabled type=""checkbox"" id=chkSearchHistory name=chkSearchHistory value=1><span style=""color:gray"" id=spnSearchHistoryText>History&nbsp;</span>"
			else
				response.write "<input " & strSaveSearchHistory & " type=""checkbox"" id=chkSearchHistory name=chkSearchHistory value=1><span id=spnSearchHistoryText>History&nbsp;</span>"
			end if
			response.write "</td>"
		case 32: 'Date Opened
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Date Opened:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth +20 & """>"
			if trim(strSavedDateOpenedCompare) = "" then
				strSavedDateOpenedCompare = "|||"
			else
				strSavedDateOpenedCompare = strSavedDateOpenedCompare & "|"
				strSavedDateOpenedCompare = strSavedDateOpenedCompare & strSavedDateOpenedDays & "|"
				strSavedDateOpenedCompare = strSavedDateOpenedCompare & strSavedDateOpenedRange1 & "|"
				strSavedDateOpenedCompare = strSavedDateOpenedCompare & strSavedDateOpenedRange2
			end if
			FieldArray = split(strSavedDateOpenedCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanCount = "none"
				ShowSpanRange = ""
			else
				ShowSpanCount = ""
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDateOpenedCompare"" name=""cboDateOpenedCompare"" onchange=""return cboDateOpenedCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<span ID=""spnDateOpenedCount"" style=""white-space:nowrap;display:" & ShowSpanCount & ";font-family:verdana;font-size:x-small""><input style=""width:55"" type=""text"" id=""txtDateOpenedDays"" name=""txtDateOpenedDays"" value=""" & FieldArray(1) & """> Days Ago</span>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDateOpenedRange""><input style=""width:75"" type=""text"" id=""txtDateOpenedRange1"" name=""txtDateOpenedRange1"" maxlength=25 value=""" & FieldArray(2) & """>&nbsp;<a href=""javascript:PickDate(1);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a>&nbsp;-&nbsp;<input style=""width:75"" type=""text"" id=""txtDateOpenedRange2"" name=""txtDateOpenedRange2"" maxlength=25 value=""" & FieldArray(3) & """>&nbsp;<a href=""javascript:PickDate(2);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a></span>"
			response.write "</td>"
		case 33: 'Days Open
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Days Open:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """><div style=""white-space:nowrap"">"
			if trim(strSavedDaysOpenCompare) = "" then
				strSavedDaysOpenCompare = "||"
			else
				strSavedDaysOpenCompare = strSavedDaysOpenCompare & "|"
				strSavedDaysOpenCompare = strSavedDaysOpenCompare & strSavedDaysOpenDays & "|"
				strSavedDaysOpenCompare = strSavedDaysOpenCompare & strSavedDaysOpenRange
			end if
			FieldArray = split(strSavedDaysOpenCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanRange = ""
			else
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDaysOpenCompare"" name=""cboDaysOpenCompare"" onchange=""return cboDaysOpenCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<input style=""width:55"" type=""text"" id=""txtDaysOpenDays"" name=""txtDaysOpenDays"" value=""" & FieldArray(1) & """>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDaysOpenRange"">&nbsp;-&nbsp;<input style=""width:55"" type=""text"" id=""txtDaysOpenRange"" name=""txtDaysOpenRange"" maxlength=25 value=""" & FieldArray(2) & """></span>"
			response.write "<span style=""font-family:verdana;font-size:x-small"">&nbsp;Days</span></div>"
			response.write "</td>"
		case 37: 'Date Closed
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Date Closed:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """>"
			if trim(strSavedDateClosedCompare) = "" then
				strSavedDateClosedCompare = "|||"
			else
				strSavedDateClosedCompare = strSavedDateClosedCompare & "|"
				strSavedDateClosedCompare = strSavedDateClosedCompare & strSavedDateClosedDays & "|"
				strSavedDateClosedCompare = strSavedDateClosedCompare & strSavedDateClosedRange1 & "|"
				strSavedDateClosedCompare = strSavedDateClosedCompare & strSavedDateClosedRange2
			end if
			FieldArray = split(strSavedDateClosedCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanCount = "none"
				ShowSpanRange = ""
			else
				ShowSpanCount = ""
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDateClosedCompare"" name=""cboDateClosedCompare"" onchange=""return cboDateClosedCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<span ID=""spnDateClosedCount"" style=""white-space:nowrap;display:" & ShowSpanCount & ";font-family:verdana;font-size:x-small""><input style=""width:55"" type=""text"" id=""txtDateClosedDays"" name=""txtDateClosedDays"" value=""" & FieldArray(1) & """> Days Ago</span>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDateClosedRange""><input style=""width:75"" type=""text"" id=""txtDateClosedRange1"" name=""txtDateClosedRange1"" maxlength=25 value=""" & FieldArray(2) & """>&nbsp;<a href=""javascript:PickDate(3);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a>&nbsp;-&nbsp;<input style=""width:75"" type=""text"" id=""txtDateClosedRange2"" name=""txtDateClosedRange2"" maxlength=25 value=""" & FieldArray(3) & """>&nbsp;<a href=""javascript:PickDate(4);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a></span>"
			response.write "</td>"
		case 34: 'Days In State
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Days In State:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """><div style=""white-space:nowrap"">"
			if trim(strSavedDaysStateCompare) = "" then
				strSavedDaysStateCompare = "||"
			else
				strSavedDaysStateCompare = strSavedDaysStateCompare & "|"
				strSavedDaysStateCompare = strSavedDaysStateCompare & strSavedDaysStateDays & "|"
				strSavedDaysStateCompare = strSavedDaysStateCompare & strSavedDaysStateRange
			end if
			FieldArray = split(strSavedDaysStateCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanRange = ""
			else
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDaysStateCompare"" name=""cboDaysStateCompare"" onchange=""return cboDaysStateCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<input style=""width:55"" type=""text"" id=""txtDaysStateDays"" name=""txtDaysStateDays"" value=""" & FieldArray(1) & """>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDaysStateRange"">&nbsp;-&nbsp;<input style=""width:55"" type=""text"" id=""txtDaysStateRange"" name=""txtDaysStateRange"" maxlength=25 value=""" & FieldArray(2) & """></span>"
			response.write "<span style=""font-family:verdana;font-size:x-small"">&nbsp;Days</span></div>"
			response.write "</td>"
		case 36: 'Date Modified
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Date Modified:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """>"
			if trim(strSavedDateModifiedCompare) = "" then
				strSavedDateModifiedCompare = "|||"
			else
				strSavedDateModifiedCompare = strSavedDateModifiedCompare & "|"
				strSavedDateModifiedCompare = strSavedDateModifiedCompare & strSavedDateModifiedDays & "|"
				strSavedDateModifiedCompare = strSavedDateModifiedCompare & strSavedDateModifiedRange1 & "|"
				strSavedDateModifiedCompare = strSavedDateModifiedCompare & strSavedDateModifiedRange2
			end if
			FieldArray = split(strSavedDateModifiedCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanCount = "none"
				ShowSpanRange = ""
			else
				ShowSpanCount = ""
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDateModifiedCompare"" name=""cboDateModifiedCompare"" onchange=""return cboDateModifiedCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<span ID=""spnDateModifiedCount"" style=""white-space:nowrap;display:" & ShowSpanCount & ";font-family:verdana;font-size:x-small""><input style=""width:55"" type=""text"" id=""txtDateModifiedDays"" name=""txtDateModifiedDays"" value=""" & FieldArray(1) & """> Days Ago</span>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDateModifiedRange""><input style=""width:75"" type=""text"" id=""txtDateModifiedRange1"" name=""txtDateModifiedRange1"" maxlength=25 value=""" & FieldArray(2) & """>&nbsp;<a href=""javascript:PickDate(5);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a>&nbsp;-&nbsp;<input style=""width:75"" type=""text"" id=""txtDateModifiedRange2"" name=""txtDateModifiedRange2"" maxlength=25 value=""" & FieldArray(3) & """>&nbsp;<a href=""javascript:PickDate(6);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a></span>"
			response.write "</td>"
		case 35: 'Days Owner
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Days Owner:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """><div style=""white-space:nowrap"">"
			if trim(strSavedDaysOwnerCompare) = "" then
				strSavedDaysOwnerCompare = "||"
			else
				strSavedDaysOwnerCompare = strSavedDaysOwnerCompare & "|"
				strSavedDaysOwnerCompare = strSavedDaysOwnerCompare & strSavedDaysOwnerDays & "|"
				strSavedDaysOwnerCompare = strSavedDaysOwnerCompare & strSavedDaysOwnerRange
			end if

			FieldArray = split(strSavedDaysOwnerCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanRange = ""
			else
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboDaysOwnerCompare"" name=""cboDaysOwnerCompare"" onchange=""return cboDaysOwnerCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<input style=""width:55"" type=""text"" id=""txtDaysOwnerDays"" name=""txtDaysOwnerDays"" value=""" & FieldArray(1) & """>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnDaysOwnerRange"">&nbsp;-&nbsp;<input style=""width:55"" type=""text"" id=""txtDaysOwnerRange"" name=""txtDaysOwnerRange"" maxlength=25 value=""" & FieldArray(2) & """></span>"
			response.write "<span style=""font-family:verdana;font-size:x-small"">&nbsp;Days</span></div>"
			response.write "</td>"
		case 38: 'Target Date
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Target Date:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """>"
			if trim(strSavedTargetDateCompare) = "" then
				strSavedTargetDateCompare = "|||"
			else
				strSavedTargetDateCompare = strSavedTargetDateCompare & "|"
				strSavedTargetDateCompare = strSavedTargetDateCompare & strSavedTargetDateDays & "|"
				strSavedTargetDateCompare = strSavedTargetDateCompare & strSavedTargetDateRange1 & "|"
				strSavedTargetDateCompare = strSavedTargetDateCompare & strSavedTargetDateRange2
			end if
			FieldArray = split(strSavedTargetDateCompare,"|")
			ListArray = split("1|Less Than,2|Exactly,3|More Than,4|Between",",")
			if trim(FieldArray(0)) = "4" then
				ShowSpanCount = "none"
				ShowSpanRange = ""
			else
				ShowSpanCount = ""
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboTargetDateCompare"" name=""cboTargetDateCompare"" onchange=""return cboTargetDateCompare_onchange()"">"
			response.write "<option value="""" selected=""selected""/>"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(FieldArray(0)) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>&nbsp;"
			response.write "<span ID=""spnTargetDateCount"" style=""white-space:nowrap;display:" & ShowSpanCount & ";font-family:verdana;font-size:x-small""><input style=""width:55"" type=""text"" id=""txtTargetDateDays"" name=""txtTargetDateDays"" value=""" & FieldArray(1) & """> Days&nbsp;Away</span>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnTargetDateRange""><input style=""width:75"" type=""text"" id=""txtTargetDateRange1"" name=""txtTargetDateRange1"" maxlength=25 value=""" & FieldArray(2) & """>&nbsp;<a href=""javascript:PickDate(7);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a>&nbsp;-&nbsp;<input style=""width:75"" type=""text"" id=""txtTargetDateRange2"" name=""txtTargetDateRange2"" maxlength=25 value=""" & FieldArray(3) & """>&nbsp;<a href=""javascript:PickDate(8);""><img SRC=""../../MobileSE/Today/images/calendar.gif"" alt=""Choose"" border=""0"" align=""absmiddle"" WIDTH=""26"" HEIGHT=""21""></a></span>"
			response.write "</td>"
		case 41: 'Sorting
			response.write "<td nowrap=""nowrap"" style=""width:120px""><b>Sort&nbsp;Order:<b></td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """>"

			'------Begin Column 1---------------------------

			ListArray = split(strSortColumnNames,",")

			response.write "<SELECT id=cboSortColumn1 name=cboSortColumn1 style=""width:180px"" onchange=""javascript: SortColumn_onchange(1);"">"

			response.write "<OPTION selected=""selected""/>"

			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSortColumn1)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"

			response.write "<SELECT style=""Width:60px"" id=cboSort1Direction name=cboSort1Direction>"
			response.write "<OPTION selected=""selected""/>"
			ListArray = split("Asc,Desc",",")
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSort1Direction)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"

			'------End Column 1---------------------------
			response.write "<span style=""font-family:verdana;font-size=small;font-weight:bold"">&nbsp;&nbsp;,&nbsp;&nbsp;</span>"

			'------Begin Column 2---------------------------

			ListArray = split(strSortColumnNames,",")

			response.write "<SELECT id=cboSortColumn2 name=cboSortColumn2 style=""width:180px"" onchange=""javascript: SortColumn_onchange(2);"">"

			response.write "<OPTION selected=""selected""/>"

			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSortColumn2)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"

			response.write "<SELECT style=""Width:60px"" id=cboSort2Direction name=cboSort2Direction>"
			response.write "<OPTION selected=""selected""/>"
			ListArray = split("Asc,Desc",",")
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSort2Direction)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"
			'------End Column 2---------------------------
			response.write "<span style=""font-family:verdana;font-size=small;font-weight:bold"">&nbsp;&nbsp;,&nbsp;&nbsp;</span>"

			'------Begin Column 3---------------------------

			ListArray = split(strSortColumnNames,",")

			response.write "<SELECT id=cboSortColumn3 name=cboSortColumn3 style=""width:180px"" onchange=""javascript: SortColumn_onchange(3);"">"

			response.write "<OPTION selected=""selected""/>"

			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSortColumn3)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"

			response.write "<SELECT style=""Width:60px"" id=cboSort3Direction name=cboSort3Direction>"
			response.write "<OPTION selected=""selected""/>"
			ListArray = split("Asc,Desc",",")
			for each strValue in ListArray
				if trim(lcase(strValue)) = trim(lcase(strSavedSort3Direction)) then
					response.write "<Option selected=""selected"" value=""" & strValue & """>" & strValue & "</OPTION>"
				else
					response.write "<Option value=""" & strValue & """>" & strValue & "</OPTION>"
				end if
			next
			response.write "</SELECT>"
			'------End Column 3---------------------------
			response.write "</td>"
		case 28: 'Product Group
			SelectedArray = split(strSavedProductGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Product Group:</b><br/>"
			response.write "<SELECT style=""WIDTH: 100%; HEIGHT: " & FieldHeight & "px"" multiple  size=2 id=lstProductGroup name=lstProductGroup>"

			'mh if CurrentUserDivision = "1" then

				if strSavedProductGroup <> "" then
					response.write "<optgroup style="""" label=""-- Saved Groups -------------------------"">"
					for each strValue in SelectedArray
						ValuePair = split(strValue,":")
						select case trim(ValuePair(0))
							case "1"
								rs.open "spGetPartnerName " & clng(ValuePair(1)),cnExcalibur
								if rs.eof and rs.bof then
									strProductGroupName = "Unknown"
								else
									strProductGroupName = rs("name")
								end if
								rs.Close
							case "2"
								rs.open "spGetCycleName " & clng(ValuePair(1)),cnExcalibur
								if rs.eof and rs.bof then
									strProductGroupName = "Unknown"
								else
									strProductGroupName = rs("fullname") & ""
								end if
								rs.Close
							case "3"
								rs.open "spGetDevCenterName " & clng(ValuePair(1)),cnExcalibur
								if rs.eof and rs.bof then
									strProductGroupName = "Unknown"
								else
									strProductGroupName = rs("name")
								end if
								rs.Close
							case "4"
								ProductPhases = split("Unknown,Definition,Development,Production,Post-Production,Inactive",",")
								if clng(ValuePair(1)) >=0 and clng(ValuePair(1)) <=5 then
									strProductGroupName = ProductPhases(clng(ValuePair(1)))
								else
									strProductGroupName = "Unknown"
								end if
							case else
								strProductGroupName = "Unknown"
						end select
						response.write "<Option selected=""selected"" value=""" & server.HTMLEncode(strValue) & """>" & server.HTMLEncode(strProductGroupName) & "</OPTION>"
					next
					response.write "</optgroup>"
				end if

				if trim(currentuserpartner) = "1" then
					strSQL = "spListPartners 2"

					rs.Open strSQL,cnExcalibur,adOpenForwardOnly
					blnFound = false
					do while not rs.EOF
						if rs("ID") <> 1 then
							if not inlist(SelectedArray,"1:" & rs("ID")) then
								if not blnFound then
									response.write  "<optgroup style="""" label=""-- ODM -------------------------"">"
								end if
								blnFound = true
								response.write   "<Option value= ""1:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
							end if
						end if
						rs.MoveNext
					loop
					rs.Close
					if blnFound then
						response.write  "</optgroup>"
					end if
				end if

				strSQL = "spListPrograms"

				rs.Open strSQL,cnExcalibur,adOpenForwardOnly
				blnFound = false
				do while not rs.EOF
					if instr(lcase(rs("Name") & ""),"softpaq") = 0 then
						if not inlist(SelectedArray,"2:" & rs("ID")) then
							if not blnFound then
								response.write "<optgroup style="""" label=""-- Cycle -------------------------"">"
							end if
							blnFound = true
							response.write   "<Option value= ""2:" & rs("ID") & """>" & replace(rs("FullName"),"BNB Common Product ","") & "</OPTION>"
						end if
					end if
					rs.MoveNext
				loop
				rs.Close
				if blnFound then
					response.write  "</optgroup>"
				end if
				strSQL = "spListDevCenters"

				rs.Open strSQL,cnExcalibur,adOpenForwardOnly

				blnFound = false
				do while not rs.EOF
					if not inlist(SelectedArray,"3:" & rs("ID")) then
						if not blnFound then
							response.write "<optgroup style="""" label=""-- Dev. Center -------------------------"">"
						end if
						blnFound = true
						response.write   "<Option value= ""3:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
					end if
					rs.MoveNext
				loop
				rs.Close
				if blnFound then
					response.write "</optgroup>"
				end if

				strSQL = "spListProductStatuses"

				rs.Open strSQL,cnExcalibur,adOpenForwardOnly

				blnFound = false
				do while not rs.EOF
					if not inlist(SelectedArray,"4:" & rs("ID")) then
						if not blnFound then
							response.write "<optgroup style="""" label=""-- Product Phase -------------------------"">"
						end if
						blnFound = true
						response.write "<Option value= ""4:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
					end if
					rs.MoveNext
				loop
				rs.Close
				if blnFound then
					response.write "</optgroup>"
				end if

			'mh end if
			response.write "</SELECT></TD>"
		case 13: 'Tester Group
			SelectedArray = split(strSavedTesterGroup,",")

			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Tester&nbsp;Group:</b><br/>"
			response.write "<select id=lstTesterGroup name=lstTesterGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedTesterGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedTesterGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and tester=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 6: 'Developer Group
			SelectedArray = split(strSavedDeveloperGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Developer&nbsp;Group:</b><br/>"
			response.write "<select id=lstDeveloperGroup name=lstDeveloperGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"

			if strSavedDeveloperGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedDeveloperGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and Developer=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 45: 'Product PM Group
			SelectedArray = split(strSavedProductPMGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Product&nbsp;PM&nbsp;Group:</b><br/>"
			response.write "<select id=lstProductPMGroup name=lstProductPMGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"

			if strSavedProductPMGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedProductPMGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and ProductPM=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 46: 'Component Test Lead
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Comp.&nbsp;Test&nbsp;Lead:</b><br/>"
			response.write "<select id=lstComponentTestLead name=lstComponentTestLead multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedComponentTestLead,",")

			if strSavedComponentTestLead <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedComponentTestLead & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and componenttestlead=1 " & _
					 strDivisionFilter

			 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 47: 'Component Test Lead Group
			SelectedArray = split(strSavedComponentTestLeadGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Comp.&nbsp;Test&nbsp;Lead&nbsp;Group:</b><br/>"
			response.write "<select id=lstComponentTestLeadGroup name=lstComponentTestLeadGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedComponentTestLeadGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedComponentTestLeadGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and ComponenttestLead=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 48: 'Product Test Lead
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Prod.&nbsp;Test&nbsp;Lead:</b><br/>"
			response.write "<select id=lstProductTestLead name=lstProductTestLead multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedProductTestLead,",")

			if strSavedProductTestLead <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedProductTestLead & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if


			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and producttestlead=1 " & _
					 strDivisionFilter

			 strSQL = strSQL & " order by DisplayName; "
			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(selectedarray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 49: 'Product Test Lead Group
			SelectedArray = split(strSavedProductTestLeadGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Prod.&nbsp;Test&nbsp;Lead&nbsp;Group:</b><br/>"
			response.write "<select id=lstProductTestLeadGroup name=lstProductTestLeadGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedProductTestLeadGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedProductTestLeadGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and ProductTestLead=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
	   case 50: 'Approver
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Approver:</b><br/>"
			response.write "<select id=lstApprover name=lstApprover multiple style=""height:" & FieldHeight & "px;width:100%"">"

			SelectedArray = split(strSavedApprover,",")
			if strSavedApprover <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedApprover & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 "and approver=1 " & _
					 strDivisionFilter

			 strSQL = strSQL & " order by DisplayName; "

			 rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not InList(SelectedArray,rs("userid")) then
					response.write "<Option value=""" & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"
		case 51: 'Approver Group
			SelectedArray = split(strSavedApproverGroup,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Approver&nbsp;Group:</b><br/>"
			response.write "<select id=lstApproverGroup name=lstApproverGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedApproverGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedApproverGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and Approver=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 53: 'Graph Scale (NO  LONGER USED)
			response.write "<td nowrap=""nowrap"" style=""width:120px;font-family:verdana;font-size:x-small;font-weight:bold"">Graph Scale:</td>"
			response.write "<td nowrap=""nowrap"" style=""width:" & FieldWidth & """><div style=""white-space:nowrap"">"
			strSavedGraphScaleType = strSavedGraphScaleType
			ListArray = split("|Auto,1|Custom",",")
			if trim(strSavedGraphScaleType) = "1" then
				ShowSpanRange = ""
			else
				ShowSpanRange = "none"
			end if
			response.write "<select style=""WIDTH:90"" id=""cboGraphScaleType"" name=""cboGraphScaleType"" onchange=""return cboGraphScaleType_onchange()"">"
			for each strValue in ListArray
				ValuePair = split(strValue,"|")
				if trim(ValuePair(0)) = trim(strSavedGraphScaleType) then
					response.write "<option selected=""selected"" value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				else
					response.write "<option value=""" & ValuePair(0) & """>" & ValuePair(1) & "</option>"
				end if
			next
			response.write "</select>"
			response.write "<span style=""white-space:nowrap;display:" & ShowSpanRange & ";font-family:verdana;font-size:x-small"" ID=""spnGraphScale"">&nbsp;<input style=""width:55"" type=""text"" id=""txtGraphScale"" name=""txtGraphScale"" maxlength=25 value=""" & server.htmlencode(strSavedGraphScale) & """>&nbsp;Observations per Gridline</span>"
			response.write "</div>"
			response.write "</td>"
		case 52: 'Component PM Group
			SelectedArray = split(strSavedComponentPMGroup,",")

			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Component&nbsp;PM&nbsp;Group:</b><br/>"
			response.write "<select id=lstComponentPMGroup name=lstComponentPMGroup multiple style=""height:" & FieldHeight & "px;width:100%"">"
			if strSavedComponentPMGroup <> "" then
				rs.open "SELECT XLS_Org_ID as ID,XLS_Org_Name as name FROM dbo.vWorkgroupPrimary with (NOLOCK) where XLS_Org_ID in (" & strSavedComponentPMGroup & ") order by XLS_Org_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if

			strSQL = "Select GroupID as ID, name " & _
					 "from dbo.list_group with (NOLOCK) " & _
					 "where active=1 " & _
					 "and ComponentPM=1 " & _
					 strDivisionFilter

			strSQL = strSQl & " order by name;"

			rs.Open strSQL,cnSIO,adOpenForwardOnly

			do while not rs.EOF
				if not inlist(SelectedArray,rs("ID")) then
					response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close

			response.write "</select>"
			response.write "</td>"
		case 14: 'Owner
			SelectedArray = split(strSavedOwner,",")
			response.write "<td style=""width:" & FieldWidth & """>"
			response.write "<b>Owner:</b><br/>"
			response.write "<select id=lstOwner name=lstOwner multiple style=""height:" & FieldHeight & "px;width:100%"">"

			if strSavedOwner <> "" then
				rs.open "Select u.user_id, u.User_Name from dbo.Users u with (NOLOCK) where user_id in (" & strSavedOwner & ") order by u.User_Name",cnSIO
				do while not rs.EOF
					response.write "<Option selected=""selected"" value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
					rs.MoveNext
				loop
				rs.Close
			end if
			 strSQL = "Select UserID, DisplayName " & _
					 "FROM dbo.list_actor with (NOLOCK) " & _
					 "where active=1 " & _
					 strDivisionFilter

			 strSQL = strSQL & " order by DisplayName; "

			 rs.Open strSQL,cnSIO,adOpenForwardOnly
			do while not rs.EOF
				if not inlist(SelectedArray,rs("userid")) then
					response.write "<Option value= """ & rs("userid") & """>" & rs("DisplayName") & "</OPTION>"
				end if
				rs.MoveNext
			loop
			rs.Close
			response.write "</select></td>"

		case else 'Write a blank cell
			response.write "<td>&nbsp;</td>"
		end select
		if RemainingColumnCount = 0 then
'			response.write "<td style=""background-color:green;width=100%"">&nbsp;</td>"
		else
			response.write "<td style=""width:10px"">&nbsp;</td>"
		end if

	end sub

	%>
	<input style="display: none" id="txtProfileUpdateType" name="txtProfileUpdateType" type="text" value="" />
	<input style="display: none" id="txtProfileType" name="txtProfileType" type="text" value="" />
	<input style="display: none" id="txtReportSections" name="txtReportSections" type="text" value="" />
	<input style="display: none" id="txtReportSectionParameters" name="txtReportSectionParameters" type="text" value="" />
	<input style="display: none" id="txtProfileUpdateID" name="txtProfileUpdateID" type="text" value="" />
	<input style="display: none" id="txtNewProfileName" name="txtNewProfileName" type="text" value="" />
	<input style="display: none" id="txtNewTodayLink" name="txtNewTodayLink" type="text" value="" />
	<input style="display: none" id="txtNewReportFormat" name="txtNewReportFormat" type="text" value="" />
	<input style="display: none" id="txtUserID" name="txtUserID" type="text" value="<%=CurrentUserID%>" />
	<textarea style="display: none" id="txtPageLayout" name="txtPageLayout" cols="120" rows="3"><%=strProfilePageLayout%></textarea>
	</form>
	<input style="display: none" id="txtReturnValue" name="txtReturnValue" type="text" value="" />
	<input style="display: none" id="txtReturnValue2" name="txtReturnValue2" type="text" value="" />
	<input style="display: none" id="txtReturnValue3" name="txtReturnValue3" type="text" value="" />
	<%
	set rs = nothing
	cnExcalibur.Close
	set cnExcalibur = nothing
	cnSIO.Close
	set cnSIO = nothing

	function InList(MyArray, strFind)
		dim strItem
		dim strFind2
		dim blnFound
		blnFound = false
		strFind2 = lcase(trim(strFind))

		for each strItem in MyArray
			if lcase(trim(strItem)) = strFind2 then
				blnFound = true
				exit for
			end if
		next
		InList = blnFound
	end function

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function

	function GetValue(strFieldname)
		dim strTemp
		if ProfileData <> "" then
			if instr("&" & profileData,"&" & strFieldname & "=") > 0 then
				strTemp = mid(ProfileData,instr("&" & profileData,"&" & strFieldname & "=")+len(strFieldname) + 1)
				strTemp = left(strTemp & "&",instr(strTemp & "&","&")-1)
				getValue = urldecode(strTemp)
			else
				GetValue=""
			end if
		elseif trim(request(strFieldname)) <> "" then
			GetValue = request(strFieldname)
		else
			GetValue= ""
		end if
	end function

	function URLDecode(byVal encodedstring)
		Dim strIn, strOut, intPos, strLeft,strRight, intLoop
		strIn  = encodedstring : strOut = "" : intPos = Instr(strIn, "+")
		Do While intPos
			strLeft = "" : strRight = ""
			If intPos > 1 then strLeft = Left(strIn, intPos - 1)
			If intPos < len(strIn) then strRight = Mid(strIn, intPos + 1)
			strIn = strLeft & " " & strRight
			intPos = InStr(strIn, "+")
			intLoop = intLoop + 1
		Loop
		intPos = InStr(strIn, "%")
		Do while intPos
			If intPos > 1 then strOut = strOut & Left(strIn, intPos - 1)
			strOut = strOut & Chr(CInt("&H" & mid(strIn, intPos + 1, 2)))
			If intPos > (len(strIn) - 3) then
				strIn = ""
			Else
				strIn = Mid(strIn, intPos + 3)
			End If
			intPos = InStr(strIn, "%")
		Loop
		URLDecode = strOut & strIn
	end function

	function ScrubSQL(strWords)

		dim badChars
		dim newChars
		dim i

'		strWords=replace(strWords,"'","''")

		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update")
		newChars = strWords

		for i = 0 to uBound(badChars)
			newChars = replace(newChars, badChars(i), "")
		next

		ScrubSQL = newChars

	end function
	%>
	<div id="mnuPopup" style="display: none; position: absolute; padding: 0px; width: 140px; background: white; border: 1px solid gainsboro; z-index: 100">
		<div style="border-right: black 1px solid; border-top: black 1px solid; left: 0px; border-left: black 1px solid; border-bottom: black 1px solid; position: relative; top: 0px">
			<div onmouseover="this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
				<span onclick="javascript:StatusReport(1);" style="font-family: arial; font-size: x-small">&nbsp;&nbsp;&nbsp;Product&nbsp;Status</span></div>
			<div onmouseover="this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
				<span onclick="javascript:StatusReport(0);" style="font-family: arial; font-size: x-small">&nbsp;&nbsp;&nbsp;Deliverable&nbsp;Status&nbsp;&nbsp;&nbsp;&nbsp;</span></div>
			<div id="CustomMenuOptions">
				<%=CustomStatusReports%></div>
			<div onmouseover="this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'" onmouseout="this.style.background='white';this.style.color='black'">
				<span onclick="javascript:CustomStatusReport();" style="font-family: arial; font-size: x-small">&nbsp;&nbsp;&nbsp;Custom&nbsp;Reports&nbsp;&nbsp;&nbsp;</span></div>
		</div>
	</div>
	<iframe style="display: none; width: 100%; height: 300px" id="ProfileFrame" name="ProfileFrame"></iframe>
	<iframe style="display: none; width: 100%; height: 300px" id="ReportMenuFrame" name="ReportMenuFrame"></iframe>
	<div id="FilterLoadingMessage" style="display: none; position: absolute; background: #FFFFCC; width: 2px; height: 2px; left: 0px; top: 0px; padding: 10px; background: cornsilk; border: 2px ridge gainsboro; z-index: 100; font-family: verdana; font-size: x-small; font-weight: bold; color: #000080">
		Loading&nbsp;Profile.&nbsp;&nbsp;Please&nbsp;Wait...
	</div>
   <div style="display: none;">
    <div id="iframeDialog" title="ExtendTables">
        <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
    </div>
</div>
</body>
</html>
