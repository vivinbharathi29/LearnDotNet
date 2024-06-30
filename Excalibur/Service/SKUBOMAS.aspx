<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_SKUBOMAS" EnableEventValidation="False" Codebehind="SKUBOMAS.aspx.vb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

    <title>SKU BOM - Advanced Search</title>
    <link href="/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="/service/style.css" />
    <link rel="stylesheet" type="text/css" href="/service/sample.css" />

<script type="text/javascript" language="javascript">


    function HeaderMouseOver() {
        window.event.srcElement.style.cursor = "hand";
        window.event.srcElement.style.color = "red";
    }

    function HeaderMouseOut() {
        window.event.srcElement.style.color = "black";
    }


    function selectDate(sElementID) {

        var sDate;
        var oDateElement = document.getElementById(sElementID);

        sDate = window.showModalDialog("../mobilese/today/caldraw1.asp", sElementID, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");

        if ((sDate != null) && (sDate != undefined)) {
            oDateElement.value = sDate;
        }

    }

    function updateStatus(oStatusElement, sMsg, sColor) {
        oStatusElement.style.color = sColor;
        oStatusElement.innerHTML = sMsg;
    }

    function updateListSelectionCount(oList, oDiv) {
        var i = 0;
        var c = 0;

        if (oList.selectedIndex != -1) {
            for (i = 0; i < oList.options.length; i++) {
                if (oList.options[i].selected) {
                    c++;
                }
            }
        }

        oDiv.innerHTML = "(" + c.toString() + " selected)";
    }


    function resetForm() {
        //document.filterForm.reset();

        try {

            document.filterForm.ddlRptProfile.selectedIndex = 0;

            document.filterForm.chkbxGeoNA.checked = false;
            document.filterForm.chkbxGeoLA.checked = false;
            document.filterForm.chkbxGeoAPJ.checked = false;
            document.filterForm.chkbxGeoEMEA.checked = false;

            document.filterForm.chkbxSKGeoNA.checked = false; 
            document.filterForm.chkbxSKGeoLA.checked = false; 
            document.filterForm.chkbxSKGeoAPJ.checked = false; 
            document.filterForm.chkbxSKGeoEMEA.checked = false; 

            document.filterForm.lstbxBrands.selectedIndex = -1; 
            document.filterForm.lstbxCats.selectedIndex = -1; 
            document.filterForm.lstbxOSSPs.selectedIndex = -1;

            document.filterForm.txtbxKMATS.value = ""; 
            document.filterForm.txtbxSKUS.value = ""; 
            document.filterForm.txtbxAVS.value = ""; 
            document.filterForm.txtbxSKs.value = ""; 
            document.filterForm.txtbxSFPNS.value = ""; 
            document.filterForm.txtbxSKUAVS.value = ""; 
            document.filterForm.txtbxSAs.value = ""; 
            document.filterForm.txtbxComps.value = ""; 

            document.filterForm.chkbxAllSKUAVs.checked = false;

            document.filterForm.chkbxAdded.checked = false; 
            document.filterForm.chkbxUpdated.checked = false; 

            document.filterForm.txtbxFromDate.value = ""; 
            document.filterForm.txtbxToDate.value = ""; 

            document.filterForm.chkbxAllAVs.checked = false; 

            document.filterForm.rdoProdDiv_0.checked = true; 
            document.filterForm.rdoProdDiv_0.click();
            
            updateListSelectionCount(document.getElementById("lstbxBrands"), document.getElementById("PBCount"));
            updateListSelectionCount(document.getElementById("lstbxCats"), document.getElementById("SCCount"));
            updateListSelectionCount(document.getElementById("lstbxOSSPs"), document.getElementById("OSSPCount"));

            updateStatus(document.getElementById("lblProfileStatus"), "<b>Select a Profile to load.</b>", "black");
            updateStatus(document.getElementById("lblStatus"), "Set the desired filter options and click the <b>Submit</b> button to retrieve applicable data.", "black");

        } catch (e) {
            alert(e.Description);
        }
    }


    function processRequest(oStatusElement, sMsg) {

        document.getElementById("lblStatus").innerHTML = "";
        document.getElementById("lblProfileStatus").innerHTML = "";

        updateStatus(oStatusElement, sMsg, "black");

    }



    function inList(oList, sValue) {
        var i;
        var bResult = false;

        try {
            for (i = 0; i < oList.options.length; i++) {
                if (oList.options[i].value == sValue) {
                    bResult = true;
                    break;
                }
            }
        }
        catch (e) {
            alert(e.Description);
        }

        return bResult;
    }


    function addSelections(oSourceList, oDestList, sOptionalDestValueSuffix, sOptionalProhibitDestValueSuffix) {
        var i;
        var iSourceSelIdx = getFirstSelection(oSourceList);
        var sValueSuffix = "";
        var sProhibitValueSuffix = "";

        if (sOptionalDestValueSuffix != null) {
            if (sOptionalDestValueSuffix.length > 0) sValueSuffix = sOptionalDestValueSuffix;
        }

        if (sOptionalProhibitDestValueSuffix != null) {
            if (sOptionalProhibitDestValueSuffix.length > 0) sProhibitValueSuffix = sOptionalProhibitDestValueSuffix;
        }

        try {
            if (iSourceSelIdx > -1) {
                for (i = 0; i < oSourceList.options.length; i++) {
                    if (oSourceList.options[i].selected) {
                        if ((!inList(oDestList, oSourceList.options[i].value.toString() + sValueSuffix)) && (!inList(oDestList, oSourceList.options[i].value.toString() + sProhibitValueSuffix))) {
                            oDestList.options[oDestList.options.length] = new Option(oSourceList.options[i].text, oSourceList.options[i].value + sValueSuffix);
                        }
                    }
                }

                updateSelectedColumns(oDestList);

            } else {
                alert("Select one or more items in the list before clicking 'Add'!");
            }
        } catch (e) {
            alert(e.Description);
        }
    }

    function removeColumnSelections() {
        var oSelList = document.getElementById("selectedColumns");
        var iSelIdx = getFirstSelection(oSelList);
        var oOrderList = document.getElementById("orderByCols");
        var iListLen = oSelList.options.length;
        var i;
        var j;
        var sColValue = "";

        var sOrdValue;
        var aOrdColIDDir;


        try {
            if (iListLen > 0) {
                if (iSelIdx > -1) {

                    for (i = oSelList.options.length - 1; i >= 0; i--) {
                        if (oSelList.options[i].selected) {
                            sColValue = oSelList.options[i].value.toString();
                            oSelList.options[i] = null;

                            // Remove the appropriate item from the order by list

                            for (j = oOrderList.options.length - 1; j >= 0; j--) {

                                sOrdValue = oOrderList.options[j].value.toString();
                                aOrdColIDDir = sOrdValue.split("-");

                                if (aOrdColIDDir[0] == sColValue) {
                                    oOrderList.options[j] = null;
                                }
                            }
                        }
                    }

                    updateSelectedColumns(oSelList);
                    updateSelectedColumns(oOrderList);

                } else {
                    alert("Select one or more items in the list before clicking 'Remove'!");
                }
            } else {
                alert("List is empty!");
            }
        } catch (e) {
            alert(e.Description);
        }
    }

    function removeOrderColumnSelections() {
        var oSelList = document.getElementById("orderByCols");
        var iSelIdx = getFirstSelection(oSelList);
        var iListLen = oSelList.options.length;
        var i;
        var j;
        var sColValue = "";



        try {
            if (iListLen > 0) {
                if (iSelIdx > -1) {

                    for (i = oSelList.options.length - 1; i >= 0; i--) {
                        if (oSelList.options[i].selected) {
                            sColValue = oSelList.options[i].value.toString();
                            oSelList.options[i] = null;
                        }
                    }

                    updateSelectedColumns(oSelList);

                } else {
                    alert("Select one or more items in the list before clicking 'Remove'!");
                }
            } else {
                alert("List is empty!");
            }
        } catch (e) {
            alert(e.Description);
        }
    }

    function getFirstSelection(oList) {
        var i;

        try {
            if (oList.selectedIndex > -1) {
                for (i = 0; i < oList.options.length; i++) {
                    if (oList.options[i].selected) {
                        break;
                    }
                }
            } else {
                i = -1;
            }
        } catch (e) {
            alert(e.Message);
        }

        return i;
    }


    function getTotalSelections(oList) {
        var i;
        var t = 0;

        try {
            if (oList.selectedIndex > -1) {
                for (i = 0; i < oList.options.length; i++) {
                    if (oList.options[i].selected) {
                        t = t + 1;
                    }
                }
            }
        } catch (e) {
            alert(e.Message);
        }

        return t;
    }


    function getListSelections(oList) {
        var sSelections = "";
        var aSelections = null;
        var i = 0;

        try {
            if (oList.selectedIndex > -1) {
                for (i = 0; i < oList.options.length; i++) {
                    if (oList.options[i].selected) {

                        if (sSelections.length == 0) {
                            sSelections = i.toString();
                        } else {
                            sSelections += "," + i.toString();
                        }
                    }
                }
            }
        } catch (e) {

        }

        aSelections = sSelections.split(",");

        return aSelections;
    }

    function getListIndexByValue(oList, sValue) {
        var i;
        var iIdx = -1;

        for (i = 0; i < oList.options.length; i++) {
            if (oList.options[i].value == sValue) {
                iIdx = i;
                break;
            }
        }

        return iIdx;
    }


    function moveSelection(oSelList, iDir) {
        // Consider reworking this routine
        //var oSelList = document.getElementById("selectedColumns");
        var iSelIdx = getFirstSelection(oSelList);
        var iSwapIdx = -1;
        var iListLen = oSelList.options.length;
        var sOptAText;
        var sOptAValue;
        var sOptBText;
        var sOptBValue;
        var iTotalSels = 0;
        var aSelValues = getListValueSelections(oSelList).split(",");
        var i;

        try {

            if (iListLen > 0) {

                if (aSelValues.length > 0) {

                    for (i = 0; i < aSelValues.length; i++) {

                        iSelIdx = getListIndexByValue(oSelList, aSelValues[i]);

                        if (iSelIdx > -1) {

                            switch (iDir) {
                                case 0: // Move Down 
                                    if (iSelIdx <= oSelList.options.length - 2) {

                                        iSwapIdx = iSelIdx + 1;

                                        sOptAValue = oSelList.options[iSelIdx].value;
                                        sOptAText = oSelList.options[iSelIdx].text;

                                        sOptBValue = oSelList.options[iSwapIdx].value;
                                        sOptBText = oSelList.options[iSwapIdx].text;

                                        oSelList.options[iSelIdx].text = sOptBText;
                                        oSelList.options[iSelIdx].value = sOptBValue;

                                        oSelList.options[iSwapIdx].text = sOptAText;
                                        oSelList.options[iSwapIdx].value = sOptAValue;

                                        oSelList.options[iSelIdx].selected = false;
                                        oSelList.options[iSwapIdx].selected = true;

                                    }
                                    break;
                                case 1: // Move Up
                                    if (iSelIdx > 0) {

                                        iSwapIdx = iSelIdx - 1;

                                        sOptAValue = oSelList.options[iSelIdx].value;
                                        sOptAText = oSelList.options[iSelIdx].text;

                                        sOptBValue = oSelList.options[iSwapIdx].value;
                                        sOptBText = oSelList.options[iSwapIdx].text;

                                        oSelList.options[iSelIdx].text = sOptBText;
                                        oSelList.options[iSelIdx].value = sOptBValue;

                                        oSelList.options[iSwapIdx].text = sOptAText;
                                        oSelList.options[iSwapIdx].value = sOptAValue;

                                        oSelList.options[iSelIdx].selected = false;
                                        oSelList.options[iSwapIdx].selected = true;

                                    }
                                    break;
                            }
                        }
                    }

                    updateSelectedColumns(oSelList);

                } else {
                    alert("Select ONE item in the list in order to move it UP or DOWN!");
                }
            } else {
                alert("List is empty!");
            }
        }
        catch (e) {
            alert(e.Message)
        }
    }


    function updateSelectedColumns(oSelList) {
        //var oSelList = document.getElementById("selectedColumns");
        var shiddenInputID = oSelList.getAttribute("hiddenInputID");
        var oSelCols = document.getElementById(shiddenInputID);
        var sSels = "";
        var i;

        try {
            for (i = 0; i < oSelList.length; i++) {

                if (sSels.length == 0) {
                    sSels = oSelList.options[i].value.toString();
                } else {
                    sSels = sSels + "," + oSelList.options[i].value.toString();
                }
            }

            oSelCols.value = sSels;

            // MAY NEED TO REINSTATE OF ACCOMMODATE DEPENDENCIES HERE updateOrderByColumn(document.getElementById("orderByCol"));

        } catch (e) {
            alert(e.Message);
        }
    }


    function updateOrderByColumn(oList) {

        try {
            var oOrderCol = document.getElementById("orderCols");

            //oOrderCol.value=oList.value.toString(); Accommodate multiple Columns 8/26/2011

            // Update the Direction Radio Buttons


            var sCurrValue = oList.value.toString();
            var aOrdColIDDir = sCurrValue.split("-");


            document.filterForm.OrderByDir.value = aOrdColIDDir[1];

            if (aOrdColIDDir[1] == "0") {
                // Ascending
                document.getElementById("orderByAsc").checked = true;
                document.getElementById("orderByDesc").checked = false;

            }
            else {
                // Descending
                document.getElementById("orderByAsc").checked = false;
                document.getElementById("orderByDesc").checked = true;
            }


        } catch (e) {
            alert(e.Message);
        }

    }


    function updateCurrentOrderColumn(sDirValue) {

        var oOrdList = document.getElementById("orderByCols");

        if (oOrdList.selectedIndex != -1) {
            var sCurrValue = oOrdList.value.toString()
            var aOrdColIDDir = sCurrValue.split("-");

            oOrdList.options[oOrdList.selectedIndex].value = aOrdColIDDir[0] + "-" + sDirValue;
            updateSelectedColumns(oOrdList);
        }
    }


    function populateResultParameters() {
        var oAllList = document.getElementById("allColumns");

        var oSelList = document.getElementById("selectedColumns");
        var oSelCols = document.getElementById("selectedCols");

        var oOrderList = document.getElementById("orderByCols");
        var oOrderCols = document.getElementById("orderCols");

        var sSels = oSelCols.value;

        var sOrdSels = oOrderCols.value;
        var aOrdColVals = null;
        var aOrdColIDDir = null;

        var i;
        var j;

        try {

            if (sSels.length == 0) {

                sSels = "4,3,2";
                oSelCols.value = sSels

                sOrdSels = "4-0";
                oOrderCols.value = sOrdSels;
            }

            var aSelCols = sSels.split(",");

            oSelList.options.length = 0;
            oOrderList.options.length = 0;

            // Populate Selected Result Columns List
            for (i = 0; i < aSelCols.length; i++) {
                for (j = 0; j < oAllList.options.length; j++) {
                    if (aSelCols[i].toString() == oAllList.options[j].value.toString()) {

                        if (!inList(oSelList, oAllList.options[j].value.toString())) {
                            oSelList.options[oSelList.options.length] = new Option(oAllList.options[j].text, oAllList.options[j].value);

                            // Only Add what has been selected for ordering 8/26/2011<--- Accommodate this outside of these nested iterations
                            //oOrderList.options[oOrderList.options.length] = new Option(oAllList.options[j].text, oAllList.options[j].value);
                        }
                    }
                }
            }


            // Populate Selected Result Columns ORDER BY Column List (ColumnID-Direction) appearance in list is implied/denoted by position in delimited string
            if (sOrdSels.length == 0) {
                sOrdSels = aSelCols[0] + "-0";
                oOrderCols.value = sOrdSels;
            }

            if (sOrdSels.length > 0) {
                //oOrderList.value = oOrderCol.value;

                if (sOrdSels.indexOf("-") != -1) {
                    aOrdColVals = sOrdSels.split(",");

                    for (i = 0; i < aOrdColVals.length; i++) {
                        aOrdColIDDir = aOrdColVals[i].split("-");

                        j = getListIndexByValue(oSelList, aOrdColIDDir[0]);

                        oOrderList.options[oOrderList.options.length] = new Option(oSelList.options[j].text, aOrdColVals[i]);
                    }
                } else {
                    // Accommodate previous versions <--- PERFORM GLOBAL UPDATE IN DATABASE PRIOR TO DEPLOYMENT OF THIS VERSION
                    // Add single Column to the List and Set to the appropriate Direction
                }
            }

        } catch (e) {
            alert(e.Message);
        }
    }


    function processProfileAction(sActionType) {
        var oPN = document.getElementById("profileName");
        var oCA = document.getElementById("continueAction");
        oCA.value = "FALSE";
        var sPName = "";
        var sNewPName = "";
        var oPList = document.getElementById("ddlRptProfile");
        var iSelID = 0;

        switch (sActionType) {
            case "A": // Apply Profile

                iSelID = oPList.selectedIndex

                if (iSelID == 0) {
                    processRequest(document.getElementById('lblProfileStatus'), "Select a Profile to load.");
                } else {
                    processRequest(document.getElementById('lblProfileStatus'), '<b>Loading profile, please wait...</b>');
                }
                break;

            case "I": // Add Profile
                sPName = window.prompt("Enter the name of the Profile to add:", sPName);

                if ((sPName != undefined) && (sPName != null)) {
                    oPN.value = sPName;
                    oCA.value = "TRUE";

                    processRequest(document.getElementById("lblProfileStatus"), '<b>Processing request, please wait...</b>');
                }

                break;

            case "U": // Update Profile
                break;

            case "D": // Delete Profile

                iSelID = oPList.selectedIndex;

                if (iSelID > 0) {
                    sPName = oPList.options[iSelID].text;

                    if (window.confirm("Delete Profile, '" + sPName + "'?")) {
                        oCA.value = "TRUE";
                        processRequest(document.getElementById("lblProfileStatus"), '<b>Processing request, please wait...</b>');
                    } else {
                        oCA.value = "FALSE";
                    }
                }

                break;

            case "R": // Rename Profile

                iSelID = oPList.selectedIndex;


                if (iSelID > 0) {
                    sPName = oPList.options[iSelID].text;
                    sNewPName = window.prompt("Enter the new name of the Profile, '" + sPName + "':", sPName);

                    if ((sNewPName != undefined) && (sNewPName != null)) {
                        oPN.value = sNewPName;
                        oCA.value = "TRUE";
                        processRequest(document.getElementById("lblProfileStatus"), '<b>Processing request, please wait...</b>');
                    }
                }

                break;

            case "RSP": // Remove Shared Profile

                iSelID = oPList.selectedIndex;

                if (iSelID > 0) {
                    sPName = oPList.options[iSelID].text;
                    var sTitle = document.getElementById("lnkBtnRemoveSharedProfile").getAttribute("title").toString();

                    if (window.confirm(sTitle + ", '" + sPName + "'?")) {
                        oCA.value = "TRUE";
                        processRequest(document.getElementById("lblProfileStatus"), '<b>Processing request, please wait...</b>');
                    } else {
                        oCA.value = "FALSE";
                    }
                }

                break;

        }
    }


    function checkReset() {
        var oReset = document.getElementById("resetFlag");

        if (oReset.value == "TRUE") {
            resetForm();
        }
    }


    //************************************************************************************
    // AJAX --- ACCOMMODATE MULTIPLE REQUESTS OR FLAG AS IN PROCESSS
    //************************************************************************************
    function getListValueSelections(oList) {
        var i;
        var sSelValues = "";

        try {
            if (oList.selectedIndex > -1) {
                for (i = 0; i < oList.options.length; i++) {
                    if (oList.options[i].selected) {
                        if (sSelValues.length == 0) {
                            sSelValues = oList.options[i].value.toString();
                        } else {
                            sSelValues += "," + oList.options[i].value.toString();
                        }
                    }
                }
            }
        } catch (e) {
            alert(e.Message);
        }

        return sSelValues;
    }


    function validateWildCards(oElement) {
        var sList = oElement.value;
        var bResult = true;

        if ((sList.indexOf(",") >= 0) && (sList.indexOf("*") >= 0)) {
            var iNumWCs = 0;
            var i;
            var aList = sList.split(",");

            for (i = 0; i < aList.length; i++) {
                if (aList[i].indexOf("*") >= 0) iNumWCs++;
            }

            if (iNumWCs != aList.length) bResult = false;

        }

        return bResult;
    }


    function validateTextParameters() {
        var bResult = true;

        // Validate all inputs that can contain wild card values (currently "mixed mode" is NOT SUPPORTED) --- i.e. - ALL values must contain wild card characters OR ALL values must be WHOLE strings without wild characters

        //SFPNS
        if (!validateWildCards(document.getElementById("txtbxSFPNS"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [Service Family P/Ns] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxSFPNS").focus();
            return false;
        }

        //KMATs
        if (!validateWildCards(document.getElementById("txtbxKMATS"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [KMAT P/Ns] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxKMATS").focus();
            return false;
        }


        //SKUs
        if (!validateWildCards(document.getElementById("txtbxSKUS"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [SKU P/Ns] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxSKUS").focus();
            return false;
        }

        //SKUAVs
        if (!validateWildCards(document.getElementById("txtbxSKUAVS"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [SKU AVs] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxSKUAVS").focus();
            return false;
        }

        //SKs
        if (!validateWildCards(document.getElementById("txtbxSKs"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [Spare Kits] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxSKs").focus();
            return false;
        }

        //AVs
        if (!validateWildCards(document.getElementById("txtbxAVS"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [AV Qualifiers] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxAVS").focus();
            return false;
        }

        //SAs
        if (!validateWildCards(document.getElementById("txtbxSAs"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [SubAssemblies] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxSAs").focus();
            return false;
        }

        //COMPs
        if (!validateWildCards(document.getElementById("txtbxComps"))) {
            updateStatus(document.getElementById("lblStatus"), "<b>All delimited values in [Components] must either contain wild cards or none at all.  Mixed values are not supported.</b>", "red");
            document.getElementById("txtbxComps").focus();
            return false;
        }

        //ACTION DATES
        var oFromDate=document.getElementById("txtbxFromDate");
        var oToDate = document.getElementById("txtbxToDate");
        var oDateColumn = document.getElementById("ddlActionDateColumn");
        var sFromDate=oFromDate.value;
        var sToDate=oToDate.value;

        if ((sFromDate.length > 0) && (sToDate.length > 0) && getListValueSelections(oDateColumn) == "0") {
            updateStatus(document.getElementById("lblStatus"), "<b>You must specify an Action Date Column with which to filter.</b>", "red");
            oDateColumn.focus();
            return false;
        } else if (((sFromDate.length > 0) && (sToDate.length == 0)) || ((sFromDate.length == 0) && (sToDate.length > 0))) {

            updateStatus(document.getElementById("lblStatus"), "<b>You must specify BOTH FROM and TO dates.</b>", "red");

            if (sFromDate.length == 0) {
                oFromDate.focus();
            } else {
                oToDate.focus();
            }
            return false;
        } else if ((sFromDate.length == 0) && (sToDate.length == 0) && getListValueSelections(oDateColumn) != "0") {
            updateStatus(document.getElementById("lblStatus"), "<b>You must specify BOTH FROM and TO dates when an Action Date Column is selected.</b>", "red");
            oDateColumn.focus();
            return false;
        }


        // SPB Change Log Filter
        var oSA = document.getElementById("txtbxSAs");
        var oCMP = document.getElementById("txtbxComps");
        var oSPBF = document.getElementById("chkbxSPBLogFilter");
        var sSA = oSA.value;
        var sCMP = oCMP.value;

        if ((oSPBF.checked) && (sSA.length == 0) && (sCMP.length == 0)) {
            updateStatus(document.getElementById("lblStatus"), "<b>You must specify at least ONE Sub-Assembly OR Component when 'Apply to SPB Change Log' is checked.</b>", "red");
            oSA.focus();
            return false;
        }

        // SKU Change Log Filter
        var oSKUAVs = document.getElementById("txtbxSKUAVS");
        var sSKUAVs=oSKUAVs.value;
        var oSKUF = document.getElementById("chkbxSKUAVLogFilter");

        if ((oSKUF.checked) && (sSKUAVs.length == 0)) {
            updateStatus(document.getElementById("lblStatus"), "<b>You must specify at least ONE SKU AV when 'Apply to SKU Change Log' is checked.</b>", "red");
            oSKUAVs.focus();
            return false;
        }

	// Added to validate 'Require All SKU AVs'
	var oReqAllSKUAVs=document.getElementById("chkbxAllSKUAVs");

	if((oReqAllSKUAVs.checked) && (sSKUAVs.length == 0))
	{
            updateStatus(document.getElementById("lblStatus"), "<b>You must specify at least ONE SKU AV when 'Require All SKU AVs' is checked.</b>", "red");
            oSKUAVs.focus();
            return false;
	}

        return bResult;
    }


    function getOrderByColumnData(sData, iMode) 
    {
	// NEED TO ADD ERROR TRAPPING...
        var sRetValue = "";
        var aData = sData.split(",");
        var aSubData = null;
        var i;

        for (i = 0; i < aData.length; i++) {

            aSubData = null;
            aSubData = aData[i].split("-");

	        if((aSubData[iMode]!=undefined)&&(aSubData[iMode]!=null))
	        {
	            if (sRetValue.length == 0) sRetValue = aSubData[iMode];
	            else sRetValue += "," + aSubData[iMode];
	        }
        }

        return sRetValue;
    }


    function getParameters() {
        var sValues = "";

        // NEED TO ADD LOGIC TO REPLACE PIPE DELIMITER ON OPEN INPUTS
        try {
            // ColumnIDs
            sValues = document.getElementById("selectedCols").value;

            // ColumnOrderByIDs
            sValues += "|" + getOrderByColumnData(document.getElementById("orderCols").value,0);

            // ColumnOrderAscDesc
            sValues += "|" + getOrderByColumnData(document.getElementById("orderCols").value,1);

            //ServiceGeoNA
            sValues += "|" + document.getElementById("chkbxGeoNA").checked.toString();

            //ServiceGeoLA
            sValues += "|" + document.getElementById("chkbxGeoLA").checked.toString();

            //ServiceGeoAPJ
            sValues += "|" + document.getElementById("chkbxGeoAPJ").checked.toString();

            //ServiceGeoEMEA
            sValues += "|" + document.getElementById("chkbxGeoEMEA").checked.toString();

            //SKGeoNA
            sValues += "|" + document.getElementById("chkbxSKGeoNA").checked.toString();

            //SKGeoLA
            sValues += "|" + document.getElementById("chkbxSKGeoLA").checked.toString();

            //SKGeoAPJ
            sValues += "|" + document.getElementById("chkbxSKGeoAPJ").checked.toString();

            //SKGeoEMEA
            sValues += "|" + document.getElementById("chkbxSKGeoEMEA").checked.toString();

            //ProductBrandIDs
            sValues += "|" + getListValueSelections(document.getElementById("lstbxBrands"));

            //ServiceCategoryIDs
            sValues += "|" + getListValueSelections(document.getElementById("lstbxCats"));

            //OSSPIDs
            sValues += "|" + getListValueSelections(document.getElementById("lstbxOSSPs"));

            //KMATs
            sValues += "|" + document.getElementById("txtbxKMATS").value.toString().replace("|", "").replace(" ","");

            //SKUs
            sValues += "|" + document.getElementById("txtbxSKUS").value.toString().replace("|", "").replace(" ","");


            //AVs
            sValues += "|" + document.getElementById("txtbxAVS").value.toString().replace("|", "").replace(" ","");


            //RequireAllAVs
            sValues += "|" + document.getElementById("chkbxAllAVs").checked.toString();

            //SKs
            sValues += "|" + document.getElementById("txtbxSKs").value.toString().replace("|", "").replace(" ","");

            //SFPNS
            sValues += "|" + document.getElementById("txtbxSFPNS").value.toString().replace("|", "").replace(" ","");


            //LastAction
            var sLastAction = "";

            if (document.getElementById("chkbxAdded").checked) {
                sLastAction = "I";
            }

            if (document.getElementById("chkbxUpdated").checked) {
                if (sLastAction.length == 0) {
                    sLastAction = "U";
                }
                else {
                    sLastAction += ",U";
                }
            }

            sValues += "|" + sLastAction

            //ActionDateFrom
            sValues += "|" + document.getElementById("txtbxFromDate").value.toString().replace("|", "").replace(" ","");

            //ActionDateTo
            sValues += "|" + document.getElementById("txtbxToDate").value.toString().replace("|", "").replace(" ","");

            //Rows/Page
            sValues += "|"; //  + getListValueSelections(document.getElementById("ddlRowsPerPage"));
                        
            //SKUAVs
            sValues += "|" + document.getElementById("txtbxSKUAVS").value.toString().replace("|", "").replace(" ","");

            //RequireAllSKUAVs
            sValues += "|" + document.getElementById("chkbxAllSKUAVs").checked.toString();

            // Report Type
            sValues += "|" + getListValueSelections(document.getElementById("ddlReportType"));

            // Action Date Column
            sValues += "|" + getListValueSelections(document.getElementById("ddlActionDateColumn"));

            //SAs
            sValues += "|" + document.getElementById("txtbxSAs").value.toString().replace("|", "").replace(" ","");

            //COMPs
            sValues += "|" + document.getElementById("txtbxComps").value.toString().replace("|", "");

            //REPORT TYPE
            var sReportType = document.getElementById("ddlReportType").value;

            if (sReportType != "HTML") {

                switch (sReportType) {
                    case "EXCEL": sFileName = "FileName.xls";
                        break;
                    case "TEXT": sFileName = "FileName.txt";
                        break;
                }

                sFileName = window.prompt("Enter the name of the file to create:", sFileName);

                if ((sFileName == undefined) && (sFileName == null)) {
                    sFileName = "";
                } else if (sFileName.length == 0) {
                    sFileName = "";
                }

            } else {
                sFileName = "";
            }

            //Product Brand Division Filter
            if (document.filterForm.rdoProdDiv_1.checked) sValues += "|1";
            else if (document.filterForm.rdoProdDiv_2.checked) sValues += "|2";
            else sValues += "|0";

            //SPB Log Filter
            sValues += "|" + document.getElementById("chkbxSPBLogFilter").checked.toString();

            //SKU Log Filter
            sValues += "|" + document.getElementById("chkbxSKUAVLogFilter").checked.toString();
            

            
        } catch (e) {
            alert("getParameters - " + e.Description);
            sValues = e.Description;
            sFileName = "";
        }

        return sValues;

    }


    function getRequestMethod() {
        var sRM = null;
	var oReq=null;

        try {
            oReq = new XMLHttpRequest();
            sRM = "A";
        } catch (e) {
            sRM = "F";
        }
        finally {
            oReq = null;
        }

	if(sRM==null)
	{
	  sRM="F";
	}


        return sRM;

    }


    var sFileName = "";
    var oRequest = null;
    var sRequestMethod = "F"; // 'A' or 'F' (AJAX or FORM POST) 
    var bProcessing = false;
    var iMode = 0;


    function processTimeOut() // Process the time out error
    {
        updateStatus(document.getElementById("lblStatus"), "<b>ERROR - The Request or Response timed out.</b>", "red");
        oRequest.abort();
        oRequest = null;
        bProcessing = false;
        iMode = 0;
    }


    function showReport(sID) {
        var iWinID;
        var sProfID = getListValueSelections(document.getElementById("ddlRptProfile"));

        try {
            iWinID = window.open("processSKUBOMAS.aspx?m=1&rg=" + sID + "&PID="+sProfID+"&fn=" + sFileName, "", "toolbar=no,menubar=no,location=no,status=no,scrollbars=yes,resizable=yes", "");
        } catch (e) {
            alert(e.Message);
        }
    }


    function processAlternateResponse(sRG, bError, sMsg) {
        window.frames["postingFrame"].location = "requestSKUBOMAS.html";

        bProcessing = false;

        if (!bError) {
            updateStatus(document.getElementById("lblStatus"), "<b>Processing request, please wait...done.</b>", "black");
            showReport(sRG);
        } else {
            updateStatus(document.getElementById("lblStatus"), "<b>" + sMsg + "</b>", "red");
        }

    }


    function processResponse() // Process the returned response
    {
        // states
        /*
        const unsigned short UNSENT = 0;
        const unsigned short OPENED = 1;
        const unsigned short HEADERS_RECEIVED = 2;
        const unsigned short LOADING = 3;
        const unsigned short DONE = 4;
        readonly attribute unsigned short readyState;
        */
        // Add error handling ---> Use switch block
        try {
            if (bProcessing) {
                if (((oRequest.readyState == 2) || (oRequest.readyState == 4)) && ((oRequest.status == 200) || (oRequest.status == 304))) // 
                {
                    // Add error flag check
                    var sResponseString = oRequest.responseText;

                    bProcessing = false
                    oRequest = null;

                    if (sResponseString.indexOf("ERROR") == -1) {

                        updateStatus(document.getElementById("lblStatus"), "<b>Processing request, please wait...done.</b>", "black");

                        // Check current Mode and process accordingly
                        showReport(sResponseString);

                    } else {
                        updateStatus(document.getElementById("lblStatus"), "<b>" + sResponseString + "</b>", "red");
                        // alert(sResponseString);
                    }

                    // Consider object ref OR launch
                } else if ((oRequest.readyState == 4) && (oRequest.status != 200) && (oRequest.status != 304)) {
                    alert(oRequest.responseText);
                }
            }
        } catch (e) {
            updateStatus(document.getElementById("lblStatus"), "<b>" + e.Message + "</b>", "red");
            bProcessing = false
            oRequest = null;
        }

    }


    function sendRequest(mode) {

        // CONSIDER IMPLEMENTING JQUERY LIBRARY CALLS TO ACCOMMODATE OTHER BROWSERS IN THE FUTURE, ADD ERROR TRAPPING
        updateStatus(document.getElementById("lblStatus"), "<b>Processing request, please wait...</b>", "black");

        try {
            if (!bProcessing) {

                var bValidTextParameters = validateTextParameters();

                if (bValidTextParameters) {
                    var sParameters = getParameters();
                   
                    bProcessing = true;
                    iMode = mode;

                    // TESTING AND DEBUGGING PURPOSES ONLY
                    //------------------------------------------------
			        // sRequestMethod = getRequestMethod();
                    //------------------------------------------------
                    var sProfID = getListValueSelections(document.getElementById("ddlRptProfile"));

                    if (sRequestMethod == "A") {

                        oRequest = new XMLHttpRequest();
                        oRequest.onreadystatechange = processResponse;
                        oRequest.ontimeout = processTimeOut;
                        oRequest.open('POST', 'processSKUBOMAS.aspx?rm=A&m=' + mode.toString() + "&fn=" + sFileName);
                        oRequest.setRequestHeader('PARAMETERS', sParameters);
                        oRequest.setRequestHeader('PID', sProfID);
                        oRequest.send();

                    } else {

                        var oPostingFrame = window.frames["postingFrame"];
                        var oPostingForm = oPostingFrame.document.forms["requestForm"];
                        var oParametersElem = oPostingFrame.document.getElementById("PARAMETERS");
                        var oProfIDElem = oPostingFrame.document.getElementById("PID");
                        
                        oParametersElem.setAttribute("value", sParameters);
                        oProfIDElem.setAttribute("value", sProfID);
                        
                        oPostingForm.submit();
                    }
                }

            } else {
                alert("Currently processing another request...please try again, later.");
            }
        } catch (e) {

            if (sRequestMethod == "A") {
                // Attempt using alternate Request method
                sRequestMethod = "F";
                sendRequest(mode);
            } else {
                updateStatus(document.getElementById("lblStatus"), "<b>" + e.Message + "</b>", "red");
                bProcessing = false;
                oRequest = null;
                iMode = 0;
            }
        }

    }

    var iTimerID = null;

    function checkProfProcessFlag() {

        var oFlagElement = document.getElementById("profileFlag");

        try {
            if (oFlagElement.value == "true") {

                if (sRequestMethod == "F") {
                    var oPostingFrame = window.frames["postingFrame"];
                    var oPostingForm = oPostingFrame.document.forms["requestForm"];
                    var oParametersElem = oPostingFrame.document.getElementById("PARAMETERS");

                    if ((oParametersElem != null) && (oParametersElem != undefined)) {

                        if (iTimerID != null) {
                            window.clearTimeout(iTimerID);
                            iTimerID = null;
                        }

                        sendRequest(0);

                    } else if (iTimerID == null) {
                        iTimerID = window.setTimeout("checkProfProcessFlag()", 1000);
                    }
                } else {
                    sendRequest(0);
                }
            }
        } catch (e) {


            if (sRequestMethod == "A") {
                // Attempt using alternate Request method
                sRequestMethod = "F";
                sendRequest(0);
            } else {

                if (iTimerID == null) {
                    iTimerID = window.setTimeout("checkProfProcessFlag()", 1000);
                }
            }
        }
    }

    //************************************************************************************
</script>

<style type="text/css">

TH
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

H3
{
    FONT-SIZE: small;
    FONT-FAMILY: Verdana;
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}






.SKURow
{
    color:Black;
    background-color:Ivory;
    border-color:LightGray;
    border-width:1px;
    border-style:Solid;
    
    
}
.HeaderRow
{
    background-color:Ivory;
    border-color:LightGrey;
    border-width:1px;
    border-style:Solid;
    width:100%;
    
    
}
</style>
<link rel="stylesheet" type="text/css" href="../style/Excalibur.css" />
<link href="/style/wizard style.css" type="text/css" rel="stylesheet" />
</head>
<body>

<h2 style="font-family: Verdana; font-size: medium">SKU BOM Advanced Search</h2>
    <form id="filterForm" name="filterForm" runat="server" method="post" target="_self">
        <div>
                <asp:Table ID="tblProfileTable" runat="server" ForeColor="Black" >

                <asp:TableRow>
                <asp:TableCell Font-Size="Small" Font-Bold="True" HorizontalAlign="Left">Profiles:</asp:TableCell>
                <asp:TableCell HorizontalAlign="Left">
                    <asp:DropDownList ID="ddlRptProfile" runat="server" AutoPostBack="true" onchange="processProfileAction('A')"  DataTextField="ProfileName" DataValueField="ID" >
                    <asp:ListItem Value="0" Text="-- Select Profile --" ></asp:ListItem>
                    </asp:DropDownList>
                </asp:TableCell>
                <asp:TableCell>
                <asp:LinkButton ID="lnkBtnApplyProfile" runat="server" Text="Load" ToolTip="Load the selected Profile" Visible="False"></asp:LinkButton>&nbsp;
                <asp:LinkButton ID="lnkBtnAddProfile" runat="server" Text="Add" ToolTip="Create a Profile from the current Filter Options" OnClientClick="processProfileAction('I');"></asp:LinkButton>&nbsp;
                <asp:LinkButton ID="lnkBtnUpdateProfile" runat="server" Text="Update" ToolTip="Update the selected Profile with the current Filter Options" OnClientClick="processProfileAction('U');"></asp:LinkButton>&nbsp;
                <asp:LinkButton ID="lnkBtnDeleteProfile" runat="server" Text="Delete" ToolTip="Delete the selected Profile" OnClientClick="processProfileAction('D');"></asp:LinkButton>&nbsp;
                <asp:LinkButton ID="lnkBtnRenameProfile" runat="server" Text="Rename" ToolTip="Rename the selected Profile" OnClientClick="processProfileAction('R');"></asp:LinkButton>&nbsp;
                <asp:LinkButton ID="lnkBtnShareProfile" runat="server" Text="Share" ToolTip="Share the selected Profile" Visible="False" Enabled="False"></asp:LinkButton>
                <asp:LinkButton ID="lnkBtnRemoveSharedProfile" runat="server" Text="Remove" ToolTip="Remove the selected Shared Profile" Visible="False" OnClientClick="processProfileAction('RSP');"></asp:LinkButton>
                </asp:TableCell>
                <asp:TableCell>&nbsp;&nbsp;<asp:Label ID="lblProfileStatus" runat="server" Text="Select the desired Profile to retrieve applicable data."></asp:Label></asp:TableCell>
                </asp:TableRow>

                <asp:TableRow>
                <asp:TableCell></asp:TableCell>
                <asp:TableCell><asp:CheckBox ID="chkBxIncludeRemSProfs" AutoPostBack="True" Visible="False" Checked="False" runat="server" Enabled="False" Text="Include Removed Shared Profiles" /></asp:TableCell>
                <asp:TableCell></asp:TableCell>
                <asp:TableCell></asp:TableCell>
                </asp:TableRow>

                </asp:Table>            


        <asp:Table ID="tblFilterOptions" runat="server" BorderColor="Black" BorderStyle="Solid">

            <asp:TableRow VerticalAlign="Top">
            <asp:TableCell VerticalAlign="Top">

                        <asp:Table ID="tblFilterOptions0" runat="server" BorderColor="Black" CaptionAlign="Left" VerticalAlign="Top">

                            <asp:TableRow>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell Font-Underline="True" Font-Bold="True" Font-Size="Small" HorizontalAlign="Left" ForeColor="Red">PRODUCT FILTERS</asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>Service Family P/Ns:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxSFPNS" Width="200px" runat="server" ToolTip="Enter 1 or more Service Family Part Numbers separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>KMAT P/Ns:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxKMATS"  Width="200px" runat="server" ToolTip="Enter 1 or more KMAT Numbers separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>Products:</b>
                            <div id="PBCount"></div>

                            <asp:RadioButtonList ID="rdoProdDiv" AutoPostBack="True" runat="server">
                            <asp:ListItem Selected="True" Text="All" Value="0" ></asp:ListItem>
                            <asp:ListItem Selected="False" Text="Commercial" Value="1" ></asp:ListItem>
                            <asp:ListItem Selected="False" Text="Consumer" Value="2" ></asp:ListItem>
                            </asp:RadioButtonList>
                            
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:ListBox ID="lstbxBrands" onchange="updateListSelectionCount(this, document.getElementById('PBCount'));" runat="server" Height="150px" Width="200px" DataSourceID="dsBrands" DataTextField="Name" DataValueField="ID" SelectionMode="Multiple" ToolTip="Select 1 or more Product Brands">
                                </asp:ListBox>
                            </asp:TableCell>
                            </asp:TableRow>


                            <asp:TableRow>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell Font-Underline="True" Font-Bold="True" Font-Size="Small" HorizontalAlign="Left" ForeColor="Red">SKU FILTERS</asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>SKU P/Ns:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxSKUS" Width="200px" runat="server" ToolTip="Enter 1 or more SKU Numbers separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>SKU AVs:</b>
                            </asp:TableCell>
                            <asp:TableCell>
                            <asp:TextBox ID="txtbxSKUAVS" Width="200px" runat="server" ToolTip="Enter 1 or more SKU AV Numbers separated by a comma"></asp:TextBox><br />
                            <asp:CheckBox ID="chkbxAllSKUAVs" runat="server" Text="Require All SKU AVs"/><br />
                            <asp:CheckBox ID="chkbxSKUAVLogFilter" Name="chkbxSKUAVLogFilter" runat="server" Text="Apply to SKU Change Log" />
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>SKU Geos:</b></asp:TableCell>
                            <asp:TableCell VerticalAlign="Top">
                                <asp:CheckBox ID="chkbxGeoNA" runat="server" Text="NA"/>&nbsp;
                                <asp:CheckBox ID="chkbxGeoLA" runat="server" Text="LA"/>&nbsp;
                                <asp:CheckBox ID="chkbxGeoAPJ" runat="server" Text="APJ"/>&nbsp;
                                <asp:CheckBox ID="chkbxGeoEMEA" runat="server" Text="EMEA"/>
                            </asp:TableCell>
                            </asp:TableRow>
                            
                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell Font-Underline="True" Font-Bold="True" Font-Size="Small" HorizontalAlign="Left" ForeColor="Red">OSSP FILTERS</asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>OSSPs</b><div id="OSSPCount"></div></asp:TableCell>
                            <asp:TableCell>
                                <asp:ListBox ID="lstbxOSSPs" onchange="updateListSelectionCount(this, document.getElementById('OSSPCount'));" runat="server" Height="55px" Width="200px"  SelectionMode="Multiple" DataSourceID="odsListPartners" DataValueField="ID" DataTextField="Name" ToolTip="Select 1 or more OSSPs">
                                </asp:ListBox>
                            </asp:TableCell>
                            </asp:TableRow>

                        </asp:Table>
            </asp:TableCell>
            <asp:TableCell VerticalAlign="Top">
                        <asp:Table ID="tblFilterOptions1" runat="server" BorderColor="Black" CaptionAlign="Left">
                        
                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell Font-Underline="True" Font-Bold="True" Font-Size="Small" HorizontalAlign="Left" ForeColor="Red">SPARE KIT FILTERS</asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Spare Kits:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxSKs" Width="200px" runat="server" ToolTip="Enter 1 or more Spare Kit Numbers separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Spare Kit Geos:</b></asp:TableCell>
                            <asp:TableCell VerticalAlign="Top">
                                <asp:CheckBox ID="chkbxSKGeoNA" runat="server" Text="NA" />&nbsp;
                                <asp:CheckBox ID="chkbxSKGeoLA" runat="server" Text="LA"/>&nbsp;
                                <asp:CheckBox ID="chkbxSKGeoAPJ" runat="server" Text="APJ" />&nbsp;
                                <asp:CheckBox ID="chkbxSKGeoEMEA" runat="server" Text="EMEA"/>
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Categories:</b><div id="SCCount"></div></asp:TableCell>
                            <asp:TableCell>
                                <asp:ListBox ID="lstbxCats" onchange="updateListSelectionCount(this, document.getElementById('SCCount'));" runat="server" Height="55px" Width="200px" DataSourceID="dsCategories" DataValueField="ID" DataTextField="CategoryName" SelectionMode="Multiple" ToolTip="Select 1 or more Spare Kit Categories">
                                </asp:ListBox>
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Wrap="False" Font-Size="Small"><b>AV Qualifiers:</b>
                            </asp:TableCell>
                            <asp:TableCell>
                            <asp:TextBox ID="txtbxAVS" Width="200px" runat="server" ToolTip="Enter 1 or more Spare Kit AV Qualifiers separated by a comma"></asp:TextBox><br />
                            <asp:CheckBox ID="chkbxAllAVs" runat="server" Text="Require All AVs"/>
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>SubAssemblies:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxSAs" Width="200px" runat="server" ToolTip="Enter 1 or more Spare Kit Sub-Assemblies separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Components:</b></asp:TableCell>
                            <asp:TableCell><asp:TextBox ID="txtbxComps" Width="200px" runat="server" ToolTip="Enter 1 or more Spare Kit Components separated by a comma"></asp:TextBox></asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Top">
                            <asp:TableCell VerticalAlign="Top"></asp:TableCell>
                            <asp:TableCell>
                            <asp:CheckBox ID="chkbxSPBLogFilter" Name="chkbxSPBLogFilter" runat="server" Text="Apply to SPB Change Log" />
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell Font-Underline="True" Font-Bold="True" Font-Size="Small" HorizontalAlign="Left" ForeColor="Red">DATE FILTERS</asp:TableCell>
                            </asp:TableRow>                                                    

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Last Action</b></asp:TableCell>
                            <asp:TableCell VerticalAlign="Top">
                                <asp:CheckBox ID="chkbxAdded" runat="server" Text="Added"/>&nbsp;
                                <asp:CheckBox ID="chkbxUpdated" runat="server" Text="Updated"/>&nbsp;
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small"><b>Action Date Range</b></asp:TableCell>
                            <asp:TableCell  VerticalAlign="Top">
                            <asp:Textbox ID="txtbxFromDate" runat="server" Width="100px"></asp:Textbox>
                            <asp:LinkButton runat="server" ID="lnkBtnFromDate" OnClientClick="selectDate('txtbxFromDate')"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose Starting Date" border="0" align="middle" width="26" height="21"/></asp:LinkButton>&nbsp;<b>TO</b>&nbsp;<asp:Textbox ID="txtbxToDate" runat="server" Width="100px"></asp:Textbox><asp:LinkButton runat="server" ID="lnkBtnToDate" OnClientClick="selectDate('txtbxToDate')"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose Ending Date" border="0" align="middle" width="26" height="21"/></asp:LinkButton>
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow>
                            <asp:TableCell VerticalAlign="Top" Font-Size="Small" Visible="True"><b>Action Date Column</b></asp:TableCell>
                            <asp:TableCell  VerticalAlign="Top">
                            <asp:DropDownList ID="ddlActionDateColumn" runat="server" DataValueField="ColumnID" DataTextField="ColumnDesc" DataSourceID="dsDateColumns" Visible="True">
                            </asp:DropDownList>
                            </asp:TableCell>
                            </asp:TableRow>

                            <asp:TableRow VerticalAlign="Bottom">
                            <asp:TableCell></asp:TableCell>
                            <asp:TableCell>
                            <input type="button" id="btnRefresh" title="Refresh/Reload this page" value="Refresh" onclick="window.location=window.location;"/>
                            &nbsp;
                            <asp:Button runat="server" ID="btnReset" ToolTip="Reset Filter Options" Text="Reset" OnClientClick="resetForm()" />
                            &nbsp;
                            <input type="button" id="btnSubmitQuery" value="Submit" onclick="sendRequest(0)" title="Submit Query"/>
                            &nbsp;
                            <asp:DropDownList ID="ddlRowsPerPage" Visible="false" runat="server" onchange="updateStatus(document.getElementById('lblStatus'),'<b>Processing request, please wait...</b>','black')" AutoPostBack="True">
                            <asp:ListItem Value="10">10</asp:ListItem>
                            <asp:ListItem Value="20">20</asp:ListItem>
                            <asp:ListItem Value="30">30</asp:ListItem>
                            <asp:ListItem Value="40">40</asp:ListItem>
                            <asp:ListItem Value="50">50</asp:ListItem>
                            <asp:ListItem Value="60">60</asp:ListItem>
                            <asp:ListItem Value="70">70</asp:ListItem>
                            <asp:ListItem Value="80">80</asp:ListItem>
                            <asp:ListItem Value="90">90</asp:ListItem>
                            <asp:ListItem Value="100">100</asp:ListItem>
                            <asp:ListItem Value="0" Selected="True">ALL</asp:ListItem>
                            </asp:DropDownList>
                            
                            </asp:TableCell>
                            </asp:TableRow> 
                        </asp:Table>

            </asp:TableCell>
            <asp:TableCell VerticalAlign="Top" BorderStyle="Solid" BorderColor="Black">
                <asp:Table runat="server">
                    <asp:TableHeaderRow VerticalAlign="Top">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Center" Wrap="False" Font-Bold="True" Font-Underline="True" Font-Size="Small" ForeColor="Red">RESULT COLUMNS</asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    </asp:TableHeaderRow>
                    <asp:TableRow VerticalAlign="Top">
                        <asp:TableCell>
                        <b>Selected</b>&nbsp;<a href="javascript:removeColumnSelections()" title="Remove the selected column(s) from the list" style="text-decoration: none">[Remove]</a>
                        </asp:TableCell>
                        <asp:TableCell>
                        </asp:TableCell>
                        <asp:TableCell>
                        <b>Available</b>&nbsp;<a href="javascript:addSelections(document.getElementById('allColumns'), document.getElementById('selectedColumns'), null, null)" title="Add the selected Available column(s) to the Selected columns list" style="text-decoration: none">[Add]</a>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow VerticalAlign="Top">
                        <asp:TableCell>
                        <asp:ListBox ID="selectedColumns" name="selectedColumns" hiddenInputID="selectedCols" runat="server" Rows="8" ondblclick="addSelections(document.getElementById('selectedColumns'), document.getElementById('orderByCols'), '-0', '-1')" SelectionMode="Multiple" Height="175px" title="Double-Click an item to add it to the Sort columns list"></asp:ListBox>
                        </asp:TableCell>
                        <asp:TableCell VerticalAlign="Middle">
                        <a href="javascript:moveSelection(document.getElementById('selectedColumns'),1)" title="Move selected column(s) Up" style="text-decoration: none">[Up]</a><br /><br />
			            <a href="javascript:moveSelection(document.getElementById('selectedColumns'),0)" title="Move selected column Down" style="text-decoration: none">[Down]</a>
                        </asp:TableCell>
                        <asp:TableCell>
                        <asp:ListBox ID="allColumns" name="allColumns" DataSourceID="dsAvailableColumns" DataTextField="ColumnDesc" DataValueField="ColumnID" runat="server" Rows="8" ondblclick="addSelections(document.getElementById('allColumns'), document.getElementById('selectedColumns'), null, null)" SelectionMode="Multiple" Height="175px"  title="Double-Click an item to add it to the Selected columns list">
                        </asp:ListBox>
                        </asp:TableCell>
                    </asp:TableRow>

			        <asp:TableRow>
				        <asp:TableCell HorizontalAlign="center">
				        <a href="javascript:addSelections(document.getElementById('selectedColumns'), document.getElementById('orderByCols'), '-0', '-1');" title="Add Selected Column(s) to Sort List" style="text-decoration: none">[Add]</a>&nbsp;<a href="javascript:removeOrderColumnSelections();" title="Remove Selected Column(s) from Sort List" style="text-decoration: none">[Remove]</a>
				        </asp:TableCell>
				        <asp:TableCell HorizontalAlign="center" style="color:Red;font-size:Small;font-weight:bold;text-decoration:underline;white-space:nowrap;">SORT COLUMNS</asp:TableCell>			
				        <asp:TableCell HorizontalAlign="left"></asp:TableCell>
			        </asp:TableRow>

			        <asp:TableRow>
				        <asp:TableCell HorizontalAlign="center">
					        <select id="orderByCols" name="orderByCols" hiddenInputID="orderCols" multiple="multiple" onchange="updateOrderByColumn(this)"  style="height:80px;">
					        </select>
				        </asp:TableCell>
				        <asp:TableCell HorizontalAlign="left" VerticalAlign="top">
					        <input id="orderByAsc" type="radio" name="OrderByDir" value="0" onclick="updateCurrentOrderColumn('0')" />Ascending&nbsp;<input id="orderByDesc" type="radio" name="OrderByDir" value="1" onclick="updateCurrentOrderColumn('1')" />Descending<br/>
                            <a href="javascript:moveSelection(document.getElementById('orderByCols'),1)" title="Move selected Sort column(s) Up" style="text-decoration: none">[Up]</a><br /><br />
					        <a href="javascript:moveSelection(document.getElementById('orderByCols'),0)" title="Move selected Sort column Down" style="text-decoration: none">[Down]</a>
				        </asp:TableCell>
				        <asp:TableCell HorizontalAlign="left">
				        </asp:TableCell>
			        </asp:TableRow>

                    <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Center">Click the <b>Submit</b> button to apply</asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow Height="20px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow VerticalAlign="Bottom">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Table BorderColor="Black" BorderStyle="Solid" runat="server">
                            <asp:TableRow>
                            <asp:TableCell><b>Report Type:</b></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                            <asp:TableCell>
                            <asp:DropDownList ID="ddlReportType" runat="server">
                            <asp:ListItem Value="HTML" Text="HTML"></asp:ListItem>
                            <asp:ListItem Value="EXCEL" Text="Excel"></asp:ListItem>
                            <asp:ListItem Value="TEXT" Text="Text (Pipe Delimited)"></asp:ListItem>
                            </asp:DropDownList>
                            </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    </asp:TableRow>

                </asp:Table>
            </asp:TableCell>
            </asp:TableRow>

        </asp:Table>

        <asp:Label ID="lblStatus" runat="server" Text="Set the desired filter options and click the <b>Submit</b> button to retrieve applicable data."></asp:Label><br />
        <br />
        </div>
	<br />

	<div style="color:Red;font-size:Small;font-weight:bold;">
	NOTE: SKUs, Spare Kits, Subassemblies and Components that have a Cross Plant Status of C1 or C6, or a Material Type of ZWAR, are not included in these reports.<br/>
    	</div>        

        <asp:HiddenField ID="selectedCols" Value="4,2,3" runat="server" />
        <asp:HiddenField ID="orderCols" Value="4-0" runat="server" />
        <asp:HiddenField ID="profileName" Value="<%=Me.CurrProfileName %>" runat="server" />
        <asp:HiddenField ID="continueAction" Value="false" runat="server" />
        <asp:HiddenField ID="profileFlag" Value="false" runat="server" />
        <asp:HiddenField ID="resetFlag" Value="false" runat="server" />

    </form>
    
    <iframe id="postingFrame" name="postingFrame" src="requestSKUBOMAS.html" style="display:none"></iframe>

    <asp:SqlDataSource ID="dsBrands" runat="server" 
        ConnectionString="<%$ ConnectionStrings:PRSConnectionString %>" 
        SelectCommand="SELECT PB.ID AS ID, DOTSName+' > '+BRD.Abbreviation AS Name, PV.DevCenter AS DevCenter FROM ProductVersion PV WITH(NOLOCK) INNER JOIN Product_Brand PB WITH(NOLOCK) ON PV.ID=PB.ProductVersionID INNER JOIN Brand BRD WITH(NOLOCK) ON PB.BrandID=BRD.ID INNER JOIN ServiceFamilyDetails SFD WITH(NOLOCK) ON PV.ServiceFamilyPn=SFD.ServiceFamilyPn WHERE PV.ProductStatusID!=5 AND SFD.AutoPublishRsl=1 ORDER BY Name ASC">
    </asp:SqlDataSource>

    <asp:SqlDataSource ID="dsCategories" runat="server" 
        SelectCommand="SELECT ID, CategoryName FROM ServiceSpareCategory WITH(NOLOCK) ORDER BY CategoryName ASC" 
        ConnectionString="<%$ ConnectionStrings:PRSConnectionString %>"></asp:SqlDataSource>

    <asp:ObjectDataSource ID="odsListPartners" runat="server" 
            OldValuesParameterFormatString="original_{0}" SelectMethod="ListPartners" 
            TypeName="HPQ.Excalibur.Data" FilterExpression="active=1" >
            <SelectParameters>
                <asp:Parameter DefaultValue="1" Name="ReportType" Type="String" />
                <asp:Parameter DefaultValue="2" Name="PartnerTypeID" Type="String" />
            </SelectParameters>
    </asp:ObjectDataSource>

    <asp:SqlDataSource ID="dsAvailableColumns" runat="server" 
        SelectCommand="SELECT ColumnID, ColumnDesc FROM BTOSSASColumns WITH(NOLOCK) WHERE Active=1 OR DevOnly=1 ORDER BY OrderIndex ASC" 
        ConnectionString="<%$ ConnectionStrings:PRSConnectionString %>"></asp:SqlDataSource>

    <asp:SqlDataSource ID="dsDateColumns" runat="server" 
        SelectCommand="SELECT 0 AS ColumnID, '-- Select Date Column --' AS ColumnDesc UNION SELECT ColumnID, ColumnDesc FROM BTOSSASColumns WITH(NOLOCK) WHERE (Active=1 OR DevOnly=1) AND DateColumn=1 ORDER BY ColumnDesc ASC" 
        ConnectionString="<%$ ConnectionStrings:PRSConnectionString %>"></asp:SqlDataSource>

    <script type="text/javascript" language="javascript">
        
        /*********************************************************************************************************/
        // CONSIDER MOVING ALL OR SOME OF THE CALLS BELOW TO SERVER SIDE LOGIC
        /*********************************************************************************************************/
        
        updateListSelectionCount(document.getElementById("lstbxBrands"), document.getElementById("PBCount"));
        updateListSelectionCount(document.getElementById("lstbxCats"), document.getElementById("SCCount"));
        updateListSelectionCount(document.getElementById("lstbxOSSPs"), document.getElementById("OSSPCount"));

        populateResultParameters();

        //checkProfProcessFlag();

        /*********************************************************************************************************/
    </script>
   
    

</body>
</html>
