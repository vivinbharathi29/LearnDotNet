<%@ Language=VBScript %>

<%Response.Expires = 0%>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../../includes/client/json2.js"></script>
<script type="text/javascript" src="../../includes/client/json_parse.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
    FindFile.submit();	
}


function window_onload() {
    //FindFile.file1.click();

    //initialize modal dialog
    modalDialog.load();
}


function file1_onchange() {
	//FindFile.submit();
}
function cmdAddEmployee_onclick() {
    var strID = new Array();
    strID = window.showModalDialog("employee.asp", "", "dialogWidth:460px;dialogHeight:415px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) != "undefined") {
        ProgramInput.cboOwner.options[ProgramInput.cboOwner.length] = new Option(strID[1], strID[0]);
        ProgramInput.cboOwner.selectedIndex = ProgramInput.cboOwner.length - 1
    }
}


function cmdAddFamily_onclick() {
    var strID = new Array();
    strID = window.showModalDialog("family.asp", "", "dialogWidth:435px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) != "undefined") {
        ProgramInput.cboFamily.options[ProgramInput.cboFamily.length] = new Option(strID[1], strID[0]);
        ProgramInput.cboFamily.selectedIndex = ProgramInput.cboFamily.length - 1;
        ProgramInput.txtProductFamily.value = ProgramInput.cboFamily.options[ProgramInput.cboFamily.selectedIndex].text;
    }

}


function cmdPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPM);
}

function cmdPDEAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPDE);
}

function cmdAccessoryPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboAccessoryPM);
}

function cmdTDCCMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboTDCCM);
}

function cmdSMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSM);
}

function cmdToolsPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboToolsPM);
}

function cmdSEPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSEPM);
}

function cmdSEPEAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSEPE);
}

function cmdPINPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPINPM);
}

function cmdSETestLeadAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSETestLead);
}

function cmdSETestAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSETest);
}

function cmdWWANTestLeadAdd_onclick() {
    ChooseEmployee(ProgramInput.cboWWANTestLead);
}

function cmdODMTestLeadAdd_onclick() {
    ChooseEmployee(ProgramInput.cboODMTestLead);
}

function cmdBIOSLeadAdd_onclick() {
    ChooseEmployee(ProgramInput.cboBIOSLead);
}

function cmdCommHWPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboCommHWPM);
}

function cmdDKCAdd_onclick() {
    ChooseEmployee(ProgramInput.cboDKC);
}

function cmdProcessorPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboProcessorPM);
}

function cmdVideoMemoryPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboVideoMemoryPM);
}

function cmdGraphicsControllerPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboGraphicsControllerPM);
}

function cmdDocPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboDocPM);
}

function cmdPDEAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPDE);
}

function cmdComMarketingAdd_onclick() {
    ChooseEmployee(ProgramInput.cboComMarketing);
}

function cmdConsMarketingAdd_onclick() {
    ChooseEmployee(ProgramInput.cboConsMarketing);
}

function cmdSMBMarketingAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSMBMarketing);
}

function cmdPlatformDevelopmentAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPlatformDevelopment);
}

function cmdSupplyChainAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSupplyChain);
}

function cmdServiceAdd_onclick() {
    ChooseEmployee(ProgramInput.cboService);
}

function cmdFinanceAdd_onclick() {
    ChooseEmployee(ProgramInput.cboFinance);
}

function cmdMarketingOpsAdd_onclick() {
    ChooseEmployee(ProgramInput.cboMarketingOps);
}

function cmdPCAdd_onclick() {
    ChooseEmployee(ProgramInput.cboPC);
}

function cmdFactoryEngineerAdd_onclick() {
    ChooseEmployee(ProgramInput.cboFactoryEngineer);
}

function cmdSustainingSEPMAdd_onclick() {
    ChooseEmployee(ProgramInput.cboSustainingSEPM);
}

function SystemTeamAdd(objectId) {
    var obj = document.getElementById(objectId);
    ChooseEmployee(obj)
}


/*function old_cmdFinanceAdd_onclick() {
var strID = new Array();
strID = window.showModalDialog("employee.asp","","dialogWidth:460px;dialogHeight:342px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
if (typeof(strID) != "undefined")
{
ProgramInput.cboFinance.options[ProgramInput.cboFinance.length] = new Option(strID[1],strID[0]);
ProgramInput.cboFinance.selectedIndex = ProgramInput.cboFinance.length - 1
//ProgramInput.cboApprover.options[ProgramInput.cboApprover.length] = new Option(strID[1],strID[0]);
ProgramInput.cboTDCCM.options[ProgramInput.cboTDCCM.length] = new Option(strID[1],strID[0]);			
ProgramInput.cboPDE.options[ProgramInput.cboPDE.length] = new Option(strID[1],strID[0]);
ProgramInput.cboPM.options[ProgramInput.cboPM.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSEPM.options[ProgramInput.cboSEPM.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSM.options[ProgramInput.cboSM.length] = new Option(strID[1],strID[0]);
ProgramInput.cboComMarketing.options[ProgramInput.cboComMarketing.length] = new Option(strID[1],strID[0]);
ProgramInput.cboConsMarketing.options[ProgramInput.cboConsMarketing.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSMBMarketing.options[ProgramInput.cboSMBMarketing.length] = new Option(strID[1],strID[0]);
ProgramInput.cboPlatformDevelopment.options[ProgramInput.cboPlatformDevelopment.length] = new Option(strID[1],strID[0]);
ProgramInput.cboService.options[ProgramInput.cboService.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSupplyChain.options[ProgramInput.cboSupplyChain.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSEPE.options[ProgramInput.cboSEPE.length] = new Option(strID[1],strID[0]);
ProgramInput.cboPINPM.options[ProgramInput.cboPINPM.length] = new Option(strID[1],strID[0]);
ProgramInput.cboSETestLead.options[ProgramInput.cboSETestLead.length] = new Option(strID[1],strID[0]);
}
}
*/

function txtDescription_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function txtObjective_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;

}

function txtBaseUnit_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;

}

function txtOSSupport_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function txtImagePO_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;

}

function txtImageChanges_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function txtCertificationStatus_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function txtSWQAStatus_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function txtPlatformStatus_onkeypress() {
    if (window.event.keyCode == 42)
        window.event.keyCode = 8226;
}

function ModelExamples() {

    var strExample;

    strExample = "Be as specific as possible about the model because this information is communicated to the customers in the softpaq text file.\r\r"
    strExample = strExample + "FORMAT: [All ] <Brand> <Model> [ ,Model2] ... [ ,Modeln]\r\r"
    strExample = strExample + "EXAMPLES\r"
    strExample = strExample + "Evo N400c\r"
    strExample = strExample + "All Armada E500, M700, M300\r"
    strExample = strExample + "All Presario 7000 Series\r"
    strExample = strExample + "Presario 5012, 5013, 5015, 5890, 5287, 5425, 8974\r"

    window.alert(strExample);
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

function rblAffectedProduct_onclick(type) {
    var radio = document.getElementsByTagName("input");
    var cboMilestones = document.getElementById("cboMilestones");
    for (i = 0; i < radio.length; i++) {
        if ((radio[i].id == "rblAffectedProduct") && (type == 3)) {
            cboMilestones.disabled = false;
        }
        else if ((radio[i].id == "rblAffectedProduct") && (type != 3)) {
            cboMilestones.disabled = true;
        }
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

function cmdRTM1Date_onclick() {
    var strID;
    strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtRTM1Date.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) != "undefined") {
        ProgramInput.txtRTM1Date.value = strID;
    }
}

function cmdRTM2Date_onclick() {
    var strID;
    strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtRTM2Date.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) != "undefined") {
        ProgramInput.txtRTM2Date.value = strID;
    }
}

function cmdRTM3Date_onclick() {
    var strID;
    strID = window.showModalDialog("calDraw1.asp", ProgramInput.txtRTM3Date.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
    if (typeof (strID) != "undefined") {
        ProgramInput.txtRTM3Date.value = strID;
    }
}

function mouseover_Column() {
    event.srcElement.style.color = "red";
    event.srcElement.style.cursor = "hand";

}
function mouseout_Column() {
    event.srcElement.style.color = "black";
}

function TestPath(strID) {
    if (strID == 1) {
        if (ProgramInput.txtPDDPath.value == "")
            alert("PDD Path not specified.");
        else
            window.open(ProgramInput.txtPDDPath.value);
    }

    else if (strID == 2) {
        if (ProgramInput.txtSCMPath.value == "")
            alert("SCM Path not specified.");
        else
            window.open(ProgramInput.txtSCMPath.value);
    }
    else if (strID == 3) {
        if (ProgramInput.txtSTLPath.value == "")
            alert("STL Status Path not specified.");
        else
            window.open(ProgramInput.txtSTLPath.value);
    }
    else if (strID == 4) {
        if (ProgramInput.txtProgramMatrixPath.value == "")
            alert("Product Data Matrices Path not specified.");
        else
            window.open(ProgramInput.txtProgramMatrixPath.value);
    }
    else if (strID == 5) {
        if (ProgramInput.txtAccessoryPath.value == "")
            alert("Accessory Documents Path not specified.");
        else
            window.open(ProgramInput.txtAccessoryPath.value);
    }
}

function EnterSeries() {
    alert("This function is under development.");
}

function BrandCheck_onclick(ID) {
    var Result = 0;
    var strTemp = "";

    if (ProgramInput.txtID.value == "") {
        if (event.srcElement.checked)
            document.all("DivSeries" + ID).style.display = "";
        else
            document.all("DivSeries" + ID).style.display = "none";
    }
    else {
        if (!event.srcElement.checked) {
            Result = window.showModalDialog("BrandDeleteWarning.asp?ProductName=" + ProgramInput.txtProductFamily.value + " " + ProgramInput.txtVersion.value + "&BrandName=" + event.srcElement.BrandName, "", "dialogWidth:700px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
            if (Result == "1") {
                document.all("DivSeries" + ID).style.display = "none";
            }
            else {
                event.srcElement.checked = true;
                document.all("DivSeries" + ID).style.display = "";
            }
        }
        else
            document.all("DivSeries" + ID).style.display = "";

    }


}

function cboDevCenter_onchange() {
    if (ProgramInput.cboDevCenter.value != 2) {
        lblPOPM.innerHTML = "Configuration&nbsp;Manager:"
        lblTDCCM.innerHTML = "Program&nbsp;Office&nbsp;Manager:"
        if (ProgramInput.cboDCRDefaultOwner.selectedIndex == 0)
            ProgramInput.cboDCRDefaultOwner.selectedIndex = 1;
        // ProgramInput.cboTDCCM.style.display = "none"
        //	POPMRequired.style.display = "none";
        // ProgramInput.cmdTDCCMAdd.style.display = "none";
        // POPMConsOnly.style.display = "";
    }
    else {
        lblPOPM.innerHTML = "Program&nbsp;Office&nbsp;Manager:"
        lblTDCCM.innerHTML = "Configuration&nbsp;Manager:"
        if (ProgramInput.cboDCRDefaultOwner.selectedIndex == 0)
            ProgramInput.cboDCRDefaultOwner.selectedIndex = 2;
        //  ProgramInput.cboTDCCM.style.display = ""
        //	POPMRequired.style.display = "";
        // POPMConsOnly.style.display = "none";
        // ProgramInput.cmdTDCCMAdd.style.display = "";
    }
    if (ProgramInput.cboReleaseTeam.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
        ProgramInput.cboReleaseTeam.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;
    if (ProgramInput.cboPreinstall.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
        ProgramInput.cboPreinstall.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;

}

function ChooseEmployee(myControl) {
    modalDialog.open({ dialogTitle: 'Select Employee', dialogURL: 'ChooseEmployee.asp', dialogHeight: 200, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('1', 'employee_type');
    globalVariable.save(myControl.id, 'employee_dropdown');
    /*var ResultArray;
    ResultArray = window.showModalDialog("ChooseEmployee.asp", "", "dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")*/
}


function ChooseEmployee2() {
    modalDialog.open({ dialogTitle: 'Select Employee', dialogURL: 'ChooseEmployee.asp', dialogHeight: 200, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('2', 'employee_type');
    /*var ResultArray;
    ResultArray = window.showModalDialog("ChooseEmployee.asp", "", "dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")*/
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
        case "2":
            if (typeof (ResultArray) != "undefined") {
                if (ResultArray[0] != 0 && ResultArray[0] != ProgramInput.cboToolsPM.options[ProgramInput.cboToolsPM.selectedIndex].value) {
                    if (document.getElementById("ToolAccessRow" + ResultArray[0]) != null) {
                        document.getElementById("ToolAccessRow" + ResultArray[0]).style.display = "";
                        document.getElementById("chkToolAccessID" + ResultArray[0]).checked = true;
                    }
                    else {
                        var Row = document.all("ToolAccessTable").insertRow();

                        Row.bgColor = "#ffffff";
                        Row.id = "ToolAccessRow" + ResultArray[0]
                        var Cell = Row.insertCell();
                        Cell.noWrap = true;
                        Cell.className = "OTSComponentCell";
                        Cell.innerHTML = "<INPUT style=\"display:none\" type=\"checkbox\" checked id=chkToolAccessID" + ResultArray[0] + " name=chkToolAccessID value=\"" + ResultArray[0] + "\">" + ResultArray[1];

                        var Cell = Row.insertCell();
                        Cell.noWrap = true;
                        Cell.className = "OTSComponentCell";
                        Cell.innerHTML = "<a href=\"javascript: RemoveToolAccess(" + ResultArray[0] + ");\">Remove</a>";
                    }
                }
            }
            break;
        default: break;
    }
}

function RemoveToolAccess(strID) {
    document.all("ToolAccessRow" + strID).style.display = "none";
    document.getElementById("chkToolAccessID" + strID).checked = false;

}

function cboType_onchange() {
    if (ProgramInput.cboType.options[ProgramInput.cboType.selectedIndex].value == 2) {
        BrandRow.style.display = "none";
        OSRow.style.display = "none";
        ApproverRow.style.display = "none";
        ActivitiesRow.style.display = "none";
        CommodityRow.style.display = "none";
        MDARow.style.display = "none";
        RegModelRow.style.display = "none";
        RegModelRow2.style.display = "none";
        ToolPMRow.style.display = "";
        ToolPMRow2.style.display = "";
        ReleaseRow.style.display = "none";
        PreinstallRow.style.display = "none";
        DevCenterRow.style.display = "none";
        DistributionRow.style.display = "none";
        NotificationRow.style.display = "";
    }
    else {
        BrandRow.style.display = "";
        OSRow.style.display = "";
        ApproverRow.style.display = "";
        ActivitiesRow.style.display = "";
        CommodityRow.style.display = "";
        MDARow.style.display = "";
        RegModelRow.style.display = "";
        RegModelRow2.style.display = "";
        ToolPMRow.style.display = "none";
        ToolPMRow2.style.display = "none";
        ReleaseRow.style.display = "";
        PreinstallRow.style.display = "";
        DevCenterRow.style.display = "";
        DistributionRow.style.display = "";
        NotificationRow.style.display = "none";
    }
    SelectTab(CurrentState);
}

function cboFamily_onchange() {
    ProgramInput.txtProductFamily.value = ProgramInput.cboFamily.options[ProgramInput.cboFamily.selectedIndex].text;
}


function AddComponents() {
    var PDMID = ProgramInput.cboPlatformDevelopment.options[ProgramInput.cboPlatformDevelopment.selectedIndex].value;
    var SEPMID = ProgramInput.cboSEPM.options[ProgramInput.cboSEPM.selectedIndex].value;
    var PINPMID = ProgramInput.cboPINPM.options[ProgramInput.cboPINPM.selectedIndex].value;
    var PEID = ProgramInput.cboSEPE.options[ProgramInput.cboSEPE.selectedIndex].value;
    var SC = ProgramInput.cboSupplyChain.options[ProgramInput.cboSupplyChain.selectedIndex].value;
    var ProductID = ProgramInput.txtID.value;
    var strMissingOwners = ""

    if (PDMID == 0 || PDMID == "")
        strMissingOwners = ", Platform Development Manager"
    if (SEPMID == 0 || SEPMID == "")
        strMissingOwners = strMissingOwners + ", SE PM"
    if (PINPMID == 0 || PINPMID == "")
        strMissingOwners = strMissingOwners + ", PIN PM"
    if (PEID == 0 || PEID == "")
        strMissingOwners = strMissingOwners + ", SE PE"
    if (SC == 0 || SC == "")
        strMissingOwners = strMissingOwners + ", Supply Chain"

    if (strMissingOwners != "") {
        strMissingOwners = strMissingOwners.substr(2);
        alert("You must assign the following people on the System Team tab to continue: " + strMissingOwners);
    }
    else {
        var strID;
        strID = window.showModalDialog("AddOTSComponents.asp?ProductID=" + ProductID + "&PDMID=" + PDMID + "&SEPMID=" + SEPMID + "&PINPMID=" + PINPMID + "&PEID=" + PEID + "&SC=" + SC, "", "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) != "undefined") {
            if (strID != "") {
                ProgramInput.txtMissingComponents.value = "0";
                OTSAddComponentMessage1.style.display = "none";
                OTSAddComponentMessage2.style.display = "none";
                OTSAddComponentMessage3.style.display = "none";
                OTSAddComponentTable.innerHTML = strID;
            }
        }
    }

}

function EditOTSCoreTeam(strID, strOTSComponentID, strCoreTeamID) {
    /*ShowPropertiesDialog("mobilese/today/ChooseComponentCoreTeam.asp?ID=" + strID + "&RoleID=3&CoreTeamID=" + strCoreTeamID, "Choose Component Core Team", 400, 200);*/
    modalDialog.open({ dialogTitle: 'Component Core Team', dialogURL: 'ChooseComponentCoreTeam.asp?OTSComponentID=' + strOTSComponentID + '&RoleID=3&CoreTeamID=' + strCoreTeamID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('1', 'role_type');
    globalVariable.save(strID, 'role_id');
    globalVariable.save(strOTSComponentID, 'role_otscomponentid');
    /*var strResult;
    strResult = window.showModalDialog("ChooseComponentCoreTeam.asp?OTSComponentID=" + strOTSComponentID + "&RoleID=3&CoreTeamID=" + strCoreTeamID, "", "dialogWidth:400px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
}

function EditOTSPM(strID, strOwnerID) {
    modalDialog.open({ dialogTitle: 'Component Owner', dialogURL: 'ChooseComponentOwner.asp?ID=' + strID + '&RoleID=1&OwnerID=' + strOwnerID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('2', 'role_type');
    globalVariable.save(strID, 'role_id');
    /*var strResult;
    strResult = window.showModalDialog("ChooseComponentOwner.asp?ID=" + strID + "&RoleID=1&OwnerID=" + strOwnerID, "", "dialogWidth:400px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
}

function EditOTSDeveloper(strID, strOwnerID) {
    modalDialog.open({ dialogTitle: 'Component Owner', dialogURL: 'ChooseComponentOwner.asp?ID=' + strID + '&RoleID=2&OwnerID=' + strOwnerID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('3', 'role_type');
    globalVariable.save(strID, 'role_id');
    /*var strResult;
    strResult = window.showModalDialog("ChooseComponentOwner.asp?ID=" + strID + "&RoleID=2&OwnerID=" + strOwnerID, "", "dialogWidth:400px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
}

function EditComponentResults() {
    var iType;
    var strID;
    var strOTSComponentID;
    var strResult;

    strResult = modalDialog.getArgument('role_query_array');
    strResult = JSON.parse(strResult);

    iType = globalVariable.get('role_type');

    switch (iType) {
        case "1":
            strID = globalVariable.get('role_id');
            strOTSComponentID = globalVariable.get('role_otscomponentid');
            if (typeof (strResult) != "undefined") {
                if (strResult[1] != "" && strResult[0] != "")
                    /*Modified By: 02/22/2016 JMalichi - Adding the missing OTSComponentID to the the link below*/
                    document.all("OTSCoreTeam" + strID).innerHTML = "<a href='javascript: EditOTSCoreTeam(" + strID + "," + strOTSComponentID + "," + strResult[0] + ")'>" + strResult[1] + "</a>";
                else
                    alert("Unable to update the core team.");
            }
            break;
        case "2":
            strID = globalVariable.get('role_id');
            if (typeof (strResult) != "undefined") {
                if (strResult[1] != "" && strResult[0] != "")
                    document.all("OTSPM" + strID).innerHTML = "<a href='javascript: EditOTSPM(" + strID + "," + strResult[0] + ")'>" + strResult[1] + "</a>";
                else
                    alert("Unable to update the selected owner.");
            }
            break;
        case "3":
            strID = globalVariable.get('role_id');
            if (typeof (strResult) != "undefined") {
                if (strResult[1] != "" && strResult[0] != "")
                    document.all("OTSDeveloper" + strID).innerHTML = "<a href='javascript: EditOTSDeveloper(" + strID + "," + strResult[0] + ")'>" + strResult[1] + "</a>";
                else
                    alert("Unable to update the selected owner.");
            }
            break;
        default: break;
    }

}


function SelectSystemBoardID() {
    modalDialog.open({ dialogTitle: 'Select System Board', dialogURL: 'ProductID.asp?TypeID=1&IDList=' + encodeURI(ProgramInput.txtSystemBoardComments.value.replace("\"", "%22")), dialogHeight: 500, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('1', 'id_type');
    /*var ResultArray;
    ResultArray = window.showModalDialog("ProductID.asp?TypeID=1&IDList=" + encodeURI(ProgramInput.txtSystemBoardComments.value.replace("\"", "%22")), "", "dialogWidth:400px;dialogHeight:370px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
}

function SelectMachinePNPID() {
    modalDialog.open({ dialogTitle: 'Select PnP', dialogURL: 'ProductID.asp?TypeID=2&IDList=' + ProgramInput.txtMachinePNPComments.value.replace("\"", "%22"), dialogHeight: 500, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
    globalVariable.save('2', 'id_type');
    /*var ResultArray;
    ResultArray = window.showModalDialog("ProductID.asp?TypeID=2&IDList=" + ProgramInput.txtMachinePNPComments.value.replace("\"", "%22"), "", "dialogWidth:400px;dialogHeight:370px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
}

function GetProductIDResult() {
    ResultArray = modalDialog.getArgument('product_id_array');
    ResultArray = JSON.parse(ResultArray);

    iTypeID = globalVariable.get('id_type');

    switch (iTypeID) {
        case "1":
            if (typeof (ResultArray) != "undefined") {
                //if (ResultArray[0] != 0)
                ProgramInput.txtSystemBoardID.value = ResultArray[0];
                if (ResultArray[0] == "")
                    SystemboardLink.innerText = "Add ID";
                else
                    SystemboardLink.innerText = ResultArray[0].replace(/,/g, ", ");
                ProgramInput.txtSystemBoardComments.value = ResultArray[1];
            }
            break;
        case "2":
            if (typeof (ResultArray) != "undefined") {
                //if (ResultArray[0] != 0)
                ProgramInput.txtMachinePNPID.value = ResultArray[0];
                if (ResultArray[0] == "")
                    MachinePNPLink.innerText = "Add ID";
                else
                    MachinePNPLink.innerText = ResultArray[0].replace(/,/g, ", ");
                ProgramInput.txtMachinePNPComments.value = ResultArray[1];
            }
            break;
        default: break;
    }
}

function chkCMTAll_onclick() {
    var i;

    if (typeof (ProgramInput.chkCMT.length) == "undefined")
        ProgramInput.chkCMT.checked = ProgramInput.chkCMTAll.checked;
    else
        for (i = 0; i < ProgramInput.chkCMT.length; i++)
        ProgramInput.chkCMT[i].checked = ProgramInput.chkCMTAll.checked;
}

/*function DeactivateCMTComponents() {
var i;
var IDList="";
var SelCount=0;
	
if(typeof(ProgramInput.chkCMT.length)=="undefined")
{
if (ProgramInput.chkCMT.checked)
IDList = ",'" + ProgramInput.chkCMT.value + "'";
}
else
{
for (i=0;i<ProgramInput.chkCMT.length;i++)
if (ProgramInput.chkCMT[i].checked)
{
if (SelCount>100)
{
ProgramInput.chkCMT[i].checked = false;
}
else
{
IDList = IDList + ",'" + ProgramInput.chkCMT[i].value + "'";
SelCount = SelCount + 1;
}
}	
}
		
if (IDList=="" )
alert("You must check at lest one component to deactivate.");
else if (SelCount > 100)
alert("You can only select up to 100 components at a time.  The excess components have been unchecked.  Please click the link again to deactivate those.");
else
{
IDList = IDList.substring(1);
if(confirm("Are you sure you want to deactivate these CMT components.  This update can not be undone."))
{
var strID;
strID = window.showModalDialog("DeactivateCMTComponents.asp?ProductID=" + ProgramInput.txtID.value + "&Components=" + IDList,"","dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
if (typeof(strID) != "undefined")
{
OTSAddComponentTable.innerHTML = strID;
}		
}
}	
}
*/

function cboPhase_onchange() {

    ProgramInput.chkEnableDCR.disabled = false;

    if (ProgramInput.cboPhase.selectedIndex < 3)
        ROMText.innerHTML = "Current&nbsp;ROM";
    else
        ROMText.innerHTML = "Current&nbsp;Factory&nbsp;ROM";

    if (ProgramInput.cboPhase.selectedIndex > 3) {
        ProgramInput.chkEnableDeliverables.checked = false;
        ProgramInput.chkEnableDCR.checked = false;
        ProgramInput.chkEnableImages.checked = false;
        ProgramInput.chkEnableSMR.checked = false;
    }
    else if (ProgramInput.cboPhase.selectedIndex > 2) {
        ProgramInput.chkEnableSMR.checked = false;
        ProgramInput.chkEnableDCR.checked = false;
        ProgramInput.chkEnableImages.checked = false;
    }
    else if (ProgramInput.cboPhase.selectedIndex == 1 || ProgramInput.cboPhase.selectedIndex == 2) {
        ProgramInput.chkEnableDCR.checked = true;
        if (ProgramInput.cboPhase.selectedIndex == 1)
            ProgramInput.chkEnableImages.checked = true;
        else
            ProgramInput.chkEnableImages.checked = false;

        if (ProgramInput.chkEnableSMR.InitialValue == "")
            ProgramInput.chkEnableSMR.checked = false;
        else
            ProgramInput.chkEnableSMR.checked = true;

        if (ProgramInput.chkEnableDeliverables.InitialValue == "")
            ProgramInput.chkEnableDeliverables.checked = false;
        else
            ProgramInput.chkEnableDeliverables.checked = true;


    }
    else if (ProgramInput.cboPhase.selectedIndex == 0) {
        ProgramInput.chkEnableDCR.disabled = true;
        ProgramInput.chkEnableDCR.checked = false;
        ProgramInput.chkEnableImages.checked = true;

        if (ProgramInput.chkEnableSMR.InitialValue == "")
            ProgramInput.chkEnableSMR.checked = false;
        else
            ProgramInput.chkEnableSMR.checked = true;

        if (ProgramInput.chkEnableDeliverables.InitialValue == "")
            ProgramInput.chkEnableDeliverables.checked = false;
        else
            ProgramInput.chkEnableDeliverables.checked = true;


    }
    else {
        if (ProgramInput.chkEnableSMR.InitialValue == "")
            ProgramInput.chkEnableSMR.checked = false;
        else
            ProgramInput.chkEnableSMR.checked = true;

        if (ProgramInput.chkEnableDeliverables.InitialValue == "")
            ProgramInput.chkEnableDeliverables.checked = false;
        else
            ProgramInput.chkEnableDeliverables.checked = true;

        if (ProgramInput.chkEnableImages.InitialValue == "")
            ProgramInput.chkEnableImages.checked = false;
        else
            ProgramInput.chkEnableImages.checked = true;

        if (ProgramInput.chkEnableDCR.InitialValue == "")
            ProgramInput.chkEnableDCR.checked = true;
        else
            ProgramInput.chkEnableDCR.checked = true;
    }
}

function cmdAddApprover_onclick() {
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtDCRApproverList.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtDCRApproverList', 'email_field');
}

function cmdAddNotification_onclick() {
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtActionNotifyList.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtActionNotifyList', 'email_field');
}

function cmdAddDistribution_onclick() {
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtDistribution.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtDistribution', 'email_field');
}

function cmdAddCvrBuildDist_onclick() {
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtCvrBuildDist.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtCvrBuildDist', 'email_field');
}

function cmdAddCvrReleaseDist_onclick() {
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../../Email/AddressBook.asp?AddressList=' + ProgramInput.txtCvrReleaseDist.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtCvrReleaseDist', 'email_field');
}

function ChooseNewBrand(strID) {
    var i;
    var strIDs = "";
    var strResult = "";

    if (ProgramInput.txtBrandsLoaded.value != "")
        strIDs = "," + ProgramInput.txtBrandsLoaded.value

    for (i = 0; i < ProgramInput.chkBrands.length; i++) {
        if (ProgramInput.chkBrands[i].checked || document.all("Brand" + ProgramInput.chkBrands[i].value).style.display != "")
            strIDs = strIDs + "," + ProgramInput.chkBrands[i].value;
    }
    if (strIDs != "")
        strIDs = strIDs.substring(1);

    modalDialog.open({ dialogTitle: 'Brand', dialogURL: 'BrandUpdate.asp?ProductID=' + ProgramInput.txtID.value + '&BrandID=' + strID + '&ExcludeIDList=' + strIDs + '', dialogHeight: 300, dialogWidth: 450, dialogResizable: false, dialogDraggable: true });
    globalVariable.save(strID, 'brand_id');
}

function ChoosenewBrandResult() {
    var strID;
    var strResult;

    strResult = modalDialog.getArgument('brand_update_result');
    strID = globalVariable.get('brand_id');

    if (typeof (strResult) != "undefined") {
        for (i = 0; i < ProgramInput.chkBrands.length; i++)
            if (ProgramInput.chkBrands[i].value == strID)
                ProgramInput.chkBrands[i].checked = false;
        document.all("Brand" + strID).style.display = "none";
        document.all("DivSeries" + strID).style.display = "none";

        document.all("DivSeries" + strResult).style.display = "";
        for (i = 0; i < ProgramInput.chkBrands.length; i++)
            if (ProgramInput.chkBrands[i].value == strResult) {
                ProgramInput.chkBrands[i].checked = true;
                ProgramInput.chkBrands[i].disabled = true;
            }
        document.all("txtSeriesA" + strResult).value = document.all("txtSeriesA" + strID).value;
        document.all("txtSeriesB" + strResult).value = document.all("txtSeriesB" + strID).value;
        document.all("txtSeriesC" + strResult).value = document.all("txtSeriesC" + strID).value;
        document.all("txtSeriesD" + strResult).value = document.all("txtSeriesD" + strID).value;
        if (document.getElementById("txtSeriesE" + strID)) {
            document.all("txtSeriesE" + strResult).value = document.all("txtSeriesE" + strID).value;
        }
        if (document.getElementById("txtSeriesF" + strID)) {
            document.all("txtSeriesF" + strResult).value = document.all("txtSeriesF" + strID).value;
        }
        document.getElementById('DIV3').scrollTop = document.all("Brand" + strResult).offsetTop - 20;
        document.all("txtBrandFrom").value = strID;
        document.all("txtBrandTo").value = strResult;
    }
}
function SelectedSites(ID) {
    // Allow users to add RCTO sites only after saving product - task 16794
    if (ID == 0) {
        alert('Add Site after saving the product')
    } else {
        modalDialog.open({ dialogTitle: 'Sites', dialogURL: 'ProductSite.asp?ID=' + ID + '&Sites=' + ProgramInput.txtRCTOSites.value + '', dialogHeight: 400, dialogWidth: 600, dialogResizable: false, dialogDraggable: true });
    }
}

function SelectedSitesResult(strResult) {
    if (typeof (strResult) != "undefined") {
        ProgramInput.txtRCTOSites.value = strResult;
        if (strResult == "")
            SiteLink.innerText = "Add Site";
        else
            SiteLink.innerText = strResult;
    }
}
function SelectedCycles(ID) {
    modalDialog.open({ dialogTitle: 'Group', dialogURL: 'ProductCycle.asp?ID=' + ID + '', dialogHeight: 400, dialogWidth: 600, dialogResizable: false, dialogDraggable: true });
    /*var strResult;
    var ResultArray;
    strResult = window.showModalDialog("ProductCycle.asp?ID=" + ID, "", "dialogWidth:500px;dialogHeight:350px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");*/
}

function SelectedCyclesResult(strResult) {
    if (typeof (strResult) != "undefined") {
        if (ID == 0) {
            ResultArray = strResult.split("|");
            if (ResultArray[0] == "") {
                CycleLink.innerText = "Add to Product Group";
                ProgramInput.txtAddCycle.value = "";
            }
            else {
                CycleLink.innerText = ResultArray[0];
                ProgramInput.txtAddCycle.value = ResultArray[1];
            }
        }
        else {
            if (strResult == "")
                CycleLink.innerText = "Add to Product Group";
            else
                CycleLink.innerText = strResult;
        }
    }
}


//-->
</SCRIPT>
</HEAD>
<body onload="return window_onload()">
<form action="MobileSE/Today/ProgramSave.asp" method="post" name="ProgramInput">
<table border="1" cellpadding="2" cellspacing="0" style="width:100%; border-collapse:collapse;" bgcolor="cornsilk" bordercolor="tan">
    <tr bgcolor="Wheat">
        <td>
            <b>OTS Common Components - Factory</b>
        </td>
    </tr>
    <%
    dim cn
	dim rs
	Dim scString
	
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
	
	cnString =Session("PDPIMS_ConnectionString")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	
    if request("PVID") = "" then
	    rs.Open "spListProductOTSComponents 0",cn,adOpenStatic
    else
	    rs.Open "spListProductOTSComponents " & clng(request("PVID")) & "",cn,adOpenStatic
    end if
    rs.Filter="errortype='Factory'"
    if rs.EOF and rs.BOF then
        %>
        <tr>
            <td>
                <font face="verdana" size="2">
                    <div id="OTSAddComponentTable">
                        none</div>
                </font>
            </td>
        </tr>
    <%else%>
    <tr>
        <td>
            <div id="OTSAddComponentTable">
                <table style="width: 100%; border-collapse:collapse;" cellpadding="2" cellspacing="0">
                    <tr>
                        <td nowrap>
                            <%if ShowCMTRemoveLink then%>
                            <input style="width: 16; height: 16" type="checkbox" id="chkCMTAll" name="chkCMTAll"
                                language="javascript" onclick="return chkCMTAll_onclick()">
                            <%end if%>
                            <font face="verdana" size="1"><b>Source</b>&nbsp;&nbsp;</font>
                        </td>
                        <td>
                            <font face="verdana" size="1"><b>Err&nbsp;Type</b></font>
                        </td>
                        <td>
                            <font face="verdana" size="1"><b>Category</b></font>
                        </td>
                        <td>
                            <font face="verdana" size="1"><b>Component</b></font>
                        </td>
                        <td>
                            <font face="verdana" size="1"><b>PM</b></font>
                        </td>
                        <td>
                            <font face="verdana" size="1"><b>Developer</b></font>
                        </td>
                    </tr>
                    <%		do while not rs.EOF
                    if trim(rs("ErrorType")) = "Factory" then
	                if rs("ID") = 0 then
		                Response.Write "<TR bgcolor=Lavender><td class=OTSComponentCell><font face=verdana size=1><INPUT style=""Display:none;WIDTH:16;Height:16"" value=""" & rs("Partnumber") & """ type=""checkbox"" id=chkCMT name=chkCMT>CMT</font></td>"
	                else
		                Response.Write "<TR bgcolor=white><td class=OTSComponentCell><font face=verdana size=1>Excalibur&nbsp;&nbsp;</font></td>"
	                end if
                    %>
                    <td class="OTSComponentCell">
                        <font face=verdana size=1><%=rs("ErrorType")%></font>
                    </td>
                    <td class="OTSComponentCell">
                        <font face=verdana size=1><%=rs("category")%></font>
                    </td>
                    <td class="OTSComponentCell">
                       <font face=verdana size=1><%=rs("Component")%></font>
                    </td>
                    <%if rs("ID") = 0 then%>
                    <td class="OTSComponentCell">
                        <font face=verdana size=1><%=rs("PM")&""%></font>
                    </td>
                    <td class="OTSComponentCell">
                        <font face=verdana size=1><%=rs("Developer")&""%></a></font>
                    </td>
                    <%else%>
                    <td id='OTSPM<%=trim(rs("ID"))%>' class="OTSComponentCell">
                        <font face=verdana size=1><a href="javascript: EditOTSPM(<%=rs("ID")%>,<%=rs("PMID")%>)">
                            <%=longname(rs("PM")&"")%></a></font>
                    </td>
                    <td id='OTSDeveloper<%=trim(rs("ID"))%>' class="OTSComponentCell">
                        <font face=verdana size=1><a href="javascript: EditOTSDeveloper(<%=rs("ID")%>,<%=rs("DeveloperID")%>)">
                            <%=longname(rs("Developer")&"")%></a></font>
                    </td>
                    <%end if%>
            </tr>
            <%end if%>
      <%rs.MoveNext
      loop
     end if
%>
</table>
</form>
</BODY>
</HTML>
