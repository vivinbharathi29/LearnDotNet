<%@  language="VBScript" %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  Response.CodePage = 65001
  Response.Charset="UTF-8"

  Dim AppRoot
  AppRoot = Session("ApplicationRoot")	
%>
<html>
<head>
    <meta http-equiv="Expires" content="0">
    <meta http-equiv="Cache-Control" content="no-cache">
    <meta http-equiv="Pragma" content="no-cache">
    <meta name="VI60_defaultClientScript" content="JavaScript">
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
    <!-- #include file="../../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="../../includes/client/json2.js"></script>
    <script type="text/javascript" src="../../includes/client/json_parse.js"></script>
    <script type="text/javascript">
        $(function () {
            $("#iframeDialog").dialog({
                modal: true,
                autoOpen: false,
                width: 900,
                height: 800,
                close: function () {
                    $("#modalDialog").attr("src", "about:blank");
                },
                resizable: false

            });
        });

        function ShowIframeDialog() {
            $("#iframeDialog iframe").attr("width", $("#iframeDialog").dialog("option", "width") - 50);
            $("#iframeDialog iframe").attr("height", $("#iframeDialog").dialog("option", "height") - 50);
            $("#iframeDialog iframe").attr("src", "../Agency/Agency.asp?ID=5000");
            $("#iframeDialog").dialog("option", "title", "Agency Status");
            $("#iframeDialog").dialog("open");
        }

        function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {
            if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
            if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
            $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
            $("#modalDialog").attr("width", "98%");
            $("#modalDialog").attr("height", "98%");
            $("#modalDialog").attr("src", QueryString);
            $("#iframeDialog").dialog("option", "title", Title);
            $("#iframeDialog").dialog("open");
        }

        function ClosePropertiesDialog(strID) {
            $("#iframeDialog").dialog("close");
            if (typeof (strID) != "undefined") window.location.reload(true);
        }
        function ClosePropertiesDialog_fromClone(strID) {

            $("#iframeDialog").dialog("close");
            if (typeof (strID) != "undefined") {
                if (strID == txtID.value)
                    window.location.reload(true);
                else
                    window.location = "/Excalibur/pmview.asp?ID=" + strID + "&Class=" + txtClass.value;
            }
        }
        function CloseIframeDialog() {
            $("#iframeDialog").dialog("close");
        }
    </script>

    <script id="clientEventHandlersJS" language="javascript">
<!--


        var CurrentState;

        function ProcessState() {
            var steptext;

            switch (CurrentState) {
                case "General":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Specification"; //"Specification"; 
                    tabGeneral.style.display = "";
                    tabSystemTeam.style.display = "none";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Enter the Specification for this Requirement.";

                    window.scrollTo(0, 0);
                    break;

                case "Files":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Specification"; //"Specification"; 
                    tabFiles.style.display = "";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Enter the Specification for this Requirement.";

                    window.scrollTo(0, 0);
                    break;

                case "Access":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "none";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "";
                    if (ProgramInput.cboToolsPM.selectedIndex == 0)
                        PMAccessCell.innerHTML = "Project Manager";
                    else
                        PMAccessCell.innerHTML = ProgramInput.cboToolsPM.options[ProgramInput.cboToolsPM.selectedIndex].text;

                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";

                    window.scrollTo(0, 0);
                    break;

                case "SystemTeam":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";

                    window.scrollTo(0, 0);
                    break;
                case "OTS":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "none";
                    tabOTS.style.display = "";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";


                    var PDMID = ProgramInput.cboPlatformDevelopment.options[ProgramInput.cboPlatformDevelopment.selectedIndex].value;
                    var SEPMID = ProgramInput.cboSEPM.options[ProgramInput.cboSEPM.selectedIndex].value;
                    var PINPMID = ProgramInput.cboPINPM.options[ProgramInput.cboPINPM.selectedIndex].value;
                    var PEID = ProgramInput.cboSEPE.options[ProgramInput.cboSEPE.selectedIndex].value;
                    var SC = ProgramInput.cboSupplyChain.options[ProgramInput.cboSupplyChain.selectedIndex].value;
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

                    if (strMissingOwners != "" && ProgramInput.txtMissingComponents.value != "0") {
                        OTSAddComponentMessage1.innerHTML = "<BR><BR>You must populate the following Roles on the System Team tab before you can add the standard OTS Components: <BR><BR>&nbsp;&nbsp;-&nbsp;" + strMissingOwners.substring(2)
                        OTSAddComponentMessage1.style.display = "";
                        OTSAddComponentMessage2.style.display = "none";
                        OTSAddComponentMessage3.style.display = "none";
                    }
                    else {
                        OTSAddComponentMessage1.style.display = "none";
                        if (ProgramInput.txtMissingComponents.value == "0") {
                            OTSAddComponentMessage2.style.display = "none";
                            OTSAddComponentMessage3.style.display = "none";
                        }
                        else {
                            OTSAddComponentMessage2.style.display = "";
                            OTSAddComponentMessage3.style.display = "";
                        }
                    }

                    window.scrollTo(0, 0);
                    break;

                case "StatusReport":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "none";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="none";

                    //			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";

                    window.scrollTo(0, 0);
                    break;
                case "Approvers":
                    steptext = "";

                    //			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
                    tabGeneral.style.display = "none";
                    tabSystemTeam.style.display = "none";
                    tabFiles.style.display = "none";
                    tabStatusReport.style.display = "none";
                    tabOTS.style.display = "none";
                    tabAccess.style.display = "none";
                    //tabApprovers.style.display="";

                    //			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";

                    window.scrollTo(0, 0);
                    break;

            }
        }
        function window_onload() {
            ProgramInput.txtPDDPath.value = ProgramInput.tagPDDPath.value;
            ProgramInput.txtSCMPath.value = ProgramInput.tagSCMPath.value;
            ProgramInput.txtSTLPath.value = ProgramInput.tagSTLPath.value;
            ProgramInput.txtProgramMatrixPath.value = ProgramInput.tagProgramMatrixPath.value;
            ProgramInput.txtAccessoryPath.value = ProgramInput.tagAccessoryPath.value;
            LoadingRow.style.display = "none";
            ButtonRow.style.display = "";
            if (txtDefaultTab.value == "SystemTeam")
                CurrentState = "SystemTeam";
            else if (txtDefaultTab.value == "OTS")
                CurrentState = "OTS";
            else if (txtDefaultTab.value == "FilePaths")
                CurrentState = "Files";
            else if (txtDefaultTab.value == "StatusData")
                CurrentState = "StatusReport";
            else
                CurrentState = "General";
            //ProcessState();
            SelectTab(CurrentState);
            self.focus();

            BusinessSegment_onchange();

            $("#cboBusinessSegmentID").change(function () {
                BusinessSegment_onchange();
            });

            //initialize modal dialog
            modalDialog.load();
        }

        function SelectTab(strStep) {
            var i;

            //Reset all tabs
            document.all("CellGeneralb").style.display = "none";
            document.all("CellGeneral").style.display = "";
            document.all("CellSystemTeamb").style.display = "none";
            document.all("CellSystemTeam").style.display = "";
            document.all("CellAccessb").style.display = "none";
            document.all("CellAccess").style.display = "none";
            document.all("CellStatusReportb").style.display = "none";
            document.all("CellStatusReport").style.display = "";
            document.all("CellOTSb").style.display = "none";
            if (ProgramInput.txtID.value != "")
                document.all("CellOTS").style.display = "";
            document.all("CellFilesb").style.display = "none";
            document.all("CellFiles").style.display = "";
        

            //Highight the selected tab
            document.all("Cell" + strStep).style.display = "none";
            document.all("Cell" + strStep + "b").style.display = "";


            CurrentState = strStep;
            ProcessState();
        }



        function cmdAddFamily_onclick() {
            modalDialog.open({ dialogTitle: 'Add Product Family', dialogURL: 'family.asp', dialogHeight: 250, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
            /*var strID = new Array();
            strID = window.showModalDialog("family.asp", "", "dialogWidth:435px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
        }

        function cmdAddFamilyResults() {
            var strResult;

            strResult = modalDialog.getArgument('family_query_array');
            strResult = JSON.parse(strResult);

            if (typeof (strResult) != "undefined") {
                ProgramInput.cboFamily.options[ProgramInput.cboFamily.length] = new Option(strResult[1], strResult[0]);
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

        function cmdQualityAdd_onclick() {
            ChooseEmployee(ProgramInput.cboQuality);
        }

        function cmdODMSEPMAdd_onclick() {
            ChooseEmployee(ProgramInput.cboODMSEPM);
        }

        function cmdProcurementPMAdd_onclick() {
            ChooseEmployee(ProgramInput.cboProcurementPM);
        }

        function cmdPlanningPMAdd_onclick() {
            ChooseEmployee(ProgramInput.cboPlanningPM);
        }

        function cmdODMPIMPMAdd_onclick() {
            ChooseEmployee(ProgramInput.cboODMPIMPM);
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

        function cmdSysEngrProgramCoordinatorAdd_onclick() {
          ChooseEmployee(ProgramInput.cboSysEngrProgramCoordinator);
        }
        
        function SystemTeamAdd(objectId) {
            var obj = document.getElementById(objectId);
            ChooseEmployee(obj)
        }

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

        function BrandCheck_onclick(ID, PVID) {
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
                    Result = window.showModalDialog("BrandDeleteWarning.asp?ProductName=" + ProgramInput.txtProductFamily.value + " " + ProgramInput.txtVersion.value + "&BrandName=" + event.srcElement.BrandName + "&BrandID=" + ID + "&PVID=" + PVID, "", "dialogWidth:700px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
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
            if (ProgramInput.cboDevCenter.value != 2 && ProgramInput.cboDevCenter.value != 6) {
                lblPOPM.innerHTML = "Configuration&nbsp;Manager:"
                lblTDCCM.innerHTML = "Program&nbsp;Office&nbsp;Manager:"
            }
            else {
                lblPOPM.innerHTML = "Program&nbsp;Office&nbsp;Manager:"
                lblTDCCM.innerHTML = "Configuration&nbsp;Manager:"
            }
            if (ProgramInput.cboReleaseTeam.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
                ProgramInput.cboReleaseTeam.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;
            if (ProgramInput.cboPreinstall.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
                ProgramInput.cboPreinstall.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;

            if (ProgramInput.txtID.value == "") ChangeDefaultDCROwner_onDevCenterChange();

        }

        function ChangeDefaultDCROwner_onDevCenterChange(){
            if (window.parent.frames["UpperWindow"].lblPOPM.innerHTML.indexOf("Configuration") > -1) {
                ProgramInput.cboDCRDefaultOwner.selectedIndex = 0; //Configuration Maager
            } else {
                ProgramInput.cboDCRDefaultOwner.selectedIndex = 1; //Program Office Manager
            }
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

        /*Modified By:	02/22/2016 JMalichi - Task 16730: Update all products assigned when generic component core team changes*/
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
                        if (ResultArray[0] == "") {
                            SystemboardLink.innerText = "Add ID";
                        } else {
                            SystemboardLink.innerText = ResultArray[0].replace(/,/g, ", ");
                        }
                        ProgramInput.txtSystemBoardComments.value = ResultArray[1];
                    }
                    break;
                case "2":
                    if (typeof (ResultArray) != "undefined") {
                        //if (ResultArray[0] != 0)
                        ProgramInput.txtMachinePNPID.value = ResultArray[0];
                        if (ResultArray[0] == "") {
                            MachinePNPLink.innerText = "Add ID";
                        } else {
                            MachinePNPLink.innerText = ResultArray[0].replace(/,/g, ", ");
                        }
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

            modalDialog.open({ dialogTitle: 'Brand', dialogURL: 'BrandUpdate.asp?ProductID='+ ProgramInput.txtID.value + '&BrandID=' + strID + '&ExcludeIDList=' + strIDs + '', dialogHeight: 300, dialogWidth: 450, dialogResizable: false, dialogDraggable: true });
            globalVariable.save(strID, 'brand_id');
        }

        function ChoosenewBrandResult() {
            var strID;
            var strResult;

            strResult = modalDialog.getArgument('brand_update_result');
            strID = globalVariable.get('brand_id');

            if (typeof (strResult) != "undefined") {
                for (i = 0; i < ProgramInput.chkBrands.length; i++) {
                    if (ProgramInput.chkBrands[i].value == strID) {
                        ProgramInput.chkBrands[i].checked = false;
                        document.all("Brand" + strID).style.display = "none";
                        document.all("DivSeries" + strID).style.display = "none";
                    }

                    if (ProgramInput.chkBrands[i].value == strResult) {
                        ProgramInput.chkBrands[i].checked = true;
                        document.all("DivSeries" + strResult).style.display = "";
                    }
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
            modalDialog.open({ dialogTitle: 'Product Group', dialogURL: 'ProductCycle.asp?ID=' + ID + '', dialogHeight: 550, dialogWidth: 705, dialogResizable: false, dialogDraggable: true });
            globalVariable.save(ID, 'product_group_id');
        }

        function SelectedCyclesResult(strResult) {
            var ID = globalVariable.get('product_group_id');
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

        function BusinessSegment_onchange() {
            var plDropdown = $("#cboProductLine");
            var brandTableRows = $("#TableBrand tbody tr");
            var businessSegmentId = $("#cboBusinessSegmentID").val();
            var currentProductLine = $("#cboProductLine").val();

            if (businessSegmentId > 0) {
                
                $.ajax({
                    url: "/Pulsar/Product/GetProductLines?BusinessSegmentId=0", // + businessSegmentId,
                    method: "GET",
                    success: function (returnData) {
                        var data = jQuery.parseJSON(returnData);
                        $(plDropdown).html('');
                        $(plDropdown).append('<option value=""></option>');

                        $(document).ready(function () {
                            for (var i = 0; i <= data.productLines.length - 1; i++) {
                                var item = data.productLines[i];
                                if (item.ProductLineId == currentProductLine) {
                                    $(plDropdown).append('<option value="' + item.ProductLineId + '" selected="selected">' + item.ShortName +
                                        ' - ' + item.Description + '</option>');
                                } else {
                                    $(plDropdown).append('<option value="' + item.ProductLineId + '">' + item.ShortName +
                                        ' - ' + item.Description + '</option>');
                                }
                            }
                        });
                    },
                    cache: false
                });

                $.ajax({
                    url: "/Pulsar/Product/GetBrands?BusinessSegmentId=" + businessSegmentId,
                    method: "GET",
                    success: function (returnData) {
                        var data = jQuery.parseJSON(returnData);
                        var rows = $(brandTableRows).each(function (rowNum, val) {
                            var row = $(this);
                            var itemID;
                            var rowID = row[0].id;
                            var bdisplay;
                            rowID = rowID + ",";
                            bdisplay = false;
                            for (var i = 0; i <= data.brands.length - 1; i++) {

                                itemID = "Brand" + data.brands[i].ID + ",";
                                if (rowID.indexOf(itemID) > -1) {
                                    bdisplay = true;
                                }
                            }
                            if (bdisplay) {
                                $(row).show();
                            } else {
                                $(row).hide();
                            }

                        });
                    },
                    cache: false
                });

            }
        }

        function cmdDate_onclick(strField, strEditServiceEOLDate) {
            var strID;
            var i;

            if (strEditServiceEOLDate == "disabled") {
                return;
            }

            strID = window.showModalDialog("../../Mobilese/today/calDraw1.asp", window.document.all(strField).value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                window.document.all(strField).value = strID;
            }
        }
//-->
    </script>

</head>
<style>
    .OTSComponentCell
    {
        border-top: solid 1px #b2b2b2;
        font-size: xx-small;
        vertical-align: middle;
        line-height: 15px;
        font-family: Verdana;
    }
</style>
<link rel="stylesheet" type="text/css" href="../../style/wizard%20style.css">
<body onload="return window_onload()">
    <font face="verdana">
        <%
	dim cn
	dim rs
	dim cm
	dim p
	dim CnString
	dim strFamily
    'LY BEGINNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
	dim strProductLine
    dim strProductLines
	'LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
	

	dim strSEPM
	dim strPM
	dim strTDCCM	
	dim strSM
	dim strVersion
    dim strProductRelease
	dim strDistribution
	dim strCvrBuildDist
	dim strCvrReleaseDist
	dim strActionNotifyList
	dim strType
	dim strApprover
	dim strDescription
	dim strObjective
	dim strOTSName
	dim strEmailActive
    ' LY BEGINNNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
	dim strProductLineID
	' LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
	dim strFamilyID
	dim CheckEmail
	dim CheckReports
	dim strDivision
	dim strProductStatus
	dim strBaseUnit
	dim strCurrentROM
	dim strCurrentWebROM
	dim strOSSupport
	dim strImagePO
	dim strImageChanges
	dim strSystemboardID
	dim strSystemboardComments
	dim strMachinePNPID
	dim strMachinePNPComments
	dim strCommonimages
	dim strCertificationStatus
	dim strSWQAStatus
	dim strPlatformStatus
    dim strBusinessSegmentID
'	dim strBrands
	dim strDCRAutoOpen
    dim strDCRNoOdm
	dim strDcrToCM
	dim strDcrToList
	strDCRAutoOpen = ""
    strDCRNoOdm = ""
	strDcrToCm = ""
	strDcrToList = ""
	dim strPOPMLabel
	dim strServiceLifeDate
	dim strServicePMNotifications
	dim strPDE
	dim strFactoryEngineer
	dim strAPM
	dim strPreinstall
	dim strReleaseTeam
	dim strComMarketing
	dim strConsMarketing
	dim strSMBMarketing
	dim strFinanceList
    dim strQualityList
	dim strServiceList
	dim strSupplyChainList
	dim PlatformEngineeringList
	'dim strAvailableBrands
	dim strRTM1
	dim strRTM2
	dim strRTM3
	dim strSystemTeam
	dim strPartner
	dim strBIOSLead
	dim strVideoMemoryPM
	dim strGraphicsControllerPM
	dim strProcessorPM
	dim strCommHWPM
	dim strDKC
	dim CheckCommodities
	dim CheckSMR
    dim CheckFusion
	dim CheckDeliverables
	dim CheckImages
	dim CheckDCR
	dim EnableDCR
	dim CheckMDA
	dim strChecked
	dim strBrandsLoaded
	dim strReleasesLoaded
	dim strPDDPath
	dim strSCMPath
	dim strAccessoryPath
	dim strSTLPath
	dim strProgramMatrixPath
	dim strReferenceList
	dim strReferenceID
	dim strDevCenter
	dim strRegulatoryModel
	dim strServiceTag
	dim strBIOSBranding
	dim strPCID
	dim strPCList
	dim strMarketingOpsID
	Dim strMarketingOpsList
	dim strBIOSLeadList
	dim strVideoMemoryPMList
	dim strGraphicsControllerPMList
	dim strFactoryEngineerList
	dim strProcessorPMList
	dim strCommHWPMList
	dim strDKCList
	dim strDCRApproverList
	dim DisplayToolsProject
	dim DisplayToolsProject2
	dim strToolAccessIDList
    dim strGplmList
	dim strSpdmList
  ' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
  dim strSysEngrProgramCoordinatorList
  ' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
	dim strSbmList
	dim strGplm
	dim strSpdm
	dim strSba
	dim strRCTOSites
    dim strDefaultDCROwner
    dim strMinRoHSLevel
    dim CheckBSAMFlag
    dim iAffectedProduct
    dim strNever
	dim strAlways
	dim strUntil
	dim strcboMilestones
	dim strReqMilestones
	dim strIsSEPM
	dim AddDCRNotificationList
    dim strWWAN : strWWAN = ""
    dim strHWStatusDisplay : strHWStatusDisplay = ""
	dim strHWStatus : strHWStatus = "" 
	dim strEditServiceEOLDate : strEditServiceEOLDate = "disabled"

	strReqMilestones = ""
    
    iAffectedProduct = 0
    
    strNever = ""
	strAlways = ""
	strUntil = ""
	strcboMilestones = ""
	strIsSEPM = "False"
	
	strServiceLifeDate = ""
	strServicePMNotifications = ""
	strToolAccessIDList = ""
	
	strRCTOSites=""
	
	strBrandsLoaded = ""
	strReleasesLoaded = ""
	
	cnString =Session("PDPIMS_ConnectionString")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


' get current user info. Using user's phone number. TDC user's phone number should not start with +1 or 01., the user is not from TDC
' If current user is TDC, display both Program office PM and configurationPM
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPhone
	dim CurrentUserID
	
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
	Set	rs = cm.Execute 
	
	set cm=nothing	
		
	if not (rs.EOF and rs.BOF) then
		CurrentUserPhone = rs("Phone") & ""
		CurrentUserID = rs("ID") & ""
	end if
	rs.Close

	if request("ID") = "" then
		Response.write "<H3>Add New Product (Legacy)</H3>" 
        strDefaultDCROwner = "" 
		strPM = ""
		strSM = ""
		strTDCCM = ""		
		strSEPM = ""
		strPDE = ""
		strAPM = ""
		strRTM1 = ""
		strRTM2 = ""
		strRTM3 = ""
		strComMarketing = ""
		strConsMarketing = ""
		strSMBMarketing = ""
		strPreinstall = ""
		strReleaseTeam = ""
		strDistribution = ""
		strCvrBuildDist = ""
		strCvrReleaseDist = ""
		strActionNotifyList = ""
		strDCRApproverList = ""
		strVersion = ""
        strProductRelease = ""
		strDescription = ""
		strObjective = ""
		strOTSName = ""
		strApprover = ""
		strEmailActive = ""
    ' LY BEGINNNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
		strProductLineID = ""
  	' LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
		strFamilyID = ""
		CheckEmail = "checked"
		CheckReports = "checked"
		strDivision = ""
        strBusinessSegmentID = 0
		strType = ""
	    strRCTOSites = ""
		strProductStatus = "1"
		strBaseUnit = ""
		strCurrentROM = ""
		strCurrentWebROM = ""
		strOSSupport = ""
		strImagePO = ""
		strImageChanges = ""
		strSystemboardID = ""
		strSystemboardComments= ""
		strMachinePNPID = ""
		strMachinePNPComments = ""
		strCommonimages = ""
		strCertificationStatus = ""
		strSWQAStatus = ""
		strPlatformStatus = ""
		'strBrands = ""
		strWWANTestLeadList = ""
		strDCRAutoOpen = "checked"
        strDCRNoOdm = ""
		strDcrToCm = ""
		strDcrToList = ""
		strPDDPath = ""
		strSCMPath = ""
		strAccessoryPath = ""
		strSTLPath = ""
		strProgramMatrixPath = ""
		strPlatformDevelopment = ""
		strSupplyChain = ""
		strFinance = ""
        strQuality = ""
		strService = ""
		strPartner = ""
		strReferenceID = ""
		strDevCenter = ""
		strDocPM = ""
        strODMSEPM = ""
        strProcurementPM = ""
        strPlanningPM = ""
        strODMPIMPM = ""
		strSEPE = ""
		strPINPM = ""
		strSETestLead = ""
		strSETest = ""
		strODMTestLead = ""
		strWWANTestLead = ""
		strBIOSLead = ""
		strCommHWPM = ""
		strDKC = ""
		strVideoMemoryPM = ""
		strGraphicsControllerPM = ""
		strFactoryEngineer = ""
		strType =""
		strProcessorPM = ""
		strSustainingMgrID = ""
		strSustainingSEPMID = ""
    ' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
    strSysEngrProgramCoordinatorID = ""
    ' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
		strPinCutoffValue = ""
		strRegulatoryModel = ""
		strServiceTag = ""
	    strBIOSBranding = ""
		strPOPMLabel = "Configuration&nbsp;Manager:"
		strPCID = ""
		strMarketingOpsID = ""
		DisplayToolsProject = ""
		DisplayToolsProject2= "none"
    	strGplm = ""
    	strSpdm = ""
    	strSba = ""
    	strMinRoHSLevel = "0"
    	CheckBSAMFlag = "" 'This is a bit field in the ProductVersion table as of 3-27-2013 I'm assuming this needs to be set true for all newly created products.
		AddDCRNotificationList = "checked"
		CheckCommodities = "checked"
		CheckSMR = ""
        CheckFusion = "checked"
		CheckDeliverables = "checked"
		CheckImages = "checked"
		CheckDCR = ""
		EnableDCR=" disabled "
		
		CheckMDA = "checked"
		'Response.Write "<font color=red size=1 face=verdana><b>Please contact Dave Whorton before adding products with PRS, PAV, ENT, TAB, WKS, or SMB in the Version field.</b></font>"
	else
		Response.write "<H3>Product Properties (Legacy)</H3>" 
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spGetProductVersion " & request("ID"),cn,adOpenForwardOnly
		strPM = rs("PMID") & ""
		strSM = rs("SMID") & ""
		strTDCCM = rs("TDCCMID") & ""		
		strSEPM = rs("SEPMID") & ""
		strPDE = rs("PDEID") & ""
		strRCTOSites = rs("RCTOSites") & ""
		strAPM = rs("AccessoryPMID") & ""
		strRTM1 = rs("RolloutP1") & ""
		strRTM2 = rs("RolloutP2") & ""
		strRTM3 = rs("RolloutP3") & ""
        strDefaultDCROwner = rs("DCRDefaultOwner") & ""
		strComMarketing = rs("ComMarketingID") & ""
		strSMBMarketing = rs("SMBMarketingID") & ""
		strConsMarketing = rs("ConsMarketingID") & ""
		strPlatformDevelopment = rs("PlatformDevelopmentID") & ""
		strSupplyChain = rs("SupplyChainID") & ""
		strQuality = rs("QualityID") & ""
		strService = rs("ServiceID") & ""
		strDistribution = rs("Distribution") & ""
		strCvrBuildDist = rs("ConveyorBuildDistribution") & ""
		strCvrReleaseDist = rs("ConveyorReleaseDistribution") & ""
		strActionNotify = rs("ActionNotifyList") & ""
		strReferenceID = rs("ReferenceID") & ""
		strDevCenter = rs("DevCenter") & ""
		strVersion = rs("Version") & ""
		strProductRelease = rs("ProductRelease") & ""
		strDescription = rs("Description") & ""
		strObjective = rs("Objectives") & ""
		strOTSName = rs("DotsName") & ""
		strPreinstall = rs("PreinstallTeam") & ""
		strReleaseTeam = rs("ReleaseTeam") & ""
		strApprover = rs("Approver") & ""
		strEmailActive = rs("EmailActive") & ""
		select case rs("DCRAutoOpen") & ""
		    case "1"
		        strDcrToCm = "checked"
		    case "2"
		        strDCRAutoOpen = "checked"
		    case "3"
		        strDcrToList = "checked"
            case "4"
                strDCRNoOdm = "checked"
		    case else
		        strDcrToCm = "checked"
		end select
		
    ' LY BEGINNNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
		strProductLineID = rs("ProductLineID") & ""
  	' LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
		strFamilyID = rs("ProductFamilyID") & ""
		'strBrands = rs("Brands") & ""
		strPartner = rs("PartnerID") & ""
		strPDDPath = rs("PDDPath") & ""
		strSCMPath = rs("SCMPath") & ""
		strAccessoryPath = rs("AccessoryPath") & ""
		strSTLPath = rs("STLStatusPath") & ""
		strProgramMatrixPath = rs("ProgramMatrixPath") & ""
		strDocPM = rs("DocPM") & ""
        strODMSEPM = rs("ODMSEPMID") & ""
        strProcurementPM = rs("ProcurementPMID") & ""
        strPlanningPM = rs("PlanningPMID") & ""
        strODMPIMPM = rs("ODMPIMPMID") & ""
		strSEPE = rs("SEPE") & ""
		strPINPM = rs("PINPM") & ""
		strBIOSLead = rs("BiosLeadID") & ""
		strDKC = rs("DKCID") & ""
		strCommHWPM = rs("CommHWPMID") & ""
		strVideoMemoryPM = rs("VideoMemoryPMID") & ""
		strGraphicsControllerPM = rs("GraphicsControllerPMID") & ""
		strFactoryEngineer = rs("SCFactoryEngineerID") & ""
		strProcessorPM = rs("ProcessorPMID") & ""
        strMinRoHSLevel = rs("MinRoHSLevel")
        if rs("BSAMFlag") then
            CheckBSAMFlag = "checked"
        else
            CheckBSAMFlag = ""
        end if
        
		strSustainingMgrID = rs("SustainingMgrID")
		strSustainingSEPMID = rs("SustainingSEPMID") & ""
		
    ' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
    strSysEngrProgramCoordinatorID = rs("SysEngrProgramCoordinatorID") & ""
    ' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
    
		strSETestLead = rs("SETestLead") & ""
		strSETest = rs("SETestID") & ""
		strODMTestLead = rs("ODMTestLeadID") & ""
		strWWANTestLead = rs("WWANTestLeadID") & ""
		strPinCutoffValue = rs("PreinstallCutoff") & ""
		strRegulatoryModel = trim(rs("RegulatoryModel") & "")
		strPCID = rs("PCID") & ""
		strMarketingOpsID = rs("MarketingOpsID") & ""
		strGplm = rs("Gplm") & ""
		strSpdm = rs("Spdm") & ""
		strSba = rs("SvcBomAnalyst")
		iAffectedProduct = rs("AffectedProduct")
        select case rs("AffectedProduct") & ""
		    case "-1"
		        strNever = "checked"
		        strcboMilestones = "disabled"
		    case "0"
		        strAlways = "checked"
		        strcboMilestones = "disabled"
		    case else
		        strUntil = "checked"
		        strcboMilestones = "enabled"
		end select
		
		if trim(strDevCenter) = "2" or trim(strDevCenter) = "6" then
			strPOPMLabel = "Program&nbsp;Office&nbsp;Manager:"
		else
			strPOPMLabel = "Configuration&nbsp;Manager:"
		end if
		'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if (strSEPM = CurrentUserID) then
		    strIsSEPM = "True"
		end if
		
		if rs("EmailActive") & "" = "" or rs("EmailActive") & "" = "0" then	
			CheckEmail = ""
		else
			CheckEmail = "checked"
		end if
		
		if rs("OnlineReports") & "" = ""  or rs("OnlineReports") & "" = "0" then	
			CheckReports = ""
		else
			CheckReports = "checked"
		end if
		if rs("OnCommodityMatrix") then	
			CheckCommodities = "checked"
		else
			CheckCommodities = ""
		end if
        if rs("Fusion") then
            CheckFusion = "checked"
        else
            CheckFusion = ""
        end if
		if rs("AllowSMR") then	
			CheckSMR = "checked"
		else
			CheckSMR = ""
		end if
		if rs("AllowDeliverableReleases") then	
			CheckDeliverables = "checked"
		else
			CheckDeliverables = ""
		end if
		if rs("AllowImageBuilds") then	
			CheckImages = "checked"
		else
			CheckImages = ""
		end if
		if rs("AllowDCR") then	
			CheckDCR = "checked"
		else
			CheckDCR = ""
		end if
		if rs("ShowOnWhql") then
		    CheckMDA = "checked"
		else
		    CheckMDA = ""
		end if
		strDivision = rs("Division") & ""
        strBusinessSegmentID = rs("BusinessSegmentID") & ""
		strType = rs("TypeID") & ""
		if trim(strType) = "2" then
			DisplayToolsProject = "none"
			DisplayToolsProject2 = ""
		else
			DisplayToolsProject = ""
			DisplayToolsProject2 = "none"
		end if
		if rs("AddDCRNotificationList") then
		    AddDCRNotificationList = "checked"
		else
		    AddDCRNotificationList = ""
		end if
	
		strServiceLifeDate = trim(rs("ServiceLifeDate") & "")
	

		strProductStatus = rs("ProductStatusID")	
		
		if trim(rs("ProductStatusID")) = "1" then
			EnableDCR = " disabled "
		else
			EnableDCR = ""
		end if

        if trim(rs("ProductStatusID")) = 4 then
            strEditServiceEOLDate = ""
        end if
                 
		strBaseUnit = rs("BaseUnit") & "" 
		strCurrentROM = rs("CurrentROM") & ""
		strCurrentWebROM = rs("CurrentWebROM") & ""
		strOSSupport = rs("OSSupport") & ""
		strImagePO = rs("ImagePO") & ""
		strImageChanges = rs("ImageChanges") & ""
		strSystemboardID = rs("SystemboardID") & ""
		strSystemboardComments = rs("SystemboardComments") & ""
		strMachinePNPID = rs("MachinePNPID") & ""
		strMachinePNPComments = rs("MachinePNPComments") & ""
		strCommonimages = rs("Commonimages") & ""
		strCertificationStatus = rs("CertificationStatus") & ""
		strSWQAStatus = rs("SWQAStatus") & ""
		strPlatformStatus = rs("PlatformStatus") & ""
		strDCRApproverList = rs("DCRApproverList") & ""
		strToolAccessIDList = rs("ToolAccessList") & ""
        
        'strServiceTag = rs("ServiceTag") & ""
	    'strBIOSBranding = rs("BIOSBranding") & ""
	    
        if rs("WWANProduct") & "" = "" then	
			strWWAN = ""
		else
			strWWAN = abs(clng(rs("WWANProduct")))
		end if

        if rs("CommodityLock") then
			strHWStatusDisplay = "none"
			strHWStatus = "checked"
		else
			strHWStatusDisplay = ""
			strHWStatus = ""
		end if
        
		rs.Close
	end if		


	TmpArray = split(strDistribution,";")
	strDistribution = ""
	for each strTemp in TmpArray 
		if trim(strTemp) <> "" then strDistribution = strDistribution & "; " & trim(strTemp)
	next
	if strDistribution <> "" then strDistribution = mid(strDistribution,3)

    TmpArray = split(strCvrBuildDist,";")
    strCvrBuildDist = ""
    For Each strTemp in TmpArray
        If Trim(strTemp) <> "" Then strCvrBuildDist = strCvrBuildDist & "; " & Trim(strTemp)
    Next
    If strCvrBuildDist <> "" Then strCvrBuildDist = mid(strCvrBuildDist,3)
  '  If strCvrBuildDist = "" Then strCvrBuildDist = "houportpreinpm@hp.com"

    TmpArray = split(strCvrReleaseDist,";")
    strCvrReleaseDist = ""
    For Each strTemp in TmpArray
        If Trim(strTemp) <> "" Then strCvrReleaseDist = strCvrReleaseDist & "; " & Trim(strTemp)
    Next
    If strCvrReleaseDist <> "" Then strCvrReleaseDist = mid(strCvrReleaseDist,3)
   ' If strCvrReleaseDist = "" Then strCvrReleaseDist = "houportpreinpm@hp.com"
    
	TmpArray = split(strActionNotify,";")
	strActionNotify = ""
	for each strTemp in TmpArray 
		if trim(strTemp) <> "" then strActionNotify = strActionNotify & "; " & trim(strTemp)
	next
	if strActionNotify <> "" then strActionNotify = mid(strActionNotify,3)


    if request("ID") <> "" then
        rs.Open "usp_SelectScheduleReqMilestoneData " & request("ID"),cn,adOpenForwardOnly
        strReqMilestones = ""
        dim iScheduleID
        iScheduleID = 0
	    do while not rs.EOF
		    if strReqMilestones = "" then
		        strReqMilestones = strReqMilestones & "<optgroup label=""" & rs("schedule_name") & """>"
		        iScheduleID = rs("schedule_id")
		    elseif rs("schedule_id") <> iScheduleID then
		        iScheduleID = rs("schedule_id")
		        strReqMilestones = strReqMilestones & "</optgroup><optgroup label=""" & rs("schedule_name") & """>"
		    end if
		    if rs("schedule_data_id") = iAffectedProduct then
			    strReqMilestones = strReqMilestones & "<OPTION selected value=""" & rs("schedule_data_id") & """>" & rs("item_description") & "</OPTION>" & vbcrlf
		    else
			    strReqMilestones = strReqMilestones & "<OPTION value=""" & rs("schedule_data_id") & """>" & rs("item_description") & "</OPTION>" & vbcrlf
		    end if
		    rs.MoveNext
	    loop
	    rs.Close
	    strReqMilestones = strReqMilestones & "</optgroup>"
    end if
    
	strFamily = ""
	rs.Open "spGetProductFamiliesAll",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(strFamilyID) = trim(rs("ID") & "" ) then
			strFamilies = strFamilies & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			strFamily = rs("Name")
		elseif rs("Active") = 1 then
			strFamilies = strFamilies & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close

    Response.Cookies("ProductFamily")= strFamily
    Response.Cookies("BusinessSegmentId") = strBusinessSegmentID

    ' LY BEGINNNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM
	strProductLine = ""
	rs.Open "spGetProductLinesAll",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(strProductLineID) = trim(rs("ID") & "" ) then
			strProductLines = strProductLines & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & " - " & rs("Description") & "</OPTION>"
			strProductLine = rs("Name")
		' elseif rs("Active") = 1 then
		else
			strProductLines = strProductLines & "<OPTION value=" & rs("ID") & ">" & rs("Name") & " - " & rs("Description") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close
  ' LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM

	strReferenceList = ""
	strInactiveList = ""
	rs.Open "spGetProductsAll 1",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(request("ID")) <> trim(rs("ID") & "" ) and rs("FusionRequirements") = 0 then 'Exclude current version
			if trim(strReferenceID) = trim(rs("ID") & "" ) then
				strReferenceList = strReferenceList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & " " & rs("Version")  & "</OPTION>"
			elseif rs("ProductStatusID") < 5 then
				strReferenceList = strReferenceList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & " " & rs("Version") & "</OPTION>"
			else
			    strInactiveList = strInactiveList & "<option>" & rs("Name") & " " & rs("Version") & "</option>"
			end if
		end if
		rs.MoveNext
	loop
	rs.Close

    strDocPMList = ""
	strSEPEList = ""
	strAPMList = ""
	strPINPMList = ""
	strSETestLeadList = ""
	strSETestList = ""
	strODMTestLeadList = ""
	strSustainingMgrList = ""
	strSustainingSEPMList = ""
  ' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
  strSysEngrProgramCoordinatorList = ""
  ' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
	strSEPMList = ""
	strTDCCMList = ""
	strPDEList = ""
	strFactoryEngineerList = ""
	strPMList = ""
	strSMList = ""
	strToolsPMList = ""
	strComMarketingList = ""
	strConsMarketingList = ""
	strSMBMarketingList = ""
	strApproverList = ""
	strFinanceList= ""
    strQualityList= ""
	strServiceList= ""
	strSupplyChainList= ""
	PlatformEngineeringList= ""
	strMarketingOpsList = ""
	strBIOSLeadList = ""
	strVideoMemoryPMList = ""
	strGraphicsControllerPMList = ""
	strFactoryEngineerList = ""
	strProcessorPMList = ""
	strCommHWPMList = ""
	strDKCList = ""
	strGplmList = ""
	strSpdmList = ""
	strSbmList = ""

    strODMSEPMList = ""
    strProcurementPMList = ""
    strPlanningPMList = ""
    strODMPIMPMList = ""

	'rs.Open "spGetEmployees",cn,adOpenForwardOnly
	if trim(request("ID")) = "" then
		rs.Open "spListSystemTeamDropdowns 0",cn,adOpenStatic
	else
		rs.Open "spListSystemTeamDropdowns " & clng( request("ID")),cn,adOpenStatic
	end if
	do while not rs.EOF
		if rs("Role") = "SE PM" then
			if trim(strSEPM) = trim(rs("ID")) then
				strSEPMList = strSEPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSEPMList = strSEPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "POPM" then
			if trim(strPM) = trim(rs("ID")) then
				strPMList = strPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strPMList = strPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "TDCCM" then
			if trim(strTDCCM) = trim(rs("ID")) then
				strTDCCMList = strTDCCMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strTDCCMList = strTDCCMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Commodity PM" then
			if trim(strPDE) = trim(rs("ID")) then
				strPDEList = strPDEList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			elseif trim(strPartner) = "" or trim(strPartner)="0" or trim(rs("PartnerID"))  = "1" or trim(rs("PartnerID")) = trim(strPartner) then
				strPDEList = strPDEList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		
		if rs("Role") = "Factory Engineer" then
			if trim(strFactoryEngineer) = trim(rs("ID")) then
				strFactoryEngineerList = strFactoryEngineerList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strFactoryEngineerList = strFactoryEngineerList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Accessory PM" then
			if trim(strAPM) = trim(rs("ID")) then
				strAPMList = strAPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else 
				strAPMList = strAPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if


		if rs("Role") = "Commercial Marketing" then
			if trim(strComMarketing) = trim(rs("ID")) then
				strComMarketingList = strComMarketingList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strComMarketingList = strComMarketingList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Consumer Marketing" then
			if trim(strConsMarketing) = trim(rs("ID")) then
				strConsMarketingList = strConsMarketingList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strConsMarketingList = strConsMarketingList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "SMB Marketing" then
			if trim(strSMBMarketing) = trim(rs("ID")) then
				strSMBMarketingList = strSMBMarketingList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSMBMarketingList = strSMBMarketingList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Platform Development" then
			if trim(strPlatformDevelopment) = trim(rs("ID")) then
				strPlatformDevelopmentList = strPlatformDevelopmentList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strPlatformDevelopmentList = strPlatformDevelopmentList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Supply Chain" then
			if trim(strSupplyChain) = trim(rs("ID")) then
				strSupplyChainList = strSupplyChainList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSupplyChainList = strSupplyChainList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Service" then
			if trim(strService) = trim(rs("ID")) then
				strServiceList = strServiceList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strServiceList = strServiceList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Quality" then
			if trim(strQuality) = trim(rs("ID")) then
				strQualityList = strQualityList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strQualityList = strQualityList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
				
		if rs("Role") = "System Manager" then
			if trim(strSM) = trim(rs("ID")) then
				strSMList = strSMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSMList = strSMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Tools PM" then
			if trim(strSM) = trim(rs("ID")) then
				strToolsPMList = strToolsPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strToolsPMList = strToolsPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if trim(strApprover) = trim(rs("ID")) then
			strApproverList = strApproverList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		else
			strApproverList = strApproverList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		end if
		
		if rs("Role") = "SE PE" then
			if trim(strSEPE) = trim(rs("ID")) then
				strSEPEList = strSEPEList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSEPEList = strSEPEList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Preinstall PM" then
			if trim(strPINPM) = trim(rs("ID")) then
				strPINPMList = strPINPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strPINPMList = strPINPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "SE Test Lead" then
			if trim(strSETestLead) = trim(rs("ID")) then
				strSETestLeadList = strSETestLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			elseif trim(strPartner) = "" or trim(strPartner)="0" or trim(rs("PartnerID"))  = "1" or trim(rs("PartnerID")) = trim(strPartner) then
				strSETestLeadList = strSETestLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "SE Test" then
			if trim(strSETest) = trim(rs("ID")) then
				strSETestList = strSETestList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			elseif trim(strPartner) = "" or trim(strPartner)="0" or trim(rs("PartnerID"))  = "1" or trim(rs("PartnerID")) = trim(strPartner) then
				strSETestList = strSETestList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "ODM Test Lead" then
			if trim(strODMTestLead) = trim(rs("ID")) then
				strODMTestLeadList = strODMTestLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			elseif trim(strPartner) = "" or trim(strPartner)="0" or trim(rs("PartnerID"))  = "1" or trim(rs("PartnerID")) = trim(strPartner) then
				strODMTestLeadList = strODMTestLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "WWAN Test Lead" then
			if trim(strWWANTestLead) = trim(rs("ID")) then
				strWWANTestLeadList = strWWANTestLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			elseif trim(strPartner) = "" or trim(strPartner)="0" or trim(rs("PartnerID"))  = "1" or trim(rs("PartnerID")) = trim(strPartner) then
				strWWANTestLeadList = strWWANTestLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "BIOS Lead" then
			if trim(strBIOSLead) = trim(rs("ID")) then
				strBIOSLeadList = strBIOSLeadList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strBIOSLeadList = strBIOSLeadList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Chipset/Processor PM" then
			if trim(strProcessorPM) = trim(rs("ID")) then
				strProcessorPMList = strProcessorPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strProcessorPMList = strProcessorPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Video Memory PM" then
			if trim(strVideoMemoryPM) = trim(rs("ID")) then
				strVideoMemoryPMList = strVideoMemoryPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strVideoMemoryPMList = strVideoMemoryPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Graphics Controller PM" then
			if trim(strGraphicsControllerPM) = trim(rs("ID")) then
				strGraphicsControllerPMList = strGraphicsControllerPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strGraphicsControllerPMList = strGraphicsControllerPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Comm HW PM" then
			if trim(strCommHWPM) = trim(rs("ID")) then
				strCommHWPMList = strCommHWPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strCommHWPMList = strCommHWPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
		
		if rs("Role") = "Doc Kit Coordinator" then
			if trim(strDKC) = trim(rs("ID")) then
				strDKCList = strDKCList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strDKCList = strDKCList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if		
		
		if rs("Role") = "SE Test Lead" then
			if trim(strSustainingMgrID) = trim(rs("ID")) then
				strSustainingMgrList = strSustainingMgrList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSustainingMgrList = strSustainingMgrList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		if rs("Role") = "Sustaining SE PM" then
			if trim(strSustainingSEPMID) = trim(rs("ID")) then
				strSustainingSEPMList = strSustainingSEPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSustainingSEPMList = strSustainingSEPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

	    ' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
		if rs("Role") = "SysEngrProgramCoordinator" then
			if trim(strSysEngrProgramCoordinatorID) = trim(rs("ID")) then
				strSysEngrProgramCoordinatorList = strSysEngrProgramCoordinatorList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSysEngrProgramCoordinatorList = strSysEngrProgramCoordinatorList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

    	' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB

		if rs("Role") = "Program Coordinator" then
			if trim(strPCID) = trim(rs("ID")) then
				strPCList = strPCList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strPCList = strPCList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
				
		if rs("Role") = "PDM Team" then
			if trim(strMarketingOpsID) = trim(rs("ID")) Then
				strMarketingOpsList = strMarketingOpsList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strMarketingOpsList = strMarketingOpsList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			End If
		end if

		if rs("Role") = "GPLM" then
			if trim(strGplm) = trim(rs("ID")) Then
				strGplmList = strGplmList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strGplmList = strGplmList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			End If
		end if

		if rs("Role") = "SPDM" then
			if trim(strSpdm) = trim(rs("ID")) Then
				strSpdmList = strSpdmList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSpdmList = strSpdmList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			End If
		end if

		if rs("Role") = "SBA" then
			if trim(strSba) = trim(rs("ID")) Then
				strSbaList = strSbaList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strSbaList = strSbaList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			End If
		end if
		
		if rs("Role") = "Doc PM" then
			if trim(strDocPM) = trim(rs("ID")) then
				strDocPMList = strDocPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strDocPMList = strDocPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

        if rs("Role") = "ODMSEPM" then
			if trim(strODMSEPM) = trim(rs("ID")) then
				strODMSEPMList = strODMSEPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strODMSEPMList = strODMSEPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
        
        if rs("Role") = "ProcurementPM" then
			if trim(strProcurementPM) = trim(rs("ID")) then
				strProcurementPMList = strProcurementPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strProcurementPMList = strProcurementPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if
              
        if rs("Role") = "PlanningPM" then
			if trim(strPlanningPM) = trim(rs("ID")) then
				strPlanningPMList = strPlanningPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strPlanningPMList = strPlanningPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

        if rs("Role") = "ODMPIMPM" then
			if trim(strODMPIMPM) = trim(rs("ID")) then
				strODMPIMPMList = strODMPIMPMList & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				strODMPIMPMList = strODMPIMPMList & "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
		end if

		rs.MoveNext
	loop
	rs.Close

	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0
	
        %>
        <form action="ProgramSave.asp" method="post" name="ProgramInput">
        <table class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0">
            <tr id="LoadingRow">
                <td>
                    Loading. Please Wait...
                </td>
            </tr>
            <tr id="ButtonRow" style="display: none" bgcolor="<%=strTitleColor%>">
                <td id="CellGeneral" style="display: none" width="10">
                    <font size="2" color="black"><b>&nbsp;<a href="javascript:SelectTab('General')">General</a>&nbsp;</b></font>
                </td>
                <td id="CellGeneralb" style="display: " width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;General&nbsp;</b></font>
                </td>
                <td id="CellSystemTeam" style="display: " width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('SystemTeam')">System&nbsp;Team</a>&nbsp;</b></font>
                </td>
                <td id="CellSystemTeamb" style="display: none" width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;System&nbsp;Team&nbsp;</b></font>
                </td>
                <td id="CellAccess" style="display: " width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Access')">Access&nbsp;List</a>&nbsp;</b></font>
                </td>
                <td id="CellAccessb" style="display: none" width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;Access&nbsp;List&nbsp;</b></font>
                </td>
                <%if request("ID") = "" then%>
                <td id="CellOTS" style="display: none" width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('OTS')">&nbsp;OTS&nbsp;</a>&nbsp;</b></font>
                </td>
                <%else%>
                <td id="CellOTS" style="display: " width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('OTS')">&nbsp;OTS&nbsp;</a>&nbsp;</b></font>
                </td>
                <%end if%>
                <td id="CellOTSb" style="display: none" width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;OTS&nbsp;</b></font>
                </td>
                <td id="CellStatusReport" style="display: " width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('StatusReport')">Status&nbsp;Data</a>&nbsp;</b></font>
                </td>
                <td id="CellStatusReportb" style="display: none" width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;Status&nbsp;Data</b></font>
                </td>
                <td id="CellFiles" style="display: " width="10">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Files')">File&nbsp;Paths</a>&nbsp;</b></font>
                </td>
                <td id="CellFilesb" style="display: none" width="10" bgcolor="wheat">
                    <font size="2" color="black"><b>&nbsp;File&nbsp;Paths</b></font>
                </td>
                <!--<td id="CellApprovers" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Approvers')">DCR&nbsp;Options</a>&nbsp;</b></font></td>
		<td id="CellApproversb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;DCR&nbsp;Options</b></font></td>-->
            </tr>
        </table>
        <font size="1" face="verdana">
            <br>
        </font>
        <table id="tabGeneral" style="display: none; width:990px; overflow:scroll; border-collapse:collapse;" border="1" cellpadding="2" cellspacing="0"
            bgcolor="cornsilk" bordercolor="tan">
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Family:</font></strong><font color="red" size="1"> *</font>
                </td>
                <td>
                    <%if request("ID") = "" then%>
                    <select id="cboFamily" name="cboFamily" style="width: 260px;" language="javascript"
                        onchange="return cboFamily_onchange()">
                        <option></option>
                        <%=strFamilies%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdAddFamily" name="cmdAddFamily"
                        language="javascript" onclick="return cmdAddFamily_onclick()">
                    <%else%>
                    <input type="text" id="txtViewFamily" style="width: 260px;" name="txtViewFamily"
                        disabled value="<%=strFamily%>"><select style="display: none" id="cboFamily" name="cboFamily"
                            style="width: 180px;">
                            <option></option>
                            <%=strFamilies%>
                        </select>&nbsp;<input style="display: none" type="button" value="Add" id="cmdAddFamily"
                            name="cmdAddFamily" language="javascript" onclick="return cmdAddFamily_onclick()">
                    <%end if%>
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Version:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td>
                    <input type="text" id="txtVersion" name="txtVersion" style="width: 200px;" value="<%=strVersion%>" maxlength="20">
                    <input type="hidden" id="tagVersion" name="tagVersion" style="width: 200px;" value="<%=strVersion%>" maxlength="20">
                    <input type="hidden" id="txtProductRelease" name="txtProductRelease" style="width: 200px;" value="<%=strProductRelease%>" maxlength="20">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">
                        <%if strDevCenter = "2" then%>
                        Reference&nbsp;Platform:
                        <%else%>
                        Lead&nbsp;Product:
                        <%end if%>
                    </font></strong>
                </td>
                <td>
                    <select id="cboReference" name="cboReference" style="width: 160px;">
                        <option value="0" selected></option>
                        <%=strReferenceList%>
                    </select>
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Product&nbsp;Type:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td>
                    <select id="cboType" name="cboType" style="width: 200px;" language="javascript" onchange="return cboType_onchange()">
                        <option selected value="0"></option>
                        <%if strType = "1" then%>
                        <option value="1" selected>Platform (Notebook,Tablet,etc.)</option>
                        <%else%>
                        <option value="1">Platform (Notebook,Tablet,etc.)</option>
                        <%end if%>
                        <%if strType = "3" then%>
                        <option value="3" selected>Dock/Port Rep./Jacket</option>
                        <%else%>
                        <option value="3">Dock/Port Rep./Jacket</option>
                        <%end if%>
                        <%if strType = "4" then%>
                        <option value="4" selected>Other</option>
                        <%else%>
                        <option value="4">Other</option>
                        <%end if%>
                    </select>
                    <select id="cboDivision" name="cboDivision" style="display: none; width: 160px;">
                        <option value="0"></option>
                        <%if strDivision = "1" or strDivision = "" or strDivision = "0" then%>
                        <option value="1" selected>Mobile</option>
                        <%else%>
                        <option value="1">Mobile</option>
                        <%end if%>
                        <%if strDivision = "2" then%>
                        <option selected value="2">bPC</option>
                        <%else%>
                        <option value="2">bPC</option>
                        <%end if%>
                        <%if strDivision = "3" then%>
                        <option selected value="3">cPC</option>
                        <%end if%>
                        <%if strDivision = "4" then%>
                        <option selected value="4">ISS</option>
                        <%else%>
                        <option value="4">ISS</option>
                        <%end if%>
                    </select>
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Partner:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td>
                    <select id="cboPartner" name="cboPartner" style="width: 160px;">
                        <option selected value="0"></option>
                        <%
				        rs.Open "spListPartners",cn,adOpenForwardOnly
				        do while not rs.EOF
					        if trim(strPartner) = trim(rs("ID") & "") then
						        Response.Write "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>"
					        elseif rs("active") then
						        Response.Write "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
					        end if
					        rs.Movenext
				        loop
			
				        rs.Close
                        %>
                    </select>
                </td>
                <td width="160" style="vertical-align: top"><strong><font size="2">Business&nbsp;Segment:</font></strong><font color="red" size="1">&nbsp;*</font></td>
				<td>
                    <select id="cboBusinessSegmentID" name="cboBusinessSegmentID" style="width: 200px;">
                         <option selected value=""></option>
                    <%  
                        rs.open "spPULSAR_Product_ListBusinessSegments",cn 
                        do while not rs.eof
                    %>
                       <%if trim(rs("BusinessSegmentID")) = trim(strBusinessSegmentID) then%>
                            <option selected value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                        <%else%>
                            <option value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                        <%end if%>
                    <%
                            rs.movenext
                        loop
                        rs.close    
                    %>
                    </select>
				</td>              
            </tr>
            <tr id="DevCenterRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Dev.&nbsp;Center:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td>
                    <select id="cboDevCenter" name="cboDevCenter" style="width: 160px;" language="javascript"
                        onchange="return cboDevCenter_onchange()">
                        <option value="0" selected></option>
                        <%if strDevCenter = "1" then%>
                        <option value="1" selected>Houston</option>
                        <%else%>
                        <option value="1">Houston</option>
                        <%end if%>
                        <%if strDevCenter = "2" then%>
                        <option value="2" selected>Taiwan - Consumer</option>
                        <%else%>
                        <option value="2">Taiwan - Consumer</option>
                        <%end if%>
                        <%if strDevCenter = "3" then%>
                        <option value="3" selected>Taiwan - Commercial</option>
                        <%else%>
                        <option value="3">Taiwan - Commercial</option>
                        <%end if%>
                        <%if strDevCenter = "4" then%>
                        <option value="4" selected>Singapore</option>
                        <%else%>
                        <option value="4">Singapore</option>
                        <%end if%>
                        <%if strDevCenter = "5" then%>
                        <option value="5" selected>Brazil</option>
                        <%else%>
                        <option value="5">Brazil</option>
                        <%end if%>
                        <%if strDevCenter = "6" then%>
                        <option value="6" selected>Mobility</option>
                        <%elseif false then%>
                        <option value="6">Mobility</option>
                        <%end if%>
                        <%if strDevCenter = "7" then%>
                        <option value="7" selected>San Diego</option>
                        <%else%>
                        <option value="7">San Diego</option>
                        <%end if%>
                        <%if strDevCenter = "8" then%>
                        <option value="8" selected>No Dev. Center</option>
                        <%else%>
                        <option value="8">No Dev. Center</option>
                        <%end if%>
                    </select>
                    <input type="hidden" id="tagDevCenter" name="tagDevCenter" value="<%=strDevCenter%>">
                </td>
                <td width="160" style="vertical-align: top"><strong><font size="2">Product&nbsp;Line:</font></strong><font color="red" size="1">&nbsp;*</font></td>
				<td>
                       <select id="cboProductLine" name="cboProductLine" style="width: 200px;" language="javascript">
                        <option></option>
                        <%=strProductLines%>
                    </select>&nbsp;
				</td>
            </tr>            
            <tr id="PreinstallRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Preinstall&nbsp;Team:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td>
                    <select id="cboPreinstall" name="cboPreinstall" style="width: 160px;">
                        <%if strPreinstall = ""  then%>
                        <option value="-1" selected></option>
                        <%else%>
                        <option value="-1"></option>
                        <%end if%>
                        <%if strPreinstall = "1" then%>
                        <option value="1" selected>Houston</option>
                        <%else%>
                        <option value="1">Houston</option>
                        <%end if%>
                        <%if strPreinstall = "2" then%>
                        <option selected value="2">Taiwan</option>
                        <%else%>
                        <option value="2">Taiwan</option>
                        <%end if%>
                        <%if strPreinstall = "3" then%>
                        <option selected value="3">Singapore</option>
                        <%else%>
                        <option value="3">Singapore</option>
                        <%end if%>
                        <%if strPreinstall = "4" then%>
                        <option value="4" selected>Brazil</option>
                        <%else%>
                        <option value="4">Brazil</option>
                        <%end if%>
                        <%if strPreinstall = "5" then%>
                        <option value="5" selected>CDC</option>
                        <%else%>
                        <option value="5">CDC</option>
                        <%end if%>
                        <%if strPreinstall = "6" then%>
                        <option value="6" selected>Houston - Thin Client</option>
                        <%else%>
                        <option value="6">Houston - Thin Client</option>
                        <%end if%>
                        <%if strPreinstall = "7" then%>
                        <option value="7" selected>Mobility</option>
                        <%else%>
                        <option value="7">Mobility</option>
                        <%end if%>
                        <%if strPreinstall = "0"  then%>
                        <option value="0" selected>No Image Changes</option>
                        <%else%>
                        <option value="0">No Image Changes</option>
                        <%end if%>
                    </select>
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">System&nbsp;Board&nbsp;ID:</font></strong>
                </td>
                <td>
                    <%if strSystemBoardID = "" then
			Response.write "<a ID=SystemboardLink href=""javascript:SelectSystemBoardID();"">Add ID</a>"
		  else
			Response.write	"<a ID=SystemboardLink  href=""javascript:SelectSystemBoardID();"">" & replace(strSystemBoardID,",",", ") & "</a>"
		  end if
                    %>
                    <input type="hidden" id="txtSystemBoardID" name="txtSystemBoardID" style="width: 160px;"
                        value="<%=strSystemBoardID%>">
                    <input type="hidden" id="txtSystemBoardComments" name="txtSystemBoardComments" style="width: 160px;"
                        value="<%=server.htmlencode(replace(strSystemBoardcomments,chr(161)&chr(168),"''"))%>">
                </td>
            </tr>
            <tr id="ReleaseRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Release&nbsp;Team:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td>
                    <select id="cboReleaseTeam" name="cboReleaseTeam" style="width: 160px;">
                        <%if strReleaseTeam = ""  then%>
                        <option value="-1" selected></option>
                        <%else%>
                        <option value="-1"></option>
                        <%end if%>
                        <%if strReleaseTeam = "1" then%>
                        <option value="1" selected>Houston</option>
                        <%else%>
                        <option value="1">Houston</option>
                        <%end if%>
                        <%if strReleaseTeam = "2" then%>
                        <option selected value="2">Taiwan</option>
                        <%else%>
                        <option value="2">Taiwan</option>
                        <%end if%>
                        <%if strReleaseTeam = "3" then%>
                        <option selected value="3">Mobility</option>
                        <%else%>
                        <option value="3">Mobility</option>
                        <%end if%>
                        <%if strReleaseTeam = "0"  then%>
                        <option value="0" selected>No Release Team</option>
                        <%else%>
                        <option value="0">No Release Team</option>
                        <%end if%>
                    </select>
                </td>

                <td width="120" style="vertical-align: top">
                    <strong><font size="2">PnP ID:</font></strong>
                </td>
                <td>
                    <%if strMachinePnPID = "" then
			Response.write "<a ID=MachinePNPLink href=""javascript:SelectMachinePNPID();"">Add ID</a>"
		  else
			Response.write	"<a ID=MachinePNPLink  href=""javascript:SelectMachinePNPID();"">" & replace(strMachinePNPID,",",", ") & "</a>"
		  end if
                    %>
                    <input type="hidden" id="txtMachinePNPID" name="txtMachinePNPID" style="width: 160px;"
                        value="<%=strMachinePnPID%>">
                    <input type="hidden" id="txtMachinePNPComments" name="txtMachinePNPComments" style="width: 160px;"
                        value="<%=replace(strMachinePnPcomments,chr(161)&chr(168),"''")%>">
                </td>
            </tr>
            <tr style="display: none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">OTS Product Name:</font></strong>
                </td>
                <td colspan="3">
                    <input type="text" id="txtOTS" name="txtOTS" style="width: 180px;" value="<%=strOTSName%>"
                        maxlength="30">
                </td>
            </tr>
            <tr style="display: none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Reports:</font></strong>
                </td>
                <td colspan="3">
                    <input type="checkbox" <%=CheckReports%> id="chkReports" name="chkReports">
                    <font face="verdana" size="2">Enable real-time reports.</font>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Phase:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td>
                    <select id="cboPhase" name="cboPhase" style="width: 160px;" language="javascript"
                        onchange="return cboPhase_onchange()">
                        <%
				rs.open "spListProductStatuses",cn,adOpenStatic
				do while not rs.EOF
					if trim(strProductStatus) = trim(rs("ID")) then
						Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					else
						Response.Write "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					end if
				
					rs.MoveNext
				loop
				rs.Close
                        %>
                    </select>
                    <input type="hidden" id="tagPhase" name="tagPhase" value="<%=strProductStatus%>">
                </td>
                <td id="RegModelRow" style="display: <%=DisplayToolsProject%>" width="160" style="vertical-align: top">
                    <strong><font size="2">Regulatory&nbsp;Model:</font></strong>
                </td>
                <td id="RegModelRow2" style="display: <%=DisplayToolsProject%>">
                    <input type="text" id="txtRegulatoryModel" name="txtRegulatoryModel" style="width: 200px;"
                        value="<%=strRegulatoryModel%>" maxlength="15">
                </td>
                <td id="ToolPMRow" style="display: <%=DisplayToolsProject2%>" width="160" style="vertical-align: top">
                    <strong><font size="2">Project&nbsp;Manager:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td id="ToolPMRow2" style="display: <%=DisplayToolsProject2%>">
                    <select id="cboToolsPM" name="cboToolsPM" style="width: 160px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option></option>
                        <%=strToolsPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdToolsPMAdd" name="cmdToolsPMAdd"
                        language="javascript" onclick="return cmdToolsPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Minimum RoHS Level:</font></strong><font color="red" size="1"></font>
                </td>
                <td>
                    <select id="cboMinRoHS" name="cboMinRoHS" style="width: 160px;" language="javascript">
                    <% rs.open "spListRoHS",cn,adOpenStatic
				       do while not rs.EOF
					      if trim(strMinRoHSLevel) = trim(rs("ID")) then
						     Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					      else
						     Response.Write "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					      end if
				          rs.MoveNext
				       loop
				       rs.Close
                    %>
                    </select>
                </td>
                <td width="160" style="vertical-align: top" colspan="2">
                    <input type="checkbox" id="chkBSAMFlag" name="chkBSAMFlag" <%=CheckBSAMFlag%> />
                    <strong><font face="verdana" size="2"> BSAM Product?</font></strong>
                </td>
            </tr> 
            <tr id="CycleRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Group:</font></strong>
                </td>
                <td>
                    <%
        if trim(request("ID")) = "" then
            rs.open "spListProgramsForProduct 0" ,cn,adOpenForwardOnly
        else
            rs.open "spListProgramsForProduct " & clng( request("ID")),cn,adOpenForwardOnly
        end if
        if rs.eof and rs.bof then
            if trim(request("ID")) = "" then
                response.write "<a id=CycleLink href=""javascript: SelectedCycles(0);"">Add to Product Group</a>"
            else
                response.write "<a id=CycleLink href=""javascript: SelectedCycles(" & clng( request("ID")) & ");"">Add to Product Group</a>"
            end if
        else
            dim strCycleList
            strCycleList = ""
            response.write "<a id=CycleLink href=""javascript: SelectedCycles(" & clng( request("ID")) & ");"">"
            do while not rs.eof
                 strCycleList= strCycleList & ", " & rs("FullName")
                rs.movenext
            loop
            if strCycleList <> "" then
                strCycleList = mid(strCycleList,3)
                response.write strCycleList & "</a>"
            end if
        end if
        rs.close
        
                    %>
                    <input id="txtAddCycle" name="txtAddCycle" type="hidden" value="">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">RCTO&nbsp;Sites:</font></strong>
                </td>
                <td nowrap>
                    <%
       if trim(request("ID")) = "" then
           response.write "<a id=SiteLink href=""javascript: SelectedSites(0);"">Add Site</a>"
       elseif trim(strRCTOSites) = "" then
           response.write "<a id=SiteLink href=""javascript: SelectedSites(" & clng( request("ID")) & ");"">Add Site</a>"
       else
           response.write "<a id=SiteLink href=""javascript: SelectedSites(" & clng( request("ID")) & ");"">" & strRCTOSites & "</a>"
       end if
    
                    %><input id="txtRCTOSites" name="txtRCTOSites" type="hidden" value="<%=strRCTOSites%>">
                </td>
            </tr>
            <tr style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Fusion:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input type="checkbox" <%=CheckFusion%> id="chkFusion" name="chkFusion">
                                <font face="verdana" size="2"> Enable IRS image builds</font>
                                <span id="IRSBuildsText">&nbsp;&nbsp;<font size="1" color="green" face="verdana">Note: Only targeted and newly supported deliverables will be copied to IRS.</font></span>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

            <tr id="ActivitiesRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Development Activities:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input InitialValue="<%=CheckSMR%>" type="checkbox" <%=CheckSMR%> id="chkEnableSMR"
                                    name="chkEnableSMR">
                                <font face="verdana" size="2">Enable developers to SMR deliverables for this product.<br>
                                    <input InitialValue="<%=CheckDeliverables%>" type="checkbox" <%=CheckDeliverables%>
                                        id="chkEnableDeliverables" name="chkEnableDeliverables">
                                    <font face="verdana" size="2">Enable developers to release new deliverables to support
                                        this product.<br>
                                        <input InitialValue="<%=CheckDCR%>" <%=EnableDCR%> type="checkbox" <%=CheckDCR%>
                                            id="chkEnableDCR" name="chkEnableDCR">
                                        <font face="verdana" size="2">Enable new DCR's to be submitted for this product.<br>
                                            <input InitialValue="<%=CheckImages%>" type="checkbox" <%=CheckImages%> id="chkEnableImages"
                                                name="chkEnableImages">
                                            <font face="verdana" size="2">Enable Preinstall to build and release images.<br>
                                            </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Distribution List:</font></strong><span id="DistributionRow"
                        style="display: <%=DisplayToolsProject%>"><font color="red" size="1">&nbsp;*</font></span><br>
                    <span style="color: Blue; text-decoration: underline; cursor: pointer;" onclick="cmdAddDistribution_onclick()">
                        Add</span>
                </td>
                <td colspan="3">
                    <textarea rows="2" style="width: 720px" id="txtDistribution" name="txtDistribution"><%=strDistribution%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Conveyor Build Distribution List:</font></strong><br>
                    <span style="color: Blue; text-decoration: underline; cursor: pointer;" onclick="cmdAddCvrBuildDist_onclick()">
                        Add</span>
                </td>
                <td colspan="3">
                    <textarea rows="2" style="width: 720px" id="txtCvrBuildDist" name="txtCvrBuildDist"><%=strCvrBuildDist%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Conveyor&nbsp;Release Distribution List:</font></strong><br>
                    <span style="color: Blue; text-decoration: underline; cursor: pointer;" onclick="cmdAddCvrReleaseDist_onclick()">
                        Add</span>
                </td>
                <td colspan="3">
                    <textarea rows="2" style="width: 720px" id="txtCvrReleaseDist" name="txtCvrReleaseDist"><%=strCvrReleaseDist%></textarea>
                </td>
            </tr>
            <tr id="NotificationRow" style="display: <%=DisplayToolsProject2%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Notify&nbsp;On&nbsp;Closure:</font></strong><br>
                    <input id="cmdNotify" type="button" value="Add" onclick="cmdAddNotification_onclick()">
                </td>
                <td colspan="3">
                    <textarea rows="2" style="width: 720px" id="txtActionNotifyList" name="txtActionNotifyList"><%=strActionNotify%></textarea>
                </td>
            </tr>
            <tr style="display: none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Email:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input type="checkbox" <%=CheckEmail%> id="chkEmail" name="chkEmail">
                                <font face="verdana" size="2">Enable Email Notifications.</font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="CommodityRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Commodities:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input type="checkbox" <%=CheckCommodities%> id="chkCommodities" name="chkCommodities">
                                <font face="verdana" size="2">Include this product on the Commodity Matrix.</font><br />
                                <font face="verdana" size="2">WWAN Product: </font>
                                <SELECT id=cboWWAN name=cboWWAN>
			                        <OPTION selected value=""></OPTION>
			                        <%if strWWAN = "1" then%>
				                        <OPTION selected value="1">Yes</OPTION>
			                        <%else%>
				                        <OPTION value="1">Yes</OPTION>
			                        <%end if%>
			                        <%if strWWAN = "0" then%>
				                        <OPTION selected value="0">No</OPTION>
			                        <%else%>
				                        <OPTION value="0">No</OPTION>
			                        <%end if%>
		                        </SELECT>
		                        <br />
                                <INPUT style="Display:<%=strHWStatusDisplay%>" type="checkbox" id=chkCommodityLock name=chkCommodityLock <%=strHWStatus%>>
                                <%if strHWStatusDisplay = "" then%>
                                    <font face="verdana" size="2">Locked Status</font>
                                <%else%>
                                    <font face="verdana" size="2">Status is locked</font>
                                <%end if%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
		        <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>End&nbsp;of&nbsp;Service&nbsp;Life:</font></strong></td>
		        <TD colspan="3">
		            <INPUT type="text" id=txtServiceEndDate name=txtServiceEndDate value="<%=strServiceLifeDate%>" <%=strEditServiceEOLDate%>>
			        <a href="javascript: cmdDate_onclick('txtServiceEndDate', '<%=strEditServiceEOLDate%>')"><img ID="picTarget" SRC="../../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
                    <%if strEditServiceEOLDate = "disabled" then%>
                        <span style="font-size: xx-small; color: Red">NOTE: Users may edit this date when the product is in Post-Production</span>
                    <%end if%>
		        </TD>
	        </tr>
            <tr id="MDARow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">MDA Compliance:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input type="checkbox" <%=CheckMDA%> id="chkMdaCompliance" name="chkMdaCompliance">
                                <font face="verdana" size="2">Include this product on the MDA Compliance Report.</font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
          
            <tr style="display: none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product&nbsp;Releases:</font></strong><font color="red" size="1">
                        *</font>
                </td>
                <td colspan="3">
                    <table width="100%" id="TableReleases">
             
                    </table>
                </td>
            </tr>
            <tr id="BrandRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Brands:</font></strong>
                </td>
                <td colspan="3">
                    <%

            response.write "<input id=""txtBrandFrom"" name=""txtBrandFrom"" type=""hidden"" value"""">"
            response.write "<input id=""txtBrandTo"" name=""txtBrandTo"" type=""hidden"" value"""">"
   
                    %>
                    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
                        border-left: steelblue 1px solid; border-bottom: steelblue 1px solid; height: 160px; width:700px; 
                        background-color: white" id="DIV3">
                        <table width="700" id="TableBrand">
                            <thead>
                                <tr style="position: relative; top: expression(document.getElementById('DIV3').scrollTop-2);">
                                    <td width="10" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;
                                    </td>
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Brand
                                    </td>
                                    <td width="302" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Series
                                    </td>
                                    <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Suffix&nbsp;
                                    </td>
                                </tr>
                            </thead>
                            <%
		if request("ID") <> "" then
			rs.open  "spListBrands4Product " & clng(request("ID")) & ",0",cn,adOpenForwardOnly
		else
			rs.open  "spListBrands ",cn,adOpenForwardOnly
		end if
		dim Brandcount 
		dim strRow
		dim strRowWait
		strRowWait = ""
        BrandCount=0
		do while not rs.eof
    		strRow = ""
			if rs("Active") or not isnull(rs("ProductBrandID")) then
				if BrandCount = 0 then
					Response.Write "<TR>"
				end if	
				BrandCount = BrandCount + 1
						
				if isnull(rs("ProductBrandID")) then
					strChecked = ""
				else
					strChecked = "checked"
					strBrandsLoaded = strBrandsLoaded & "," & trim(rs("ID"))
				end if

				strAbbr = rs("name") & "" 'rs("Abbreviation") & ""
                if request("ID") ="" then
                    strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""checkbox"" " & strChecked & " BrandName=""" & rs("Name") & """ title=" & rs("ID") & " id=chkBrands name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & "0" & ")"" value=""" & rs("ID") & """></TD>"
                else
		            strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""checkbox"" " & strChecked & " BrandName=""" & rs("Name") & """ title=" & rs("ID") & " id=chkBrands name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & request("ID") & ")"" value=""" & rs("ID") & """></TD>"
                end if
                if strChecked <> "" then 'currentuserid = 31 and
                    if rs("Active") then
                        strRow = strRow &  "<TD nowrap><font face=verdana size=2><a href=""javascript: ChooseNewBrand(" & rs("ID")& ");"">" & strAbbr & "</a>&nbsp;</font></TD>"
                    else
                        strRow = strRow &  "<TD nowrap><font face=verdana size=2><a href=""javascript: ChooseNewBrand(" & rs("ID")& ");"">" & strAbbr & "&nbsp;(old)</a></font></TD>"
                    end if
                else
                    if rs("Active") then
                        strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & strAbbr & "&nbsp;</font></TD>"
                    else
                        strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & strAbbr & "&nbsp;(old)</font></TD>"
                    end if
                end if
                strRow = strRow &  "<TD><INPUT type=""text"" id=tagSeries name=tagSeries style=""Display:none"" value=""" & rs("SeriesSummary") & """>"
			
				if strChecked = "" then
					showSeries = "none"
				else
					showSeries = ""
				end if
				
				'Pull series
				strSeriesID = ""
				strSeriesName = ""
				if not isnull(rs("ProductBrandID")) then
					set rs2 = server.CreateObject("ADODB.recordset")
					rs2.Open "spListSeries4Brand " &  rs("ProductBrandID"),cn,adOpenForwardOnly
					strSeriesID = ""
					strSeriesName = ""
					do while not rs2.EOF
						strSeriesID = strSeriesID & "," & rs2("ID")
						strSeriesName = strSeriesName & "," & rs2("Name")
						rs2.MoveNext
					loop
					set rs2 = nothing
					if strSeriesID <> "" then
						strSeriesID = mid(strSeriesID,2)
						strSeriesName = mid(strSeriesName,2)
					end if
				end if
				SeriesIDArray = split(strSeriesID & ",,,,,,,",",")
				SeriesNameArray = split(strSeriesName & ",,,,,,,",",")
				strRow = strRow &  "<DIV style=""display:" & showSeries & """ id=DivSeries" & trim(rs("ID")) & ">"
				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesA" & trim(rs("ID")) & " name=txtSeriesA" & trim(rs("ID")) & " value=""" & SeriesNameArray(0) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesA" & trim(rs("ID")) & " name=tagSeriesA" & trim(rs("ID")) & " value=""" & SeriesNameArray(0) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDA" & trim(rs("ID")) & " name=txtSeriesIDA" & trim(rs("ID")) & " value=""" & SeriesIDArray(0) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesB" & trim(rs("ID")) & " name=txtSeriesB" & trim(rs("ID")) & " value=""" & SeriesNameArray(1) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesB" & trim(rs("ID")) & " name=tagSeriesB" & trim(rs("ID")) & " value=""" & SeriesNameArray(1) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDB" & trim(rs("ID")) & " name=txtSeriesIDB" & trim(rs("ID")) & " value=""" & SeriesIDArray(1) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesC" & trim(rs("ID")) & " name=txtSeriesC" & trim(rs("ID")) & " value=""" & SeriesNameArray(2) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesC" & trim(rs("ID")) & " name=tagSeriesC" & trim(rs("ID")) & " value=""" & SeriesNameArray(2) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDC" & trim(rs("ID")) & " name=txtSeriesIDC" & trim(rs("ID")) & " value=""" & SeriesIDArray(2) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesD" & trim(rs("ID")) & " name=txtSeriesD" & trim(rs("ID")) & " value=""" & SeriesNameArray(3) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesD" & trim(rs("ID")) & " name=tagSeriesD" & trim(rs("ID")) & " value=""" & SeriesNameArray(3) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDD" & trim(rs("ID")) & " name=txtSeriesIDD" & trim(rs("ID")) & " value=""" & SeriesIDArray(3) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesE" & trim(rs("ID")) & " name=txtSeriesE" & trim(rs("ID")) & " value=""" & SeriesNameArray(4) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesE" & trim(rs("ID")) & " name=tagSeriesE" & trim(rs("ID")) & " value=""" & SeriesNameArray(4) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDE" & trim(rs("ID")) & " name=txtSeriesIDE" & trim(rs("ID")) & " value=""" & SeriesIDArray(4) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesF" & trim(rs("ID")) & " name=txtSeriesF" & trim(rs("ID")) & " value=""" & SeriesNameArray(5) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesF" & trim(rs("ID")) & " name=tagSeriesF" & trim(rs("ID")) & " value=""" & SeriesNameArray(5) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDF" & trim(rs("ID")) & " name=txtSeriesIDF" & trim(rs("ID")) & " value=""" & SeriesIDArray(5) & """>"

				strRow = strRow &  "</DIV></TD>"
				
				strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & rs("Suffix") & "&nbsp;</font></TD></TR>"

				if strChecked = "checked" then
				    Response.Write strRow
				else
				    strRowWait = strRowWait & strRow
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
		if strRowWait <> "" then
		    response.write strRowWait
		end if
		if strBrandsLoaded <> "" then
			strBrandsLoaded = mid(strBrandsLoaded,2)
		end if
                            %>
                        </table>
                    </div>
                </td>
            </tr>
            <tr style="display:none">
                <td id="Td1" width="160" style="vertical-align: top">
                    <strong><font size="2">Service&nbsp;Tag:</font></strong>
                </td>
                <td id="Td2" style="display: <%=DisplayToolsProject%>" colspan="3">
                    <input type="text" id="txtServiceTag" name="txtServiceTag" style="width: 720px;"
                        value="<%=strServiceTag%>" maxlength="100">
                </td>
            </tr>
            <tr style="display:none">
                <td id="Td3" width="160" style="vertical-align: top" >
                    <strong><font size="2">BIOS&nbsp;Branding:</font></strong>
                </td>
                <td id="Td4" style="display: <%=DisplayToolsProject%>" colspan="3">
                    <input type="text" id="txtBIOSBranding" name="txtBIOSBranding" style="width: 720px;"
                        value="<%=strBIOSBranding%>" maxlength="100">
                </td>
            </tr>
            <tr id="OSRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">OS Support:</font></strong>
                </td>
                <td colspan="3">
                    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
                        border-left: steelblue 1px solid; border-bottom: steelblue 1px solid; height: 160px; width:700px;
                        background-color: white" id="DIV2">
                        <table width="700" id="TableOS">
                            <thead>
                                <tr style="position: relative; top: expression(document.getElementById('DIV2').scrollTop-2);">
                                    <td width="282" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Name
                                    </td>
                                    <td bgcolor="#c9ddff" width="70" style="border-right: 1px outset; border-top: 1px outset;
                                        border-left: 1px outset; border-bottom: 1px outset">
                                        &nbsp;Preinstall&nbsp;
                                    </td>
                                    <td width="90" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" align="middle" bgcolor="#c9ddff">
                                        &nbsp;Web&nbsp;
                                    </td>
                                </tr>
                            </thead>
                            <%
				strFullOSList=""
				if request("ID") = "" then
					rs.Open "spListProductOSAll 0",cn,adOpenForwardOnly
				else
					rs.Open "spListProductOSAll " & clng(request("ID")),cn,adOpenForwardOnly
				end if
				do while not rs.EOF
					strFullOSList = strFullOSList & "," & rs("ID")
					Response.Write "<TR>"
					Response.Write "<TD>" & rs("Name") & "</TD>"
					if rs("Preinstall") & "" = "" or lcase(trim(rs("Preinstall") & "")) = "false" then
						Response.Write "<TD align=middle><INPUT type=""checkbox"" id=chkPreinstallOS LANGUAGE=javascript onclick=""ProgramInput.txtOSListChanged.value=1;"" name=chkPreinstallOS value=""" & rs("ID") & """></TD>"
					else
						Response.Write "<TD align=middle><INPUT checked type=""checkbox"" id=chkPreinstallOS LANGUAGE=javascript onclick=""ProgramInput.txtOSListChanged.value=1;"" name=chkPreinstallOS value=""" & rs("ID") & """></TD>"
					end if
					if rs("Web") & "" = "" or lcase(trim(rs("Web") & "")) = "false" then
						Response.Write "<TD align=middle><INPUT type=""checkbox"" id=chkWebOS LANGUAGE=javascript onclick=""ProgramInput.txtOSListChanged.value=1;"" name=chkWebOS value=""" & rs("ID") & """></TD>"
					else
						Response.Write "<TD align=middle><INPUT checked type=""checkbox"" id=chkWebOS LANGUAGE=javascript onclick=""ProgramInput.txtOSListChanged.value=1;"" name=chkWebOS value=""" & rs("ID") & """></TD>"
					end if
					
					Response.Write "</TR>"
					rs.MoveNext	
				loop
				rs.Close
				if strFullOSList <> "" then
					strFullOSList = mid(strFullOSList,2)
				end if
                            %>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Description:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td colspan="3">
                    <textarea rows="2" cols="20" id="txtDescription" name="txtDescription" style="width: 720px;
                        height: 120px;" language="javascript" onkeypress="return txtDescription_onkeypress()"><%=strDescription%></textarea>
                </td>
            </tr>
            <tr id="ApproverRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <span style="font-size: x-small; font-weight: bold;">DCR&nbsp;Setup:</span>
                </td>
                <td valign="top" colspan="3">
                    &nbsp;Default DCR Owner: 
                    <select id="cboDCRDefaultOwner" name="cboDCRDefaultOwner">
                    <%

                        if trim(strDefaultDCROwner) = "1" then
                            response.write "<option selected value=1>Configuration Manager</option>"
                            response.write "<option value=2>Program Office Manager</option>"
                        else
                            response.write "<option value=1>Configuration Manager</option>"
                            response.write "<option selected value=2>Program Office Manager</option>"
                        end if

                    %>
                    </select>
                    <hr />
                    <input type="checkbox" <%=AddDCRNotificationList%> id="chkAddDCRNotificationList" name="chkAddDCRNotificationList">
                    <font face="verdana" size="2">Add the DCR notification list (Mobile Excal Notification - DCRs) for critical changes.</font>
                    <hr />
                    <table>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDcrToCm%> type="radio" id="chkDCRAutoOpen" name="chkDCRAutoOpen" 
                                    value="1" />
                            </td>
                            <td>
                                DCR is assigned to the CM/POPM for review.
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDCRAutoOpen%> type="radio" id="chkDCRAutoOpen" name="chkDCRAutoOpen"
                                    value="2" />
                            </td>
                            <td>
                                Automatically assign Primary System Team members as approvers and set status to "Investigating" for new DCRs.
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDCRNoOdm%>  type="radio" id="Radio3" name="chkDCRAutoOpen"
                                    value="4" />
                            </td>
                            <td>
                                Automatically assign Primary System Team members <span style="color: Red"> (Excluding ODM) </span> as approvers and set status to "Investigating" for new DCRs.
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDcrToList%> type="radio" id="Radio1" name="chkDCRAutoOpen" 
                                    value="3" />
                            </td>
                            <td>
                                Automatically assign the following users as approvers and set status to "Investigating"
                                for new DCRs. <span style="font-size: xx-small; color: Red">NOTE:Separate email addresses
                                    with a semicolon ( ; )</span><br />
                                <textarea rows="2" style="width: 690px" id="txtDCRApproverList" name="txtDCRApproverList"><%=strDCRApproverList%></textarea>
                                <input id="btnAddApprover" type="button" value="Add" onclick="cmdAddApprover_onclick()">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table id="tabStatusReport" style="display: none; width:900px; border-collapse: collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Preinstall Cut-Off:</font></strong>
                </td>
                <td>
                    <select id="cboPinCutoff" name="cboPinCutoff">
                        <option>
                            <%=strPinCutoffValue%></option>
                        <option>Mon 8:00 AM</option>
                        <option>Tue 8:00 AM</option>
                        <option>Wed 8:00 AM</option>
                        <option>Thu 8:00 AM</option>
                        <option>Fri 8:00 AM</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Objectives:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtObjective" name="txtObjective" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtObjective_onkeypress()"><%=strObjective%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Base Unit:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtBaseUnit" name="txtBaseUnit" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtBaseUnit_onkeypress()"><%=strBaseUnit%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2" id="ROMText">
                        <%if clng(strProductStatus) < 3 then%>
                        Current&nbsp;ROM:
                        <%else%>
                        Current&nbsp;Factory&nbsp;ROM:
                        <%end if%>
                    </font></strong>
                </td>
                <td>
                    <input type="text" id="txtCurrentROM" name="txtCurrentROM" style="width: 100%;" value="<%=strCurrentROM%>"
                        maxlength="200">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Current&nbsp;Web&nbsp;ROM:</font></strong>
                </td>
                <td>
                    <input type="text" id="txtCurrentWebROM" name="txtCurrentWebROM" style="width: 100%;"
                        value="<%=strCurrentWebROM%>" maxlength="200">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">OS Support:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtOSSupport" name="txtOSSupport" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtOSSupport_onkeypress()"><%=strOSSupport%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Image PO:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtImagePO" name="txtImagePO" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtImagePO_onkeypress()"><%=strImagePO%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Image Changes:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtImageChanges" name="txtImageChanges" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtImageChanges_onkeypress()"><%=strImageChanges%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Common Images:</font></strong>
                </td>
                <td>
                    <input type="text" id="txtCommonIMages" name="txtCommonImages" style="width: 100%;"
                        value="<%=strCommonImages%>" maxlength="300">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Certification Status:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtCertificationStatus" name="txtCertificationStatus"
                        style="width: 100%; height: 80px;" language="javascript" onkeypress="return txtCertificationStatus_onkeypress()"><%=strCertificationStatus%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Software QA Status:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtSWQAStatus" name="txtSWQAStatus" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtSWQAStatus_onkeypress()"><%=strSWQAStatus%></textarea>
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Platform Status:</font></strong>
                </td>
                <td>
                    <textarea rows="2" cols="20" id="txtPlatformStatus" name="txtPlatformStatus" style="width: 100%;
                        height: 80px;" language="javascript" onkeypress="return txtPlatformStatus_onkeypress()"><%=strPlatformStatus%></textarea>
                </td>
            </tr>
        </table>
        <span id="tabAccess" style="display: none"><font size="1" face="verdana"><a href="javascript: ChooseEmployee2();">
            Add Person</a><br>
            <br>
        </font>
            <table style="border-collapse:collapse; width:900px;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
                <tr bgcolor="Wheat">
                    <td>
                        <b>Product Access List - Grant Access to Add/Update Tasks, Roadmaps, and Issues</b>
                    </td>
                </tr>
                <tr>
                    <td nowrap width="100%">
                        <table id="ToolAccessTable" style="width: 100%" cellpadding="2" cellspacing="0">
                            <tr>
                                <td>
                                    <font size="1" face="verdana"><b>Person</b></font>
                                </td>
                                <td>
                                    <font size="1" face="verdana"><b>Access</b></font>
                                </td>
                            </tr>
                            <tr bgcolor="white">
                                <td class="OTSComponentCell">
                                    Owners of Open Action Items
                                </td>
                                <td class="OTSComponentCell">
                                    Required
                                </td>
                            </tr>
                            <tr bgcolor="white">
                                <td class="OTSComponentCell" id="PMAccessCell">
                                    Project Manager
                                </td>
                                <td class="OTSComponentCell">
                                    Required
                                </td>
                            </tr>
                            <%
				strToolAccessIDList = CleanIDList(strToolAccessIDList)
				if strToolAccessIDList <> "" then
					strSQL = "Select ID, Name from employee where Id in (" & strToolAccessIDList & ")"
					rs.Open strSQL,cn,adOpenKeyset
					do while not rs.EOF
						Response.Write "<TR id=""ToolAccessRow" & trim(rs("ID")) & """ bgcolor=white><TD class=OTSComponentCell><INPUT style=""display:none"" type=""checkbox"" checked id=chkToolAccessID" & rs("ID") & " name=chkToolAccessID value=""" & rs("ID") & """>" & rs("Name") & "</TD><TD nowrap class=OTSComponentCell><a href=""javascript: RemoveToolAccess(" & rs("ID") & ");"">Remove</a></TD></TR>"
						rs.MoveNext
					loop
					rs.Close
				end if
                            %>
                        </table>
                    </td>
                </tr>
            </table>
        </span><span id="tabFiles" style="display: none">
            <table style="width:900px; border-collapse:collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">PDD Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtPDDPath" name="txtPDDPath" maxlength="256"
                            value="">
                        <input type="hidden" id="tagPDDPath" name="tagPDDPath" value="<%=strPDDPath%>" maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">SCM Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtSCMPath" name="txtSCMPath" maxlength="256"
                            value="">
                        <input type="hidden" id="tagSCMPath" name="tagSCMPath" value="<%=strSCMPath%>" maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">STL Status Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtSTLPath" name="txtSTLPath" maxlength="256"
                            value="">
                        <input type="hidden" id="tagSTLPath" name="tagSTLPath" value="<%=strSTLPath%>" maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">Product Data<br>
                            Matrices Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtProgramMatrixPath" name="txtProgramMatrixPath"
                            maxlength="256" value="">
                        <input type="hidden" id="tagProgramMatrixPath" name="tagProgramMatrixPath" value="<%=strProgramMatrixPath%>"
                            maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">Accessory<br>
                            Documents Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtAccessoryPath" name="txtAccessoryPath"
                            maxlength="256" value="">
                        <input type="hidden" id="tagAccessoryPath" name="tagAccessoryPath" value="<%=strAccessoryPath%>"
                            maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">Test Paths:</font></strong>
                    </td>
                    <td>
                        <font size="1" face="verdana" color="green">Click a link to test the path. This will
                            open the location you specified in a new window. If will not open very fast if you
                            are not logged into the domain.</font><br>
                        <br>
                        <font size="1" face="verdana"><a href="javascript:TestPath(1);">PDD</a> | <a href="javascript:TestPath(2);">
                            SCM</a> | <a href="javascript:TestPath(3);">STL Status</a> | <a href="javascript:TestPath(4);">
                                Product Data Matrices</a> | <a href="javascript:TestPath(5);">Accessory Documents</a></font>
                    </td>
                </tr>
            </table>
        </span><span style="display: none" id="tabOTS">
            <%
	dim ShowComponentAddLink
	dim ShowCMTRemoveLink
	dim blnOTSDown
	blnOTSDown = false
	
	if request("ID") = "" then 'new Product
		ShowComponentAddLink = false
		ShowCMTRemoveLink = false
		Response.Write "<INPUT type=""hidden"" id=txtMissingComponents name=txtMissingComponents value=""0"">"
	else
		on error resume next
		rs.Open "spGetOTSComponentCount " & clng(request("ID")),cn,adOpenStatic
		if cn.Errors.Count > 0 then
		    blnOTSDown = true
		end if
		on error goto 0
		if blnOTSDown then
            response.write "OTS Is Currently Down"
   	        ShowComponentAddLink = false
	        ShowCMTRemoveLink = false
	        Response.Write "<INPUT type=""hidden"" id=txtMissingComponents name=txtMissingComponents value=""0"">"
        else
            if rs.EOF and rs.BOF then
    	        ShowComponentAddLink = false
		        ShowCMTRemoveLink = false
		        Response.Write "<INPUT type=""hidden"" id=txtMissingComponents name=txtMissingComponents value=""0"">"
	        else
		        ShowCMTRemoveLink = false 
		        if rs("ExcaliburMissing") > 0 then
			        ShowComponentAddLink = true
		        else
			        ShowComponentAddLink = false
		        end if
		        Response.Write "<INPUT type=""hidden"" id=txtMissingComponents name=txtMissingComponents value=""" &rs("ExcaliburMissing")  & """>"
            end if
	        rs.Close
        end if
    end if
Response.Write "<font size=1 face=verdana>"

if strIsSEPM = "True" then
    response.write "<span>Include this product in SI's Affected Product evaluation function:</span>"
    %>
    <table>
        <tr>
            <td style="vertical-align: top;">
              <input <%=strAlways%> type="radio" id="rblAffectedProduct" name="rblAffectedProduct" value="1" language="javascript" onclick="return rblAffectedProduct_onclick(1)"/>
            </td>
            <td>
              <font face="verdana" size="1">Always</font> 
            </td>
        </tr>
        <tr>
            <td style="vertical-align: top;">
              <input <%=strNever%> type="radio" id="rblAffectedProduct" name="rblAffectedProduct" value="2" language="javascript" onclick="return rblAffectedProduct_onclick(2)"/>
            </td>
            <td>
              <font face="verdana" size="1">Never</font> 
            </td>
        </tr>
        <tr>
            <td style="vertical-align: top;">
                <input <%=strUntil%> type="radio" id="rblAffectedProduct" name="rblAffectedProduct" value="3" language="javascript" onclick="return rblAffectedProduct_onclick(3)"/>
            </td>
            <td>
               <font face="verdana" size="1">Until</font>&nbsp;&nbsp;             
            </td>
            <td>
               <select <%=strcboMilestones%> id="cboMilestones" name="cboMilestones" style="width: 275px;" language="javascript" onkeypress="return combo_onkeypress()"
                 onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                 <option></option>
                 <%=strReqMilestones%>
               </select>            
            </td>      
        </tr>
    </table>
    <%
end if

if ShowComponentAddLink then
	response.write "<span ID=OTSAddComponentMessage2><a href=""javascript: AddComponents();"">Add Standard OTS Component Set</a></span>"
	if ShowCMTRemoveLink then
		Response.Write "<span id=OTSAddComponentMessage3>&nbsp;|&nbsp;</span>"	
	else
		Response.Write "<span id=OTSAddComponentMessage3></span>"	
	end if
else
	Response.Write "<span id=OTSAddComponentMessage3></span>"	
	Response.Write "<span id=OTSAddComponentMessage2></span>"	
end if

'if ShowCMTRemoveLink then
'	response.write "<a href=""javascript: DeactivateCMTComponents();"">Deactivate Selected CMT Components</a>"
'end if
if ShowComponentAddLink then 'or ShowCMTRemoveLink then
	Response.Write "<span ID=OTSAddComponentMessage1>Fields Missing</span><BR><BR>"
else
	Response.Write "<span ID=OTSAddComponentMessage1></span>"
end if
            %>
    </font>
    <%if not blnOTSDown then %>
    <table border="1" cellpadding="2" cellspacing="0" style="width:900px; border-collapse:collapse;" bgcolor="cornsilk"
        bordercolor="tan">
        <tr bgcolor="Wheat">
            <td>
                <b>OTS Common Components</b>
            </td>
        </tr>
        <%
	if request("ID") = "" then
		rs.Open "spListProductOTSComponents 0",cn,adOpenStatic
	else
		rs.Open "spListProductOTSComponents " & clng(request("ID")) & "",cn,adOpenStatic
	end if
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
        <%	else%>
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
                                <font face="verdana" size="1"><b>Core Team</b></font>
                            </td>
                            <td>
                                <font face="verdana" size="1"><b>PM</b></font>
                            </td>
                            <td>
                                <font face="verdana" size="1"><b>Developer</b></font>
                            </td>
                        </tr>
                        <%		do while not rs.EOF
		if rs("ID") = 0 then
			Response.Write "<TR bgcolor=Lavender><td class=OTSComponentCell><INPUT style=""Display:none;WIDTH:16;Height:16"" value=""" & rs("Partnumber") & """ type=""checkbox"" id=chkCMT name=chkCMT>CMT</td>"
		else
			Response.Write "<TR bgcolor=white><td class=OTSComponentCell>Excalibur&nbsp;&nbsp;</td>"
		end if
                        %>
                        <td class="OTSComponentCell">
                            <%=rs("ErrorType")%>
                        </td>
                        <td class="OTSComponentCell">
                            <%=rs("category")%>
                        </td>
                        <td class="OTSComponentCell">
                            <%=rs("Component")%>
                        </td>
                        <%if rs("ID") = 0 then%>
                        <td class="OTSComponentCell">
                            <%=rs("CoreTeam")&""%>
                        </td>
                        <td class="OTSComponentCell">
                            <%=rs("PM")&""%>
                        </td>
                        <td class="OTSComponentCell">
                            <%=rs("Developer")&""%></a>
                        </td>
                        <%else%>
                        <td id='OTSCoreTeam<%=trim(rs("ID"))%>' class="OTSComponentCell">
                            <a href="javascript: EditOTSCoreTeam(<%=rs("ID")%>,<%=rs("OTSComponentID")%>,<%=rs("CoreTeamID")%>)">
                                <%=longname(rs("CoreTeam")&"")%></a>
                        </td>
                        <td id='OTSPM<%=trim(rs("ID"))%>' class="OTSComponentCell">
                            <a href="javascript: EditOTSPM(<%=rs("ID")%>,<%=rs("PMID")%>)">
                                <%=longname(rs("PM")&"")%></a>
                        </td>
                        <td id='OTSDeveloper<%=trim(rs("ID"))%>' class="OTSComponentCell">
                            <a href="javascript: EditOTSDeveloper(<%=rs("ID")%>,<%=rs("DeveloperID")%>)">
                                <%=longname(rs("Developer")&"")%></a>
                        </td>
                        <%end if%>
        </tr>
        <%			rs.MoveNext
		loop
        %>
    </table>
    </div> </td></tr>
    <%		
	end if
	rs.Close
    %>
    </Table>
    <%end if 'OTS Down Check%>
    </span>
    <table id="tabSystemTeam" style="display: none; width:900px; border-collapse:collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
        <tr bgcolor="Wheat">
            <td colspan="4">
                <b>Primary Team Members</b>
            </td>
        </tr>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">System&nbsp;Manager:</font></strong><font color="red" size="1">&nbsp;*</font>
            </td>
            <td>
                <select id="cboSM" name="cboSM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                    onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                    <option></option>
                    <%=strSMList%>
                </select>&nbsp;<input type="button" value="Add" id="cmdSMAdd" name="cmdSMAdd" language="javascript"
                    onclick="return cmdSMAdd_onclick()">
            </td>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Commercial&nbsp;Marketing:</font></strong>
            </td>
            <td>
                <select id="cboComMarketing" name="cboComMarketing" style="width: 140px;" language="javascript"
                    onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                    onkeydown="return combo_onkeydown()">
                    <option selected value="0"></option>
                    <%=strComMarketingList%>
                </select>&nbsp;<input type="button" value="Add" id="cmdComMarketingAdd" name="cmdComMarketingAdd"
                    language="javascript" onclick="return cmdComMarketingAdd_onclick()">
            </td>
        </tr>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2" id="lblPOPM">
                    <%=strPOPMLabel%></font></strong><font color="red" size="1">&nbsp;*</font>
            </td>
            <td>
                <select id="cboPM" name="cboPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                    onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                    <option></option>
                    <%=strPMList%>
                </select>&nbsp;<input type="button" value="Add" id="cmdPMAdd" name="cmdPMAdd" language="javascript"
                    onclick="return cmdPMAdd_onclick()">
            </td>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Consumer&nbsp;Marketing:</font></strong>
            </td>
            <td>
                <select id="cboConsMarketing" name="cboConsMarketing" style="width: 140px;" language="javascript"
                    onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                    onkeydown="return combo_onkeydown()">
                    <option selected value="0"></option>
                    <%=strConsMarketingList%>
                </select>&nbsp;<input type="button" value="Add" id="cmdConsMarketingAdd" name="cmdConsMarketingAdd"
                    language="javascript" onclick="return cmdConsMarketingAdd_onclick()">
            </td>
            <tr>
                <% if trim(strDevCenter) <> "2" and trim(strDevCenter) <> "6"then 'This is a Commercial Product                     
                    %>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2" id="lblTDCCM">Program&nbsp;Office&nbsp;Manager:</font><span
                        style="display: none" id="POPMRequired"><font color="red" size="1">&nbsp;*</font>
                    </span></strong>
                </td>
                <td valign="top">
                    <span style="display: none" id="POPMConsOnly"><font face="verdana" size="2" color="green">&nbsp;Consumer Products
                        Only.</font></span>
                    <select id="cboTDCCM" name="cboTDCCM" style="width: 140px;"
                        language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()"
                        onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option></option>
                        <%=strTDCCMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdTDCCMAdd"
                        name="cmdTDCCMAdd" language="javascript" onclick="return cmdTDCCMAdd_onclick()">
                </td>
                <%else 'This is a Consumer Product%>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2" id="lblTDCCM">Configuration&nbsp;Manager:</font><span style="display: none"
                        id="POPMRequired"><font color="red" size="1">&nbsp;*</font></span></strong>
                </td>
                <td>
                    <span style="display: none" id="POPMConsOnly"><font face="verdana" size="2" color="green">
                        &nbsp;Consumer Products Only.</font></span>
                    <select id="cboTDCCM" name="cboTDCCM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option></option>
                        <%=strTDCCMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdTDCCMAdd" name="cmdTDCCMAdd"
                        language="javascript" onclick="return cmdTDCCMAdd_onclick()">
                </td>
                <%end if%>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">SMB&nbsp;Marketing:</font></strong>
                </td>
                <td>
                    <select id="cboSMBMarketing" name="cboSMBMarketing" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSMBMarketingList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSMBMarketingAdd" name="cmdSMBMarketingAdd"
                        language="javascript" onclick="return cmdSMBMarketingAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Systems&nbsp;Engineering&nbsp;PM:</font></strong><font color="red"
                        size="1">&nbsp;*</font>
                </td>
                <td>
                    <select id="cboSEPM" name="cboSEPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option></option>
                        <%=strSEPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSEPMAdd" name="cmdSEPMAdd"
                        language="javascript" onclick="return cmdSEPMAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Supply&nbsp;Chain:</font></strong>
                </td>
                <td>
                    <select id="cboSupplyChain" name="cboSupplyChain" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSupplyChainList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSupplyChainAdd" name="cmdSupplyChainAdd"
                        language="javascript" onclick="return cmdSupplyChainAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Platform&nbsp;Development&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPlatformDevelopment" name="cboPlatformDevelopment" style="width: 140px;"
                        language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()"
                        onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strPlatformDevelopmentList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdPlatformDevelopmentAdd" name="cmdPlatformDevelopmentAdd"
                        language="javascript" onclick="return cmdPlatformDevelopmentAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Service:</font></strong>
                </td>
                <td>
                    <select id="cboService" name="cboService" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strServiceList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdServiceAdd" name="cmdServiceAdd"
                        language="javascript" onclick="return cmdServiceAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Commodity&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPDE" name="cboPDE" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strPDEList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdPDEAdd" name="cmdPDEAdd" language="javascript"
                        onclick="return cmdPDEAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Quality:</font></strong>
                </td>
                <td>
                    <select id="cboQuality" name="cboQuality" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strQualityList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdQualityAdd" name="cmdQualityAdd"
                        language="javascript" onclick="return cmdQualityAdd_onclick()">
                </td>
            </tr>

            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">ODM&nbsp;System&nbsp;Engineering&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboODMSEPM" name="cboODMSEPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strODMSEPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdODMSEPMAdd" name="cmdODMSEPMAdd" language="javascript"
                        onclick="return cmdODMSEPMAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Procurement&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboProcurementPM" name="cboProcurementPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strProcurementPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdProcurementPMAdd" name="cmdProcurementPMAdd"
                        language="javascript" onclick="return cmdProcurementPMAdd_onclick()">
                </td>
            </tr>

            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Planning&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPlanningPM" name="cboPlanningPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strPlanningPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdPlanningPMAdd" name="cmdPlanningPMAdd" language="javascript"
                        onclick="return cmdPlanningPMAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
        
                </td>
                <td>
                </td>
            </tr>

            <tr bgcolor="Wheat">
                <td colspan="4">
                    <b>Extended Team Members</b>
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Program&nbsp;Coordinator:</font></strong>
                </td>
                <td>
                    <select id="cboPC" name="cboPC" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strPCList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdPCAdd" name="cmdPCAdd" language="javascript"
                        onclick="return cmdPCAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Chipset/Processor&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboProcessorPM" name="cboProcessorPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strProcessorPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdProcessorPMAdd" name="cmdProcessorPMAdd"
                        language="javascript" onclick="return cmdProcessorPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">PDM&nbsp;Team&nbsp;Member:</font></strong>
                </td>
                <td>
                    <select id="cboMarketingOps" name="cboMarketingOps" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strMarketingOpsList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdMarketingOpsAdd" name="cmdMarketingOpsAdd"
                        language="javascript" onclick="return cmdMarketingOpsAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Graphics&nbsp;Controller&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboGraphicsControllerPM" name="cboGraphicsControllerPM" style="width: 140px;"
                        language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()"
                        onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strGraphicsControllerPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdGraphicsControllerPMAdd" name="cmdGraphicsControllerPMAdd"
                        language="javascript" onclick="return cmdGraphicsControllerPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Documentation&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboDocPM" name="cboDocPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strDocPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdDocPMAdd" name="cmdDocPMAdd"
                        language="javascript" onclick="return cmdDocPMAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Video&nbsp;Memory&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboVideoMemoryPM" name="cboVideoMemoryPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strVideoMemoryPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdVideoMemoryPMAdd" name="cmdVideoMemoryPMAdd"
                        language="javascript" onclick="return cmdVideoMemoryPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Doc&nbsp;Kit&nbsp;Coordinator:</font></strong>
                </td>
                <td>
                    <select id="cboDKC" name="cboDKC" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strDKCList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdDKCAdd" name="cmdDKCAdd" language="javascript"
                        onclick="return cmdDKCAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Comm&nbsp;HW&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboCommHWPM" name="cboCommHWPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strCommHWPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdCommHWPMAdd" name="cmdCommHWPMAdd"
                        language="javascript" onclick="return cmdCommHWPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">SE&nbsp;Program&nbsp;Engineer:</font></strong>
                </td>
                <td>
                    <select id="cboSEPE" name="cboSEPE" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSEPEList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSEPEAdd" name="cmdSEPEAdd"
                        language="javascript" onclick="return cmdSEPEAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Accessory&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboAccessoryPM" name="cboAccessoryPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strAPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdAccessoryPMAdd" name="cmdAccessoryPMAdd"
                        language="javascript" onclick="return cmdAccessoryPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Preinstall&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPINPM" name="cboPINPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strPINPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdPINPMAdd" name="cmdPINPMAdd"
                        language="javascript" onclick="return cmdPINPMAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">SE&nbsp;Test&nbsp;Lead&nbsp;(Pri):</font></strong>
                </td>
                <td>
                    <select id="cboSETestLead" name="cboSETestLead" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSETestLeadList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSETestLeadAdd" name="cmdSETestLeadAdd"
                        language="javascript" onclick="return cmdSETestLeadAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">BIOS&nbsp;Lead:</font></strong>
                </td>
                <td>
                    <select id="cboBIOSLead" name="cboBIOSLead" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strBIOSLeadList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdBIOSLeadAdd" name="cmdBIOSLeadAdd"
                        language="javascript" onclick="return cmdBIOSLeadAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">SE&nbsp;Test&nbsp;Lead&nbsp;(Sec):</font></strong>
                </td>
                <td>
                    <select id="cboSETest" name="cboSETest" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSETestList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSETestAdd" name="cmdSETestAdd"
                        language="javascript" onclick="return cmdSETestAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">SC&nbsp;Factory&nbsp;Engineer:</font></strong>
                </td>
                <td>
                    <select id="cboFactoryEngineer" name="cboFactoryEngineer" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strFactoryEngineerList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdFactoryEngineerAdd" name="cmdFactoryEngineerAdd"
                        language="javascript" onclick="return cmdFactoryEngineerAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">ODM&nbsp;HW&nbsp;Test&nbsp;Lead:</font></strong>
                </td>
                <td>
                    <select id="cboODMTestLead" name="cboODMTestLead" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strODMTestLeadList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdODMTestLeadAdd" name="cmdODMTestLeadAdd"
                        language="javascript" onclick="return cmdODMTestLeadAdd_onclick()">
                </td>
            </tr>
            <tr style="vertical-align: top">
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Global&nbsp;Product&nbsp;LM:</font></strong>
                </td>
                <td>
                    <select id="cboGplm" name="cboGplm" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strGplmList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdGplmAdd" name="cmdGplmAdd"
                        language="javascript" onclick="return SystemTeamAdd('cboGplm')">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">WWAN&nbsp;Test&nbsp;Lead:</font></strong>
                </td>
                <td>
                    <select id="cboWWANTestLead" name="cboWWANTestLead" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strWWANTestLeadList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdWWANTestLeadAdd" name="cmdWWANTestLeadAdd"
                        language="javascript" onclick="return cmdWWANTestLeadAdd_onclick()">
                </td>
            </tr>
            <tr style="vertical-align: top">
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Service&nbsp;BOM&nbsp;Analyst:</font></strong>
                </td>
                <td>
                    <select id="cboSpdm" name="cboSpdm" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSpdmList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSpdmAdd" name="cmdSpdmAdd"
                        language="javascript" onclick="return SystemTeamAdd('cboSpdm')">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Sustaining&nbsp;SEPM:</font></strong>
                </td>
                <td>
                    <select id="cboSustainingSEPM" name="cboSustainingSEPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSustainingSEPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSustainingSEPMAdd" name="cmdSustainingSEPMAdd"
                        language="javascript" onclick="return cmdSustainingSEPMAdd_onclick()">
                </td>
            </tr>
            <tr style="vertical-align: top; display: none">
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Service&nbsp;BOM&nbsp;Analyst:</font></strong>
                </td>
                <td>
                    <select id="cboSba" name="cboSba" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSbaList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSbaAdd" name="cmdSbaAdd" language="javascript"
                        onclick="return SystemTeamAdd('cboSba')">
                </td>
                <td width="120" style="display: none; vertical-align: top">
                    <strong><font size="2">Sustaining&nbsp;Manager:</font></strong>
                </td>
                <td>
                    <select id="cboSustainingMgr" name="cboSustainingMgr" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSustainingMgrList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSustainingMgrAdd" name="cmdSustainingMgrAdd"
                        language="javascript" onclick="return cmdSustainingMgrAdd_onclick()">
                </td>
            </tr>
                
                <!-- LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB -->
            <tr style="vertical-align: top">
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Systems&nbsp;Engineering&nbsp;PC:&nbsp;</font></strong>
                </td>
                <td>
                    <select id="cboSysEngrProgramCoordinator" name="cboSysEngrProgramCoordinator" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strSysEngrProgramCoordinatorList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdSysEngrProgramCoordinatorAdd" name="cmdSysEngrProgramCoordinatorAdd"
                        language="javascript" onclick="return cmdSysEngrProgramCoordinatorAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">ODM&nbsp;PIN&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboODMPIMPM" name="cboODMPIMPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <option selected value="0"></option>
                        <%=strODMPIMPMList%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdODMPIMPMAdd" name="cmdODMPIMPMAdd"
                        language="javascript" onclick="return cmdODMPIMPMAdd_onclick()">
                </td>
            </tr>
                <!-- LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB -->
                
    </table>
    <!-- End tabSystemTeam -->
    <input style="display: none" type="text" id="txtID" name="txtID" value="<%=request("ID")%>">
    <input type="hidden" id="txtProductName" name="txtProductName">
    <input type="hidden" id="txtServiceLifeDate" name="txtServiceLifeDate" value="<%=strServiceLifeDate%>">
    <input type="hidden" id="txtProductFamily" name="txtProductFamily" value="<%=strFamily%>">
    <input type="hidden" id="txtOSListChanged" name="txtOSListChanged" value="">
    <input type="hidden" id="txtFullOSList" name="txtFullOSList" value="<%=strFullOSList%>">
    <input type="hidden" id="txtBrandsLoaded" name="txtBrandsLoaded" value="<%=strBrandsLoaded%>">
    <input type="hidden" id="txtReleasesLoaded" name="txtReleasesLoaded" value="<%=strReleasesLoaded%>">
    <input type="hidden" id="txtBrands" name="txtBrands" value="">
    <input type="hidden" id="txtInitialSystemBoardID" name="txtInitialSystemBoardID"
        value="<%=strsystemBoardId%>">
    <input type="hidden" id="txtInitialMachinePnPID" name="txtInitialMachinePnPID" value="<%=strMachinePnPID%>">
    <input type="hidden" id="txtInitialAffectedProduct" name="txtInitialAffectedProduct" value="<%=iAffectedProduct%>">
    <input type="hidden" id="txtIsSEPM" name="txtIsSEPM" value="<%=strIsSEPM%>">
    </form>
    <%
	set rs = nothing
	cn.Close
	set cn = nothing
	
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function	
	
	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
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
	
	function CleanIDList (strIDList)
		dim IDArray
		dim strID
		dim strOut
		IDArray = split(strIDList,",")
		strOut = ""
		for each strID in IDArray
			if isnumeric(strID) then
				strOut = strOut & "," & trim(clng(strID))
			end if
		next
		if strOut <> "" then
			strOut = mid(strOut,2)
		end if
		CleanIDList = strOut
	end function
	
    %>
    <input type="hidden" id="txtDefaultTab" name="txtDefaultTab" value="<%=request("Tab")%>">
    <select style="display: none" id="cboInactiveProducts">
        <%=strInactiveList%>
    </select>
</body>
</html>
