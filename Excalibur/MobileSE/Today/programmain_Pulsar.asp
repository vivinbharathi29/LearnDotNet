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
    <script id="clientEventHandlersJS" type="text/javascript" language="javascript">
<!--
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

            if ($("#tagPhase").val() != 1) {
                $("#txtProductNameBase").prop("readonly", true);
                $('#txtProductNameBase').css('background-color', '#DEDEDE');
                $('#txtProductNameBase').css('color', '#A9A9A9');
            }

            $(".ModelNumber").on("keyup", function (e) {

                var AllowedCharacters = /^[[0-9a-zA-Z \-]+$/;
                if ($(this).val() != "" && !($(this).val().match(AllowedCharacters))) 
                  {
                    alert("Model Number can only contain alphanumeric and dashes");
                    $(this).val("");
                    return;
                  }
                 // do not allow spaces
                $(this).val(this.value.replace(/^\s+|\s+$/g, ""));

            });

            // hide/show DCR approvers onload base on the checkbox selected
            var chkID = $("input[type=radio][name=chkDCRAutoOpen]:checked").val(); 
                $("tr.DCRApprover").hide();
                $("#TeamRoster" + chkID).show();

            // hide/show DCR approvers on radio button click
            $("input[name$='chkDCRAutoOpen']").click(function () {
                var ID = $(this).val();
                $("tr.DCRApprover").hide();
                $("#TeamRoster" + ID).show();
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

        function AddRTPandEMDate() {
            //ShowPropertiesDialog("../../schedule/TargetRTPandEMUpdateFrame.asp?ID=" + $("#txtID").val(), "Schedule Properties", 800, 275);

            var strResult;
            strResult = window.showModalDialog("../../schedule/TargetRTPandEMUpdateFrame.asp?ID=" + $("#txtID").val(), "", "dialogwidth:" + 800 + "px; dialogheight:" + 275 + "px");
            if (typeof strResult === "undefined") {
                return false;
            } else {
                return true;
            }
        }

        function adjustWidth(percent) {
            return document.documentElement.offsetWidth * (percent / 100);
        }

        function adjustHeight(percent) {
            return document.documentElement.offsetHeight * (percent / 100) + 100;
        }
        
        function ClosePlatformFormFactor(refresh, PlatformID, ProductVersionID) {
            CloseDialog2();
            if (refresh) {
                CloseDialog1();
                try {
                    OpenDialog1("platform.asp?ID=" + PlatformID + "&ProductVersionID=" + ProductVersionID, "Platform - Update", 0, 0, true, true, true);
                }
                catch (e) {
                    if (window.event.srcElement.className == "cell") {
                        var sheight = $(window).height() * (90 / 100);
                        if (sheight < 600 && screen.height < 600)
                            sheight = 600;
                        else
                            sheight = screen.height * (70 / 100);

                        var sWidth = $(window).width() * (50 / 100);
                        if (sWidth < 900)
                            sWidth = 900;

                        window.showModalDialog("platform.asp?ID=" + PlatformID + "&ProductVersionID=" + ProductVersionID, "", "dialogwidth:" + sWidth + "px; dialogheight:" + sheight + "px");
                        window.location.reload();
                    }
                }
            }
        }

        function OpenDialog1(url, title, width, height, resizable, modal, sbuttons, disableScrollBar) {
            if (width == 0)
                width = adjustWidth(98);

            if (height == 0)
                height = adjustHeight(100);

            $("#Dialog1").dialog({
                title: title,
                resizable: resizable,
                width: width,
                minWidth: 400,
                height: height,
                minHeight: 140,
                modal: modal,
                closeOnEscape: true,
                open: function (event, ui) {
                    $("#DialogIframe").attr("src", url);
                    if (disableScrollBar) {
                        $('#Dialog').css('overflow', 'hidden');
                        $("#DialogIframe").css('overflow', 'hidden');
                    }
                },
                close: function () {
                    window.parent.frames["LowerWindow"].cmdSubmit.disabled = false;
                    window.parent.frames["LowerWindow"].cmdEditCancel.disabled = false;
                    window.parent.frames["LowerWindow"].cmdClear.disabled = false;
                    $("#DialogIframe").attr("src", "about:blank");
                }
            });
            window.parent.frames["LowerWindow"].cmdSubmit.disabled = true;
            window.parent.frames["LowerWindow"].cmdEditCancel.disabled = true;
            window.parent.frames["LowerWindow"].cmdClear.disabled = true;
            $('#Dialog1').dialog('open');
        }

        function OpenDialog2(url, title, width, height, resizable, modal, sbuttons, disableScrollBar) {
            if (width == 0)
                width = adjustWidth(98);

            if (height == 0)
                height = adjustHeight(100);

            $("#Dialog2").dialog({
                title: title,
                resizable: resizable,
                width: width,
                minWidth: 400,
                height: height,
                minHeight: 140,
                modal: modal,
                closeOnEscape: true,
                open: function (event, ui) {
                    $("#DialogIframe2").attr("src", url);
                    if (disableScrollBar) {
                        $('#Dialog').css('overflow', 'hidden');
                        $("#DialogIframe2").css('overflow', 'hidden');
                    }
                },
                close: function () {
                    window.parent.frames["LowerWindow"].cmdSubmit.disabled = false;
                    window.parent.frames["LowerWindow"].cmdEditCancel.disabled = false;
                    window.parent.frames["LowerWindow"].cmdClear.disabled = false;
                    $("#DialogIframe2").attr("src", "about:blank");
                }
            });

            $('#Dialog2').dialog('open');
        }

    function ClosePlatFormDetail(DialogID, withRefresh) {
            $("#" + DialogID).dialog("close");
        if (withRefresh)
                $("#PlatformFrame").attr('src', $('#PlatformFrame').attr('src'));

            window.parent.frames["LowerWindow"].cmdSubmit.disabled = false;
            window.parent.frames["LowerWindow"].cmdEditCancel.disabled = false;
            window.parent.frames["LowerWindow"].cmdClear.disabled = false;
            $("#DialogIframe").attr("src", "about:blank");
        }

        function CloseDialog1() {
            $("#Dialog1").dialog("close");            
            $("#DialogIframe").attr("src", "about:blank");
        }

        function CloseDialog2() {
            $("#Dialog2").dialog("close");
            $("#DialogIframe2").attr("src", "about:blank");
        }
        
    var CurrentState;

    function ProcessState() {
        var steptext;

        switch (CurrentState) {
            case "General":
                steptext = "";

                tabGeneral.style.display = "";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";                
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

                window.scrollTo(0, 0);
                break;

            case "Platforms":
                steptext = "";

                tabFiles.style.display = "none";
                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "";

                window.scrollTo(0, 0);
                break;

            case "Files":
                steptext = "";

                tabFiles.style.display = "";
                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

                window.scrollTo(0, 0);
                break;

            case "Access":
                steptext = "";

                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "";
                if (ProgramInput.cboToolsPM.selectedIndex == 0)
                    PMAccessCell.innerHTML = "Project Manager";
                else
                    PMAccessCell.innerHTML = ProgramInput.cboToolsPM.options[ProgramInput.cboToolsPM.selectedIndex].text;
                tabPlatforms.style.display = "none";

                window.scrollTo(0, 0);
                break;

            case "SystemTeam":
                steptext = "";

                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "";
                tabsystemTeamMsg.style.display = "";
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

                window.scrollTo(0, 0);
                break;
            case "OTS":

                steptext = "";

                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabOTS.style.display = "";
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

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

                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

                window.scrollTo(0, 0);
                break;
            case "Approvers":
                steptext = "";

                tabGeneral.style.display = "none";
                tabSystemTeam.style.display = "none";
                tabsystemTeamMsg.style.display = "none";
                tabFiles.style.display = "none";
                tabStatusReport.style.display = "none";
                tabOTS.style.display = "none";
                tabAccess.style.display = "none";
                tabPlatforms.style.display = "none";

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
        ProgramInput.txtIDInformationPath.value = ProgramInput.tagIDInformationPath.value;
        ProgramInput.txtMSPEKSExecutionPath.value = ProgramInput.tagMSPEKSExecutionPath.value;
        LoadingRow.style.display = "none";
        ButtonRow.style.display = "";
        if (txtDefaultTab.value == "SystemTeam")
            CurrentState = "SystemTeam";
        else if (txtDefaultTab.value == "OTS")
            CurrentState = "OTS";
        else if (txtDefaultTab.value == "FilePaths")
            CurrentState = "Files";
        else if (txtDefaultTab.value == "Platforms")
            CurrentState = "Platforms";
        else if (txtDefaultTab.value == "StatusData")
            CurrentState = "StatusReport";
        else
            CurrentState = "General";
        //ProcessState();
        SelectTab(CurrentState);
        self.focus();

        BusinessSegment_onchange();
        BusinessSegment_onchange2();

        $("#cboBSBrandFilter").change(function () {
            BusinessSegment_onchange2();
        });

        // Identify Business Sector and busisiness operation
        BusinessSectorAndOperation($("#cboBusinessSegmentID").val());

        //initialize modal dialog
        modalDialog.load();
    }

    function SelectTab(strStep) {
        var i;
        CurrentState = strStep;
        //Reset all tabs

        document.all("CellGeneralb").style.display = "none";
        document.all("CellGeneral").style.display = "";
        document.all("CellSystemTeamb").style.display = "none";
        document.all("CellSystemTeam").style.display = "";
        document.all("CellAccessb").style.display = "none";
        document.all("CellAccess").style.display = "none";
        document.all("CellStatusReportb").style.display = "none";
        document.all("CellStatusReport").style.display = "";


        if (document.all("CellOTSb") != null) {
            document.all("CellOTSb").style.display = "none";
        }
        if (ProgramInput.txtID.value != "") {
            if (document.all("CellOTS") != null) {
                document.all("CellOTS").style.display = "";
            }
        }

        document.all("CellFilesb").style.display = "none";
        document.all("CellFiles").style.display = "";

        if (document.all("CellPlatformsb") != null) {
            document.all("CellPlatformsb").style.display = "none";
        }
        if (ProgramInput.txtID.value != "") {
            if (document.all("CellPlatforms") != null) {
                document.all("CellPlatforms").style.display = "";
            }
        }
   

        //Highight the selected tab
        document.all("Cell" + strStep).style.display = "none";
        document.all("Cell" + strStep + "b").style.display = "";

        ProcessState();
    }



    function cmdAddFamily_onclick() {
        modalDialog.open({ dialogTitle: 'Add Product Family', dialogURL: 'family.asp', dialogHeight: 250, dialogWidth: 550, dialogResizable: true, dialogDraggable: true });
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

    function cmdODMHWPMAdd_onclick() {
        ChooseEmployee(ProgramInput.cboODMHWPM);
    }

    function cmdHWPCAdd_onclick() {
        ChooseEmployee(ProgramInput.cboHWPC);
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

    function cmdSharedAvMarketing_onclick() {
        ChooseEmployee(ProgramInput.cboSharedAvMarketing);
    }
    
    function cmdSharedAVPC_onclick() {
        ChooseEmployee(ProgramInput.cboSharedAVPC);
    }
    /* LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB */
    function cmdSysEngrProgramCoordinatorAdd_onclick() {
        ChooseEmployee(ProgramInput.cboSysEngrProgramCoordinator);
    }
    /* LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB */

    function cmdProgramBusinessManagerAdd_onclick() {
        ChooseEmployee(ProgramInput.cboProgramBusinessManager);
    }
    function SystemTeamAdd(objectId) {
        var obj = document.getElementById(objectId);
        ChooseEmployee(obj)
    }

    function cmdSCMOwner_onclick() {
        ChooseEmployee(ProgramInput.cboSCMOwner);
    }
    function cmdEngineeringDataManagement_onclick() {
        ChooseEmployee(ProgramInput.cboEngineeringDataManagement);
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
        else if (strID == 6) {
            if (ProgramInput.txtIDInformationPath.value == "")
                alert("ID Information Path not specified.");
            else
                window.open(ProgramInput.txtIDInformationPath.value);
        }
        else if (strID == 7) {
            if (ProgramInput.txtMSPEKSExecutionPath.value == "")
                alert("MSPEKS(Execution) Path not specified.");
            else
                window.open(ProgramInput.txtMSPEKSExecutionPath.value);
        }
    }




    function EnterSeries() {
        alert("This function is under development.");
    }

    function BrandCheck_onclick(ID, PVID) {
        var Result = 0;
        var strTemp = "";
        var removedItem = "";

        if (ProgramInput.txtID.value == "") {
            if (event.srcElement.checked) {
                document.all("DivSeries" + ID).style.display = "";
                ProgramInput.txtBrandsAdded.value = ProgramInput.txtBrandsAdded.value + ',' + ID;
            }
            else { 
                document.all("DivSeries" + ID).style.display = "none";
                removedItem = ID;
            }
           
        }
        else {
            if (!event.srcElement.checked) {

                //if the brand has published SCM, do not allow user to delete the brand
                if (document.getElementById("txtSCMEnabled" + ID).value != "1") {
                    alert("The brand has already been published in an SCM and cannot be deleted");
                    event.srcElement.checked = true;
                    return false;
                }

                ProgramInput.txtProductName.value = ProgramInput.txtProductNameBase.value;

                Result = window.showModalDialog("BrandDeleteWarning.asp?ProductName=" + ProgramInput.txtProductName.value + "&BrandName=" + event.srcElement.BrandName + "&BrandID=" + ID + "&PVID=" + PVID, "", "dialogWidth:700px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
                if (Result == "1") {
                    document.all("DivSeries" + ID).style.display = "none";
                    removedItem = ID;
                }
                else {
                    event.srcElement.checked = true;
                    document.all("DivSeries" + ID).style.display = "";
                }

                if (removedItem != "") {
                    var brands = ProgramInput.txtBrandsAdded.value;
                    if (brands[0] == ',')
                        brands = brands.substring(1);
                    var temp = "";
                    $.each(brands.split(','), function (index, value) {
                        if (value != "" && removedItem != value) {
                            temp = temp + ',' + value;
            }
                    });
                    if (temp[0] == ',')
                        temp = temp.substring(1);
                    ProgramInput.txtBrandsAdded.value = temp;
                }
            }
            else {
                document.all("DivSeries" + ID).style.display = "";
                ProgramInput.txtBrandsAdded.value = ProgramInput.txtBrandsAdded.value + ',' + ID;
            }

        }

    }

    function cboDevCenter_onchange() {
        if (ProgramInput.cboReleaseTeam.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
            ProgramInput.cboReleaseTeam.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;
        if (ProgramInput.cboPreinstall.selectedIndex == 0 && (ProgramInput.cboDevCenter.selectedIndex == 1 || ProgramInput.cboDevCenter.selectedIndex == 2))
            ProgramInput.cboPreinstall.selectedIndex = ProgramInput.cboDevCenter.selectedIndex;
    }

    function ChooseEmployee(myControl) {
        modalDialog.open({ dialogTitle: 'Select Employee', dialogURL: 'ChooseEmployee.asp', dialogHeight: 200, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('1', 'employee_type');
        globalVariable.save(myControl.id, 'employee_dropdown');
    }


    function ChooseEmployee2() {
        modalDialog.open({ dialogTitle: 'Select Employee', dialogURL: 'ChooseEmployee.asp', dialogHeight: 200, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('2', 'employee_type');
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
            default:break;
        }
    }

    function RemoveToolAccess(strID) {
        document.all("ToolAccessRow" + strID).style.display = "none";
        document.getElementById("chkToolAccessID" + strID).checked = false;

    }

    function cboType_onchange() {

        BrandRow.style.display = "";
        ApproverRow.style.display = "";
        ActivitiesRow.style.display = "";
        CommodityRow.style.display = "";
        MDARow.style.display = "";
        ToolPMRow.style.display = "none";
        ToolPMRow2.style.display = "none";
        ReleaseRow.style.display = "";
        PreinstallRow.style.display = "";
        DevCenterRow.style.display = "";
        DistributionRow.style.display = "";
        NotificationRow.style.display = "none";

        tdReleaseTitle.style.display = "";
        tdReleaseText.style.display = "";
    
        SelectTab(CurrentState);
    }

    function cboFamily_onchange() {
        ProgramInput.txtProductFamily.value = ProgramInput.cboFamily.options[ProgramInput.cboFamily.selectedIndex].text;
    }

    function cboFamily_selectcurrent(current) {
        $("#cboFamily option[value='" + current + "']").prop('selected', 'selected');
    }

    /* LY BEGINNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM */
    function cboProductLine_onchange() {
        ProgramInput.txtProductLine.value = ProgramInput.cboProductLine.options[ProgramInput.cboProductLine.selectedIndex].text;
    }
    /* LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM */

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
        modalDialog.open({ dialogTitle: 'Component Core Team', dialogURL: 'ChooseComponentCoreTeam.asp?OTSComponentID=' + strOTSComponentID + '&RoleID=3&CoreTeamID=' + strCoreTeamID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('1', 'role_type');
        globalVariable.save(strID, 'role_id');
        globalVariable.save(strOTSComponentID, 'role_otscomponentid');
    }

    function EditOTSPM(strID, strOwnerID) {
        modalDialog.open({ dialogTitle: 'Component Owner', dialogURL: 'ChooseComponentOwner.asp?ID=' + strID + '&RoleID=1&OwnerID=' + strOwnerID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('2', 'role_type');
        globalVariable.save(strID, 'role_id');
    }

    function EditOTSDeveloper(strID, strOwnerID) {
        modalDialog.open({ dialogTitle: 'Component Owner', dialogURL: 'ChooseComponentOwner.asp?ID=' + strID + '&RoleID=2&OwnerID=' + strOwnerID, dialogHeight: 250, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('3', 'role_type');
        globalVariable.save(strID, 'role_id');
    }

    function EditComponentResults() {
        var iType;
        var strID;
        var strOTSComponentID;
        var strResult;

        strResult = modalDialog.getArgument('role_query_array');
        strResult = JSON.parse(strResult);

        iType = globalVariable.get('role_type');

        switch(iType) {
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
        modalDialog.open({ dialogTitle: 'System Board', dialogURL: 'ProductID.asp?TypeID=1&IDList=' + encodeURI(ProgramInput.txtSystemBoardComments.value.replace("\"", "%22")), dialogHeight: 500, dialogWidth: 450, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('1', 'id_type');
    }

    function SelectMachinePNPID() {
        modalDialog.open({ dialogTitle: 'PnP', dialogURL: 'ProductID.asp?TypeID=2&IDList=' + ProgramInput.txtMachinePNPComments.value.replace("\"", "%22"), dialogHeight: 500, dialogWidth: 700, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('2', 'id_type');
    }

    function GetProductIDResult() {
        ResultArray = modalDialog.getArgument('product_id_array');
        ResultArray = JSON.parse(ResultArray);

        iTypeID = globalVariable.get('id_type');

        switch(iTypeID) {
            case "1":
                if (typeof (ResultArray) != "undefined") {
                    ProgramInput.txtSystemBoardID.value = ResultArray[0];
                    ProgramInput.txtSystemBoardComments.value = ResultArray[1];
                }
                break;
            case "2":
                if (typeof (ResultArray) != "undefined") {
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
        if ($("#tagPhase").val() != "" && $("#tagPhase").val() != 1 && $("#cboPhase").val() == 1)
        {
            alert("The Product Phase can not be changed back to Definition due to problems with Sudden Impact.")
            $("#cboPhase").prop('selectedIndex', $("#tagPhase").val()-1);
        }
       

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
        var currentBSegmentID = $("#cboBSBrandFilter").val();

        if (ProgramInput.txtBrandsLoaded.value != "")
            strIDs = "," + ProgramInput.txtBrandsLoaded.value;
        if (ProgramInput.txtBrandsAdded.value != "")
            strIDs = "," + ProgramInput.txtBrandsAdded.value;

        if (strIDs != "")
            strIDs = strIDs.substring(1);

        modalDialog.open({ dialogTitle: 'Brand', dialogURL: 'BrandUpdate_Pulsar.asp?ProductID=' + ProgramInput.txtID.value + '&BrandID=' + strID + '&ExcludeIDList=' + strIDs + '&BSegmentID=' + currentBSegmentID + '', dialogHeight: 300, dialogWidth: 450, dialogResizable: false, dialogDraggable: true });
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

            ProgramInput.txtBrandsAdded.value = ProgramInput.txtBrandsAdded.value + ',' + strResult;

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

    function SelectedSitesResult(strResult){
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

    function cboBusinessSegment_onchange() {
        var newBusinessSegmentId = $("#cboBusinessSegmentID").val();
        var oldBusinessSegmentId = $("#hdnBusinessSegmentID").val();
        var response;
        response = true;

       if  ($("#cboSEPM").val() > 0 ||  
        $("#cboPM").val() > 0 ||  
        $("#cboTDCCM").val() > 0 ||  
        $("#cboPDE").val() > 0 ||  
        $("#cboFactoryEngineer").val() > 0 || 
        $("#cboAccessoryPM").val() > 0 ||  
        $("#cboComMarketing").val() > 0 ||  
        $("#cboPlatformDevelopment").val() > 0 ||  
        $("#cboSupplyChain").val() > 0 ||  
        $("#cboService").val() > 0 ||  
        $("#cboQuality").val() > 0 ||  
        $("#cboSM").val() > 0 || 
        $("#cboToolsPM").val() > 0 ||  
        $("#cboPINPM").val() > 0 ||  
        $("#cboSEPE").val() > 0 ||  
        $("#cboSETestLead").val() > 0 || 
        $("#cboSETest").val() > 0 ||  
        $("#cboODMTestLead").val() > 0 ||  
        $("#cboWWANTestLead").val() > 0 ||  
        $("#cboBIOSLead").val() > 0 ||  
        $("#cboProcessorPM").val() > 0 ||  
        $("#cboEngineeringDataManagement").val() > 0 || 
        $("#cboSCMOwner").val() > 0 ||  
        $("#cboHWPC").val() > 0 ||  
        $("#cboODMPIMPM").val() > 0 ||  
        $("#cboODMHWPM").val() > 0 ||  
        $("#cboPlanningPM").val() > 0 ||  
        $("#cboProcurementPM").val() > 0 || 
        $("#cboODMSEPM").val() > 0 ||  
        $("#cboDocPM").val() > 0 ||  
        $("#cboSpdm").val() > 0 ||  
        $("#cboGplm").val() > 0 || 
        $("#cboMarketingOps").val() > 0 ||  
        $("#cboPC").val() > 0 ||  
        $("#cboProgramBusinessManager").val() > 0 ||  
        $("#cboSysEngrProgramCoordinator").val() > 0 ||  
        $("#cboSustainingSEPM").val() > 0 ||  
        $("#cboDKC").val() > 0 ||  
        $("#cboCommHWPM").val() > 0 ||  
        $("#cboSharedAVPC").val() > 0 ||  
        $("#cboSharedAvMarketing").val() > 0 ||  
        $("#cboGraphicsControllerPM").val() > 0 ||  
        $("#cboVideoMemoryPM").val() > 0 )
        { //if atleast one of the system team drop down has been selected
            response = confirm("Changing the business segment will remove all existing System Team Roster selections.");           
        }
        if (response)
        {
            ReleaseLink.innerText = "Add New Release";
            $("#txtProductRelease").val("");
            $("#txtProductReleaseIDs").val("");
            // change brand filter only in creating New product
            if (ProgramInput.txtID.value == "") {
                $("#hdnBusinessSegmentID").val(newBusinessSegmentId);
                $("#cboBSBrandFilter").val(newBusinessSegmentId);
                BusinessSegment_onchange2();
            }
            // Identify Business Sector and busisiness operation
            BusinessSectorAndOperation(newBusinessSegmentId);
            LoadSystemTeamDropDown(newBusinessSegmentId);
        }
        else
        {
            $("#cboBusinessSegmentID").val(oldBusinessSegmentId);
        }
    }

    function AddRelease(ID) {
        var productTypeId = $("#cboType").val();
        var businessSegmentId = $("#cboBusinessSegmentID").val();

        if (productTypeId == 0 || businessSegmentId == "") {
            alert("Please select Product Type and Business Segment to add release")
            return;
        }

        var sWidth = $(window).width() * 70 / 100;
        modalDialog.open({ dialogTitle: 'Select Product (Release) / Lead Product (Release) for', dialogURL: 'ProductRelease.asp?ID=' + ID + '&ProductTypeID=' + productTypeId + '&ProductBusinessSegmentID=' + businessSegmentId + '&isClone=' + $("#hdnIsClone").val(), dialogHeight: 600, dialogWidth: sWidth, dialogResizable: false, dialogDraggable: true });
    }

    function AddReleaseResult(strNames, strIDs) {
        if (typeof strNames === "undefined") {
            return;
        }

        if (strNames == "")
            ReleaseLink.innerText = "Add New Release";
        else {
            ReleaseLink.innerText = strNames;
        }

        $("#txtProductRelease").val(strNames);
        $("#txtProductReleaseIDs").val(strIDs);
    }

    function AddPlatform() {
        window.showModalDialog("platform.asp");
    }

    function SetCreateSimpleAvStatus(PulsarProduct) {
        if (PulsarProduct == 1) {
            ProgramInput.optCreateSimpleAvTypeAuto.disabled = false;
            ProgramInput.optCreateSimpleAvTypeManual.disabled = false;
            if (ProgramInput.optCreateSimpleAvTypeAuto.checked == false && ProgramInput.optCreateSimpleAvTypeManual.checked == false) {
                ProgramInput.optCreateSimpleAvTypeAuto.checked = true;
                ProgramInput.optCreateSimpleAvTypeManual.checked = false;
            }
        } else {
            ProgramInput.optCreateSimpleAvTypeAuto.disabled = true;
            ProgramInput.optCreateSimpleAvTypeManual.disabled = true;
            ProgramInput.optCreateSimpleAvTypeAuto.checked = false;
            ProgramInput.optCreateSimpleAvTypeManual.checked = false;
        }
    }
    function EliminateFirstSapce() {
        var sName = document.getElementById("txtProductNameBase").value;
        sName = (sName.replace(/^\W+/, '')).replace(/\W+$/, '');

        if (sName == "") {
            document.getElementById("txtProductNameBase").value = "";
            return false;
        }

        if (document.getElementById("txtProductNameBase").value.indexOf("<") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace('<', '');
            return false;
        }
        if (document.getElementById("txtProductNameBase").value.indexOf(">") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace('>', '');
            return false;
        }
        if (document.getElementById("txtProductNameBase").value.indexOf(";") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace(';', '');
            return false;
        }
        if (document.getElementById("txtProductNameBase").value.indexOf("\"") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace('\"', '');
            return false;
        }

        if (document.getElementById("txtProductNameBase").value.indexOf("&") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace('&', '');
            return false;
        }
        if (document.getElementById("txtProductNameBase").value.indexOf("\'") > -1) {
            document.getElementById("txtProductNameBase").value = sName.replace('\'', '');
            return false;
        }
    }

    function LoadSystemTeamDropDown(BusinessSegmentID)
    {
        var ProductVersionID = 0;
        if (ProgramInput.txtID.value != "")
            ProductVersionID = ProgramInput.txtID.value;
        var strPartner;
        strPartner = $("#hdnProductPartner").val();
        $.ajax({
            url: "/Pulsar/Product/GetSystemTeamDropDown?ProductVersionID= " + ProductVersionID + "&BusinessSegmentID=" + BusinessSegmentID,
            method: "GET",
            async: false,
            success: function (returnData) {
                var data = jQuery.parseJSON(returnData);
                $("#cboSEPM").html('<option value="0"></option>');
                $("#cboPM").html('<option value="0"></option>');
                $("#cboTDCCM").html('<option value="0"></option>');
                $("#cboPDE").html('<option value="0"></option>');
                $("#cboFactoryEngineer").html('<option value="0"></option>');
                $("#cboAccessoryPM").html('<option value="0"></option>');
                $("#cboComMarketing").html('<option value="0"></option>');
                $("#cboPlatformDevelopment").html('<option value="0"></option>');
                $("#cboSupplyChain").html('<option value="0"></option>');
                $("#cboService").html('<option value="0"></option>');
                $("#cboQuality").html('<option value="0"></option>');
                $("#cboSM").html('<option value="0"></option>');
                $("#cboToolsPM").html('<option value="0"></option>');
                $("#cboPINPM").html('<option value="0"></option>');
                $("#cboSEPE").html('<option value="0"></option>');
                $("#cboSETestLead").html('<option value="0"></option>');
                $("#cboSETest").html('<option value="0"></option>');
                $("#cboODMTestLead").html('<option value="0"></option>');
                $("#cboWWANTestLead").html('<option value="0"></option>');
                $("#cboBIOSLead").html('<option value="0"></option>');
                $("#cboProcessorPM").html('<option value="0"></option>');
                $("#cboEngineeringDataManagement").html('<option value="0"></option>');
                $("#cboSCMOwner").html('<option value="0"></option>');
                $("#cboHWPC").html('<option value="0"></option>');
                $("#cboODMPIMPM").html('<option value="0"></option>');
                $("#cboODMHWPM").html('<option value="0"></option>');
                $("#cboPlanningPM").html('<option value="0"></option>');
                $("#cboProcurementPM").html('<option value="0"></option>');
                $("#cboODMSEPM").html('<option value="0"></option>');
                $("#cboDocPM").html('<option value="0"></option>');
                $("#cboSpdm").html('<option value="0"></option>');
                $("#cboGplm").html('<option value="0"></option>');
                $("#cboMarketingOps").html('<option value="0"></option>');
                $("#cboPC").html('<option value="0"></option>');
                $("#cboProgramBusinessManager").html('<option value="0"></option>');
                $("#cboSysEngrProgramCoordinator").html('<option value="0"></option>');
                $("#cboSustainingSEPM").html('<option value="0"></option>');
                $("#cboDKC").html('<option value="0"></option>');
                $("#cboCommHWPM").html('<option value="0"></option>');
                $("#cboSharedAVPC").html('<option value="0"></option>');
                $("#cboSharedAvMarketing").html('<option value="0"></option>');
                $("#cboGraphicsControllerPM").html('<option value="0"></option>');
                $("#cboVideoMemoryPM").html('<option value="0"></option>');


                $(document).ready(function () {
                    for (var i = 0; i <= data.dropdownresults.length - 1; i++) {
                        var item = data.dropdownresults[i];
                        if (item.Role == "SE PM") {
                            if (item.bUsed == 1)
                                $("#cboSEPM").append('<option value="' + item.ID + '" selected="selected">' + item.Name + '</option>');
                            else
                                $("#cboSEPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');                                                                          
                        }
                        if (item.Role == "POPM") {
                            if (item.bUsed == 1)
                                $("#cboPM").append('<OPTION value="' + item.ID + '" selected="selected">' + item.Name + '</OPTION>');
                            else 
                                $("#cboPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "TDCCM") {
                            if (item.bUsed == 1)
                                $("#cboTDCCM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else 
                                $("#cboTDCCM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Commodity PM") {
                            if (item.bUsed == 1)
                                $("#cboPDE").append('<option value="' + item.ID + '" selected="selected">' + item.Name + '</option>');
                            else
                                $("#cboPDE").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Factory Engineer") {
                            if (item.bUsed == 1)
                                $("#cboFactoryEngineer").append('<OPTION selected="selected" value="' + item.ID + '">' + item.Name + '</OPTION>');
                            else 
                                $("#cboFactoryEngineer").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Accessory PM") {
                            if (item.bUsed == 1)
                                $("#cboAccessoryPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboAccessoryPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Commercial Marketing") {
                            if (item.bUsed == 1)
                                $("#cboComMarketing").append('<option value="' + item.ID + '" selected="selected">' + item.Name + '</option>');
                            else
                                $("#cboComMarketing").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Platform Development") {
                            if (item.bUsed == 1)
                                $("#cboPlatformDevelopment").append('<OPTION selected="selected" value="' + item.ID + '">' + item.Name + '</OPTION>');
                            else 
                                $("#cboPlatformDevelopment").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Supply Chain") {
                            if (item.bUsed == 1)
                                $("#cboSupplyChain").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSupplyChain").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        
                        if (item.Role == "Service") {
                            if (item.bUsed == 1)
                                $("#cboService").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboService").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                       
                        if (item.Role == "Quality") {
                            if (item.bUsed == 1)
                                $("#cboQuality").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboQuality").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                       
                        if (item.Role == "System Manager") {
                            if (item.bUsed == 1)
                            {
                                $("#cboSM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                                $("#cboToolsPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            }
                            else
                            {
                                $("#cboSM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                                $("#cboToolsPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                            }
                        }
                        
                        if (item.Role == "SE PE") {
                           if (item.bUsed == 1)
                                $("#cboSEPE").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                           else
                                $("#cboSEPE").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        
                        if (item.Role == "Preinstall PM") {
                            if (item.bUsed == 1)
                                $("#cboPINPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboPINPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        
                        if (item.Role == "SE Test Lead") {
                            if (item.bUsed == 1)
                                $("#cboSETestLead").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else if (strPartner = "" || strPartner == "0" || item.PartnerID  == 1 || item.PartnerID == strPartner)                             
                                $("#cboSETestLead").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                           
                        }
                       
                        if (item.Role == "SE Test") {
                            if (item.bUsed == 1)
                                $("#cboSETest").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSETest").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                       
                        if (item.Role == "ODM Test Lead") {
                            if (item.bUsed == 1)
                                $("#cboODMTestLead").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboODMTestLead").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        
                        if (item.Role == "WWAN Test Lead") {
                            if (item.bUsed == 1)
                                $("#cboWWANTestLead").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else if (strPartner = "" || strPartner == "0" || item.PartnerID  == 1 || item.PartnerID == strPartner) {
                                $("#cboWWANTestLead").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                            }
                        }
                        
                        if (item.Role == "BIOS Lead") {
                            if (item.bUsed == 1)
                                $("#cboBIOSLead").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboBIOSLead").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                       
                        if (item.Role == "Chipset/Processor PM") {
                            if (item.bUsed == 1)
                                $("#cboProcessorPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboProcessorPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "Video Memory PM") {
                            if (item.bUsed == 1)
                                $("#cboVideoMemoryPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboVideoMemoryPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Graphics Control PM" ) { 
                            if (item.bUsed == 1)
                                $("#cboGraphicsControllerPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboGraphicsControllerPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "SharedAVMarketing") {
                            if (item.bUsed == 1)
                                $("#cboSharedAvMarketing").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSharedAvMarketing").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "SharedAVPC") {
                            if (item.bUsed == 1)
                                $("#cboSharedAVPC").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSharedAVPC").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Comm HW PM") {
                            if (item.bUsed == 1)
                                $("#cboCommHWPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboCommHWPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                        if (item.Role == "Doc Kit Coordinator") {
                            if (item.bUsed == 1)
                                $("#cboDKC").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboDKC").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }		
                        if (item.Role == "Sustaining SE PM") {
                            if (item.bUsed == 1)
                                $("#cboSustainingSEPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboSustainingSEPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role =="SysEngrPC") { 
                            if (item.bUsed == 1)
                                $("#cboSysEngrProgramCoordinator").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSysEngrProgramCoordinator").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "ProgramBusiManager") {  
                            if (item.bUsed == 1)
                                $("#cboProgramBusinessManager").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboProgramBusinessManager").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "Program Coordinator") {
                            if (item.bUsed == 1)
                                $("#cboPC").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboPC").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "PDM Team") {
                            if (item.bUsed == 1)
                                $("#cboMarketingOps").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboMarketingOps").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "GPLM") {
                            if (item.bUsed == 1)
                                $("#cboGplm").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboGplm").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "SPDM") {
                            if (item.bUsed == 1)
                                $("#cboSpdm").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboSpdm").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }                       

                        if (item.Role == "Doc PM") {
                            if (item.bUsed == 1)
                                $("#cboDocPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboDocPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "ODMSEPM") {
                            if (item.bUsed == 1)
                                $("#cboODMSEPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboODMSEPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "ProcurementPM") {
                            if (item.bUsed == 1)
                                $("#cboProcurementPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboProcurementPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "PlanningPM") {
                            if (item.bUsed == 1)
                                $("#cboPlanningPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboPlanningPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "ODMHWPM") {
                            if (item.bUsed == 1)
                                $("#cboODMHWPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboODMHWPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "ODMPIMPM") {
                            if (item.bUsed == 1)
                                $("#cboODMPIMPM").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboODMPIMPM").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "HWPC") {
                            if (item.bUsed == 1)
                                $("#cboHWPC").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                            $("#cboHWPC").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "SCMOwnerID") {
                            if (item.bUsed == 1)
                                $("#cboSCMOwner").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboSCMOwner").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }

                        if (item.Role == "EngDataManagementID") { 
                            if (item.bUsed == 1)
                                $("#cboEngineeringDataManagement").append('<option selected="selected" value="' + item.ID + '">' + item.Name + '</option>');
                            else
                                $("#cboEngineeringDataManagement").append('<option value="' + item.ID + '">' + item.Name + '</option>');
                        }
                    }

                });
            },
            cache: false
        });
    }

    function BusinessSegment_onchange() {
        var plDropdown = $("#cboProductLine");
        var businessSegmentId = $("#cboBusinessSegmentID").val();
        var currentProductLine = $("#cboProductLine").val();

        // change brand filter only in creating New product
        if (ProgramInput.txtID.value == "") {
            $("#cboBSBrandFilter").val(businessSegmentId);
            BusinessSegment_onchange2();
        }
       
        // Identify Business Sector and busisiness operation
        BusinessSectorAndOperation(businessSegmentId);
        LoadSystemTeamDropDown(businessSegmentId);

        if (businessSegmentId > 0) {
            $.ajax({
                //url: "/Pulsar/Product/GetProductLines?BusinessSegmentId=" + businessSegmentId,
                url: "/Pulsar/Product/GetProductLines?BusinessSegmentId=0",
                method: "GET",
                success: function (returnData) {
                    var data = jQuery.parseJSON(returnData);
                    $(plDropdown).html('<option value=""></option>');
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
        }
    }

    function BusinessSegment_onchange2() {
        var plDropdown = $("#cboProductLine");
        var brandTableRows = $("#TableBrand tbody tr");
        var businessSegmentId = $("#cboBSBrandFilter").val();

        if (businessSegmentId > 0) {
            var brands = ProgramInput.txtBrandsLoaded.value + ProgramInput.txtBrandsAdded.value;
            if (brands[0] == ',')
                brands = brands.substring(1);

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
                                break;
                            }
                        }
                        if (!bdisplay && brands != "") {
                            $.each(brands.split(','), function (index, value) {
                                itemID = "Brand" + value + ",";
                                if (rowID.indexOf(itemID) > -1) {
                                    bdisplay = true;
                                }
                            });

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
        else {
            $.ajax({
                url: "/Pulsar/Product/GetBrands?BusinessSegmentId=0",
                method: "GET",
                success: function (returnData) {
                    var data = jQuery.parseJSON(returnData);
                    var rows = $(brandTableRows).each(function (rowNum, val) {
                        var row = $(this);
                        $(row).show();
                    });
                },
                cache: false
            });
        }
    }

    // Identify Business Sector and busisiness operation then call to show or hide system team roles
    function BusinessSectorAndOperation(businessSegmentId) {
        var IsDesktop = 0;
        var IsCommercial = 0;
        var arrBSs = $("#hdnBusinessSegmentList").val().split(';');
        for (var i = 0; i < arrBSs.length - 1 && IsDesktop == 0; i++) {
            var arrBS = arrBSs[i].split(",");
            if (arrBS[0] == businessSegmentId) {
                if (arrBS[1] == 1)
                    IsDesktop = 1;
                if (arrBS[2] == 1)
                    IsCommercial = 1;
            }
        }
        ProgramInput.hdnIsDesktop.value = IsDesktop == 0 ? "" : "YES";
        ProgramInput.hdnIsCommercial.value = IsCommercial == 0 ? "" : "YES";
        HideandShow(IsCommercial, IsDesktop);
        if (ProgramInput.txtID.value == "") ChangeDefaultDCROwner_onSegmentChange(IsCommercial);
    }

    // show or hide system team roles
    function HideandShow(IsCommercial, IsDesktop) {
        if (IsDesktop == 1) {
            lblSCMOwner.style.display = "inline";
            lblEngineeringDataManagement.style.display = "inline";
            lblConfigurationManager.style.display = "none";
            lblProgramOfficeManager.style.display = "none";
        }
        else {
            lblSCMOwner.style.display = "none";
            lblEngineeringDataManagement.style.display = "none";

            lblConfigurationManager.style.display = IsCommercial == 1 ? "inline" : "none";
            lblProgramOfficeManager.style.display = IsCommercial == 1 ? "none" : "inline";
        }
    }

    function ChangeDefaultDCROwner_onSegmentChange(IsCommercial) {
        if (IsCommercial) {
            ProgramInput.cboDCRDefaultOwner.selectedIndex = 0; //automatically select Configuration Manager
        } else {
            ProgramInput.cboDCRDefaultOwner.selectedIndex = 1; //automatically select Program Office Manager
        }
    }

    //*****************************************************************
    //Description:  Add Existing Base Unit Group Modal Dialog
    //Function:     AddExistingBaseUnitGroup();
    //Modified:     Harris, Valerie (10/11/2016) 
    //*****************************************************************
    function AddExistingBaseUnitGroup(ID) {
        var url = "/IPulsar/SCM/AddExistingPlatforms.aspx?PVID=" + ID;
        modalDialog.open({ dialogTitle: 'Add Existing Base Unit Group', dialogURL: '' + url + '', dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: false, dialogDraggable: true });
    }

    //*****************************************************************
    //Description:  View Base Units
    //Function:     ViewPlatformBaseUnits();
    //Modified:     Harris, Valerie (10/12/2016) 
    //*****************************************************************
    function ViewPlatformBaseUnits(ID, ProductVersionID, PlatformName) {
        var url = "platformBaseUnitList.asp?ID=" + ID + "&ProductVersionID=" + ProductVersionID + "&PlatformName=" + PlatformName;
        modalDialog.open({ dialogTitle: 'Base Units', dialogURL: '' + url + '', dialogHeight: GetWindowSize('height'), dialogWidth: GetWindowSize('width'), dialogResizable: false, dialogDraggable: true });
    }

    function ViewPlatformBaseUnits_return() {
        var oIFrame = document.getElementById('PlatformFrame');
        oIFrame.contentWindow.location.reload(true);

    }
    //*****************************************************************
    //Description:  Refresh Platform List within Iframe
    //Function:     RefreshPlatformList();
    //Modified:     Harris, Valerie (10/12/2016) 
    //*****************************************************************
    function RefreshPlatformList(returnValue) {
        if (returnValue === 1) {
            //refresh platform list within iframe
            var oIFrame = document.getElementById('PlatformFrame');
            oIFrame.contentWindow.location.reload(true);
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

    function getPageLocation()
    {
        return "ProductProperties";
    }
    //-->
    </script>
</head>
<style type="text/css">
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
<body onload="return window_onload();">
    <font face="verdana">
        <%
   	dim cn
	dim rs
	dim cm
	dim p
	dim CnString
	dim strFamily
    dim isClone
	dim strProductLine	
	dim strSEPM
	dim strPM
	dim strSM
	dim strVersion
    dim strProductRelease
    dim strProductReleaseIDs
    dim strProductNameBase
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
	dim strProductLineID
	dim strFamilyID
	dim CheckEmail
	dim CheckReports
    dim StableConsistent
	dim strDivision
	dim strProductStatus
	dim strBaseUnit
	dim strCurrentROM
	dim strCurrentWebROM
	dim strOSSupport
	dim strImagePO
	dim strImageChanges
    dim strSystemboardIDs
	dim strSystemboardComments
	dim strMachinePNPID
	dim strMachinePNPComments
	dim strCommonimages
	dim strCertificationStatus
	dim strSWQAStatus
	dim strPlatformStatus
    dim strBusinessSegmentID
	dim strBusinessSegment
    dim strBusinessSegmentName
	dim strDCRAutoOpen
    dim strDCRNoOdm
	dim strDcrToCM
	dim strDcrToList
	strDCRAutoOpen = ""
    strDCRNoOdm = ""
	strDcrToCm = ""
	strDcrToList = ""
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
    dim hdnFollowMarketingName
	dim CheckDCR
	dim EnableDCR
	dim CheckMDA
	dim strChecked
	dim strBrandsLoaded
	dim strReleasesLoaded
	dim strPDDPath
	dim strSCMPath
	dim strAccessoryPath
	dim strIDInformationPath
	dim strMSPEKSExecutionPath
	dim strSTLPath
	dim strProgramMatrixPath
	dim strReferenceList
    dim strCompleteProductList
	dim strReferenceID
	dim strDevCenter
	dim strRegulatoryModel
	dim strServiceTag
	dim strBIOSBranding
	dim strPCID
	dim strPCList
	dim strMarketingOpsID
	dim strBIOSLeadList
	dim strVideoMemoryPMList
	dim strGraphicsControllerPMList
	dim strProcessorPMList	
	dim strDKCList
	dim strDCRApproverList
	dim DisplayToolsProject
	dim DisplayToolsProject2
	dim strToolAccessIDList
    dim strGplmList
	dim strSpdmList
    dim strSharedAVPCList
    dim strSharedAVMarketingList
    dim strSharedAVMarketingPM
    dim strSharedAVPCPM
    dim strIsDesktop
    dim strIsCommercial
    dim strBusinessSegmentList
    dim strSCMOwnerID
    dim strSCMOwnerList
    dim strEngineeringDataManagementList
    dim strEngineeringDataManagementID
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
    dim FusionRequirements
    dim FusionReqChecked
    dim FusionLegacyReqChecked
    dim strProductName
    dim CreateSimpleAvTypeAuto
    dim CreateSimpleAvTypeAutoChecked
    dim CreateSimpleAvTypeManualChecked
    dim strFactoryId
	dim strIsPlatformRTM
    dim strBrandsChecked
    dim strBrandsAdded
    dim strWWAN : strWWAN = ""
    dim strHWStatusDisplay : strHWStatusDisplay = ""
	dim strHWStatus : strHWStatus = "" 
	dim strEditServiceEOLDate : strEditServiceEOLDate = "disabled"
    dim sTeamRosterApprovers
    dim sInitialDCR
    dim followMarketingName: followMarketingName = 1
    dim showBrandsRow: showBrandsRow="none"
    dim chkMKT: chkMKT ="checked"

    strIsPlatformRTM = ""
	strReqMilestones = ""
    strIsPlatformRTM = ""
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
    strBrandsAdded = ""
	
    strProductRelease = ""
    strProductReleaseIDs = ""
    
    strIsDesktop = ""
    strIsCommercial = ""
    strBusinessSegmentList = ""
    strSCMOwnerID = ""
    strSCMOwnerList = ""
    strEngineeringDataManagementID = ""
    strEngineeringDataManagementList = ""

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
    dim intCMProductCount 
    dim CurrentUserSysAdmin
	dim isCMPermission: isCMPermission = false

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
        intCMProductCount = rs("CMProductCount") 
        CurrentUserSysAdmin = rs("SystemAdmin")
	end if
	rs.Close
    if (intCMProductCount > 0 or CurrentUserSysAdmin) then
        isCMPermission = true
    end if

	if request("ID") = "" then
        isClone = false
		Response.write "<H3>Add New Product (Pulsar)</H3>" 
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
        strProductReleaseIDs = ""
        strProductNameBase = ""
		strDescription = ""
		strObjective = ""
		strOTSName = ""
		strApprover = ""
		strEmailActive = ""
		strProductLineID = ""
		strFamilyID = ""
		CheckEmail = "checked"
		CheckReports = "checked"
        StableConsistent = ""
		strDivision = ""
        strBusinessSegmentID = 0
		strBusinessSegment = ""
        strBusinessSegmentName = ""
        strProductName = ""
		strType = ""
	    strRCTOSites = ""
		strProductStatus = "1"
		strBaseUnit = ""
		strCurrentROM = ""
		strCurrentWebROM = ""
		strOSSupport = ""
		strImagePO = ""
		strImageChanges = ""
        strSystemboardIDs = ""
		strSystemboardComments= ""
		strMachinePNPID = ""
        FusionRequirements = "true"
        CreateSimpleAvTypeAuto = "true"
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
		strIDInformationPath = ""
		strMSPEKSExecutionPath = ""
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
        strODMHWPM = ""
        strODMPIMPM = ""
        strHWPC = ""
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
        strProgramBusinessManagerID = ""
		strPinCutoffValue = ""
		strRegulatoryModel = ""
		strServiceTag = ""
	    strBIOSBranding = ""
		'strPOPMLabel = "Configuration&nbsp;Manager:"
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
        CheckFusion = ""
		CheckDeliverables = "checked"
		CheckImages = "checked"
		CheckDCR = ""
		EnableDCR=" disabled "
		
		CheckMDA = "checked"
	  strFactoryId = "0"

        FusionReqChecked = "checked"
        FusionLegacyReqChecked = ""

        strSharedAVMarketingPM = ""
        strSharedAVPCPM = ""
        sInitialDCR = "checked"
	else
        if Request("Clone") = "1" then isClone = true
		Response.write "<H3>Product Properties (Pulsar)</H3>" 

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion_Pulsar"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
           
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
        strProductReleaseIDs = rs("ProductReleaseIDs") & ""
        strProductNameBase= rs("Name") & ""
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
		strPartner = rs("PartnerID") & ""
		strPDDPath = rs("PDDPath") & ""
		strSCMPath = rs("SCMPath") & ""
		strAccessoryPath = rs("AccessoryPath") & ""
		strIDInformationPath = rs("IDInformationPath") & ""
		strMSPEKSExecutionPath = rs("MSPEKSExecutionPath") & ""
		strSTLPath = rs("STLStatusPath") & ""
		strProgramMatrixPath = rs("ProgramMatrixPath") & ""
		strDocPM = rs("DocPM") & ""
        strODMSEPM = rs("ODMSEPMID") & ""
        strProcurementPM = rs("ProcurementPMID") & ""
        strPlanningPM = rs("PlanningPMID") & ""
        strODMHWPM = rs("ODMHWPMID") & ""
        strODMPIMPM = rs("ODMPIMPMID") & ""
        strHWPC = rs("HWPCID") & ""
		strSEPE = rs("SEPE") & ""
        FusionRequirements = rs("FusionRequirements") & ""
        CreateSimpleAvTypeAuto = rs("AutoSimpleAV") & ""
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
        strProductName = strOTSName
		strSustainingMgrID = rs("SustainingMgrID")
		strSustainingSEPMID = rs("SustainingSEPMID") & ""
		
    	' LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
    	strSysEngrProgramCoordinatorID = rs("SysEngrProgramCoordinatorID") & ""
    	' LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB

        strProgramBusinessManagerID = rs("ProgramBusinessManagerID") & ""
    
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
        strSCMOwnerID = rs("SCMOwnerId") & ""
        strEngineeringDataManagementID = rs("EngineeringDataManagementId") & ""
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

        if rs("bStableConsistent") & "" = "" or rs("bStableConsistent") & "" = "0" then
            StableConsistent = ""
        else
            StableConsistent = "checked"
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
		   
        if Not(rs("AllowFollowMarketingName")) and Not (isClone) then
            followMarketingName = 0
            showBrandsRow =""
            chkMKT=""
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

        if (DisplayToolsProject = "none" and showBrandsRow="") then
                showBrandsRow = "none"
        end if

		if rs("AddDCRNotificationList") then
		    AddDCRNotificationList = "checked"
		else
		    AddDCRNotificationList = ""
		end if
	
		strServiceLifeDate = trim(rs("ServiceLifeDate") & "")

        if trim(rs("ProductStatusID")) = 4 then
            strEditServiceEOLDate = ""
        end if

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

	    if  isclone= true then
             strProductStatus = "1"
        else
		strProductStatus = rs("ProductStatusID")	
        end if
		
				
		if trim(rs("ProductStatusID")) = "1" then
			EnableDCR = " disabled "
		else
			EnableDCR = ""
		end if
		strBaseUnit = rs("BaseUnit") & "" 
		strCurrentROM = rs("CurrentROM") & ""
		strCurrentWebROM = rs("CurrentWebROM") & ""
		strOSSupport = rs("OSSupport") & ""
		strImagePO = rs("ImagePO") & ""
		strImageChanges = rs("ImageChanges") & ""
        strSystemboardIDs = rs("SystemBoardIDs") & ""
		strSystemboardComments = rs("SystemboardComments") & ""
		strMachinePNPID = rs("MachinePNPID") & ""
		strMachinePNPComments = rs("MachinePNPComments") & ""
		strCommonimages = rs("Commonimages") & ""
		strCertificationStatus = rs("CertificationStatus") & ""
		strSWQAStatus = rs("SWQAStatus") & ""
		strPlatformStatus = rs("PlatformStatus") & ""
		strDCRApproverList = rs("DCRApproverList") & ""
		strToolAccessIDList = rs("ToolAccessList") & ""
        strFactoryId = rs("FactoryId") & ""
	    strIsPlatformRTM = rs("IsPlatformRTM") & ""
        strSharedAVMarketingPM = rs("SharedAVMarketingPMID") & ""
        strSharedAVPCPM = rs("SharedAVPCID") & ""
        sTeamRosterApprovers = rs("TeamRosterApprovers") & ""
		rs.Close
	end if		
    if lcase(trim(CreateSimpleAvTypeAuto)) = "true" then
       CreateSimpleAvTypeAutoChecked = "checked"
       CreateSimpleAvTypeManualChecked = ""
    else
       CreateSimpleAvTypeAutoChecked = ""
       CreateSimpleAvTypeManualChecked = "checked"
    end if
    if lcase(trim(FusionRequirements)) = "true" then
        FusionReqChecked = "checked"
        FusionLegacyReqChecked = ""
    else
        FusionReqChecked = ""
        FusionLegacyReqChecked = "checked"
        CreateSimpleAvTypeAutoChecked = "disabled"
        CreateSimpleAvTypeManualChecked = "disabled"
    end if


	TmpArray = split(strDistribution,";")
	strDistribution = ""
	for each strTemp in TmpArray 
		if trim(strTemp) <> "" then strDistribution = strDistribution & "; " & trim(strTemp)
	next
	if strDistribution <> "" then strDistribution = mid(strDistribution,3)


	TmpArray = split(strActionNotify,";")
	strActionNotify = ""
	for each strTemp in TmpArray 
		if trim(strTemp) <> "" then strActionNotify = strActionNotify & "; " & trim(strTemp)
	next
	if strActionNotify <> "" then strActionNotify = mid(strActionNotify,3)



    if request("ID") <> "" then
        Response.Cookies("ProductVersionID") = request("ID")
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
	
    Response.Cookies("ProductName") = strProductName
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
    strCompleteProductList = ""

    rs.Open "spGetProductsAll 1",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(request("ID")) <> trim(rs("ID") & "" ) or isClone  then 'Exclude current version, in the cloning situation, the source product should be in the list

            'Malichi, 4/13/2016, Bug 19360: Pulsar allows product to be created with duplicate name (Gather complete product list for duplicate name check on save)
            strCompleteProductList = strCompleteProductList & "<OPTION selected value=" & rs("ID") & ">" & rs("ProductName")  & "</OPTION>"

			if trim(strReferenceID) = trim(rs("ID") & "" ) then
				strReferenceList = strReferenceList & "<OPTION selected value=" & rs("ID") & ">" & rs("Product")  & "</OPTION>"
			elseif rs("ProductStatusID") < 5 then
				strReferenceList = strReferenceList & "<OPTION value=" & rs("ID") & ">" & rs("Product") & "</OPTION>"
			else
			    strInactiveList = strInactiveList & "<option>" & rs("Product") & "</option>"
			end if
		end if
		rs.MoveNext
	loop
	rs.Close
	
	strSbmList = ""
	

	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0

'---------------------------------------------------PBI 16164 ---------------------------------------------------
    isSCMPublished = 0

    rs.Open "usp_ProductProperties_IsSCMPublished " & request("ID"),cn,adOpenStatic
	if not rs.BOF and not rs.EOF then
			isSCMPublished = rs("IsSCMPublished")
	end if 
	rs.Close

'--------------------------------------------------- PBI 92708 ---------------------------------------------------
dim teamRosterList 
    teamRosterList = ""
    teamRosterList = "," & sTeamRosterApprovers & ","

dim retTeamRoster

function isChecked(TeamRosterId)
    if(strDCRAutoOpen = "checked") then
        retTeamRoster = ""
        TeamRosterId =  "," & TeamRosterId & ","
            if inStr(teamRosterList, TeamRosterId) > 0 then  
                retTeamRoster = " checked "   
             end if
        isChecked = retTeamRoster
    else
        retTeamRoster = " checked " 
        isChecked = retTeamRoster
    end if
end function

function isCheckedNoODM(TeamRosterId)
   if(strDCRNoOdm = "checked") then
        retTeamRoster = ""
        TeamRosterId =  "," & TeamRosterId & ","
            if inStr(teamRosterList, TeamRosterId) > 0 then  
                retTeamRoster = " checked "   
             end if
        isCheckedNoODM = retTeamRoster
    else
        retTeamRoster = " checked " 
        isCheckedNoODM = retTeamRoster
    end if
end function

'------------------------------------------------------------------------------------------------------------------            	
        %>
    <span id="spnErrorMessage" class="font-red hide"></span>
    <form action="ProgramSave_Pulsar.asp" method="post" name="ProgramInput">
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
                 <%if request("ID") = "" then%>
                <td id="CellPlatforms" style="display: none" width="150px">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Platforms')">Base Unit Groups</a>&nbsp;</b></font></td>
                    
                <%else%>
                <td id="CellPlatforms" style="display: " width="150px">
                    <font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Platforms')">Base Unit Groups</a>&nbsp;</b></font></td>

                <%end if%>
                 <td id="CellPlatformsb" style="display: none" width="150px" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Base Unit Groups</b></font></td>
           </tr>
            </table>
        <font size="1" face="verdana">
            <br>
        </font>
        <table id="tabGeneral" style="display: none; width:990px; overflow:scroll; border-collapse:collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Name:</font></strong><font color="red" size="1"> *</font>
                </td>
                <td>
                    <!-- Remove check for not being able to edit Product Name in Pulsar Product properties when base unit groups is RTMED-->
                    <input type="text" id="txtProductNameBase" name="txtProductNameBase" style="width: 160px;" value="<%=strProductNameBase%>"  maxlength="30" onkeyup="return EliminateFirstSapce();" />
                    <input type="hidden" id="tagProductNameBase" name="tagProductNameBase" style="width: 160px;" value="<%=strProductNameBase%>" />
                    
                    <!--Malichi, 4/13/2016, Bug 19360: Pulsar allows product to be created with duplicate name (Gather complete product list for duplicate name check on save)-->
                    <select id="cboCompleteProductList" name="cboCompleteProductList" style="display:none"><%=strCompleteProductList%></select>
                </td>                

                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Product&nbsp;Type:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td>
                    <select id="cboType" name="cboType" style="width: 200px;" language="javascript" onchange="return cboType_onchange()">
                        <option value="1" selected>PC (Notebook,Desktop,etc.)</option>           
                        <%if strType = "3" then%>
                        <option value="3" selected>Dock/Port Rep./Jacket</option>
                        <%else%>
                        <option value="3">Dock/Port Rep./Jacket</option>
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
            <tr style="display:none">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Name:</font></strong><font color="red" size="1"> *</font>
                </td>
                <td>
                    <%if request("ID") = "" then%>
                    <select id="cboFamily1" name="cboFamily1" style="width: 160px;" language="javascript"
                        onchange="return cboFamily_onchange()">
                        <option></option>
                        <%=strFamilies%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdAddFamily1" name="cmdAddFamily"
                        language="javascript" onclick="return cmdAddFamily_onclick()">
                    <%else%>
                    <input type="text" id="txtViewFamily1" style="width: 160px;" name="txtViewFamily"
                        disabled value="<%=strFamily%>"><select style="display: none" id="Select1" name="cboFamily1"
                            style="width: 180px;">
                            <option></option>
                            <%=strFamilies%>
                        </select>&nbsp;<input style="display: none" type="button" value="Add" id="Button1"
                            name="cmdAddFamily" language="javascript" onclick="return cmdAddFamily_onclick()">
                    <%end if%>
                </td>

                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Business&nbsp;Segment:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td>
                    <input type="text" id="txtVersion1" name="txtVersion1" style="width: 200px;" value="<%=strVersion%>"
                        maxlength="20">
                    <input type="hidden" id="tagVersion1" name="tagVersion1" style="width: 200px;" value="<%=strVersion%>"
                        maxlength="20">
                </td>
            </tr>

            <tr>
               <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Family:</font></strong><font color="red" size="1"> *</font>
                </td>
                <td>
                    <%if request("ID") = "" then%>
                    <select id="cboFamily" name="cboFamily" style="width: 380px;" language="javascript"
                        onchange="return cboFamily_onchange()">
                        <option></option>
                        <%=strFamilies%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdAddFamily" name="cmdAddFamily"
                        language="javascript" onclick="return cmdAddFamily_onclick()" />
                    <% elseif request("clone") = "1" then %>
                    <select id="cboFamily" name="cboFamily" style="width: 370px;" language="javascript"
                        onchange="return cboFamily_onchange()">
                        <option></option>
                        <%=strFamilies%>
                    </select>&nbsp;<input type="button" value="Add" id="Button2" name="cmdAddFamily"
                        language="javascript" onclick="return cmdAddFamily_onclick()" />
                    <script>cboFamily_selectcurrent("<%=strFamily%>");</script>
                    <%else%>
                    <input type="text" id="txtViewFamily" style="width: 450px;" name="txtViewFamily"
                        disabled value="<%=strFamily%>"><select style="display: none" id="Select3" name="cboFamily"
                            style="width: 180px;">
                            <option></option>
                            <%=strFamilies%>
                        </select>&nbsp;<input style="display: none" type="button" value="Add" id="Button3"
                            name="cmdAddFamily" language="javascript" onclick="return cmdAddFamily_onclick()">
                    <%end if%>
                </td>
				
                <td width="160" style="vertical-align: top"><strong><font size="2">Business&nbsp;Segment:</font></strong><font color="red" size="1">&nbsp;*</font></td>
                <td>

                   <!-- Check if any SCM(s) are published, if pusblished, dont allow user to change business segment  -  PBI 16164-->
                    <%if isSCMPublished = 0 then%>
                    <select id="cboBusinessSegmentID" name="cboBusinessSegmentID" style="width: 200px;" language="javascript" onchange="return cboBusinessSegment_onchange()">
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
                    <%elseif isSCMPublished <> 0 and isClone = true then%>
                    <select id="cboBusinessSegmentID" name="cboBusinessSegmentID" style="width: 200px;" language="javascript" onchange="return cboBusinessSegment_onchange()">
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
                    
                <%else%>
                    <select id="cboBusinessSegmentID" name="cboBusinessSegmentID" style="width: 200px; display:none;">
                         <option selected value=""></option>
                    <%  
                        rs.open "spPULSAR_Product_ListBusinessSegments",cn 
                        do while not rs.eof
                    %>
                       <%if trim(rs("BusinessSegmentId")) = trim(strBusinessSegmentID) then%>
                            <option selected value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                            <% strBusinessSegment = rs("name")%>
                        <%else%>
                            <option value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                        <%end if%>
                    <%
                            rs.movenext
                        loop
                        rs.close    
                    %>
                    </select>
                   <input type="text" id="txtViewBusinessSegment" style="width: 200px;" name="txtViewBusinessSegment" disabled value="<%=strBusinessSegment%>">
                <%end if%>
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
                <td width="120" style="display: <%=DisplayToolsProject%>" id="tdReleaseTitle">
                    <strong><font size="2">Release:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td style="display: <%=DisplayToolsProject%>" id="tdReleaseText">
                    <%  
                        if request("ID") = "" then
                            response.write "<a id=ReleaseLink href=""javascript: AddRelease(0);"">" & "Add New Release" & "</a>"
                        elseif strProductRelease = "" then
                            response.write "<a id=ReleaseLink href=""javascript: AddRelease(" & clng(request("ID")) & ");"">" & "Add New Release" & "</a>"
                        else
                            response.write "<a id=ReleaseLink href=""javascript: AddRelease(" & clng(request("ID")) & ");"">" & strProductRelease & "</a>"
                        end if
                    %>
                        <input type="hidden" id="txtVersion" name="txtVersion" style="width: 200px;" value="<%=strVersion%>" maxlength="20">
                        <input type="hidden" id="tagVersion" name="tagVersion" style="width: 200px;" value="<%=strVersion%>" maxlength="20">
                        <input type="hidden" id="txtProductRelease" name="txtProductRelease" style="width: 200px;" value="<%=strProductRelease%>" maxlength="500">
                        <input type="hidden" id="txtProductReleaseIDs" name="txtProductReleaseIDs" style="width: 200px;" value="<%=strProductReleaseIDs%>" maxlength="500">
                </td>
			</tr>
            <tr id="DevCenterRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top" class="tdClassBSV">
                    <strong><font size="2">Dev.&nbsp;Center:</font></strong><font color="red" size="1">&nbsp;*</font>
                </td>
                <td class="tdClassBSV">
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
                        <option value="6">Mobility</option>k
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
                       <select id="cboProductLine" name="cboProductLine" style="width: 200px;" language="javascript"
                        onchange="return cboProductLine_onchange()">
                        <option></option>
                        <%=strProductLines%>
                    </select>&nbsp;
				</td>
            </tr>
            <tr id="PreinstallRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top" class="tdClassBSV">
                    <strong><font size="2">Preinstall&nbsp;Team:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td class="tdClassBSV">
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
                        <%if strPreinstall = "8"  then%>
                        <option value="8" selected>No Preinstall Team</option>
                        <%else%>
                        <option value="8">No Preinstall Team</option>
                        <%end if%>
                    </select>
                </td>

                <td width="160" style="vertical-align: top">
                    <strong><font size="2">System&nbsp;Board&nbsp;ID:</font></strong>
                </td>
                <td>
                   
                    <%if strSystemboardIDs <> "" then
                        Response.write strSystemboardIDs 
                        end if %>
                    
                    <input type="hidden" id="txtSystemBoardID" name="txtSystemBoardID" style="width: 160px;"
                        value="<%=strSystemboardIDs%>">
                    <input type="hidden" id="txtSystemBoardComments" name="txtSystemBoardComments" style="width: 160px;"
                        value="<%=server.htmlencode(replace(strSystemBoardcomments,chr(161)&chr(168),"''"))%>">
                </td>
            </tr>
            <tr id="ReleaseRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top" class="tdClassBSV">
                    <strong><font size="2">Release&nbsp;Team:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td class="tdClassBSV">
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
                <td colspan="3">
                    <select id="cboPhase" name="cboPhase" style="width: 160px;" language="javascript"
                        onchange="return cboPhase_onchange()">
                        <%
				rs.open "spListProductStatuses",cn,adOpenStatic
				do while not rs.EOF
                    if request("ID") = "" or request("clone") = "1" then 'create new or clone just display 'definition'
                            
					    if trim(strProductStatus) = trim(rs("ID")) then
						    Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					   
					    end if
                    else
					if trim(strProductStatus) = trim(rs("ID")) then
						Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					else
						Response.Write "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
					end if
				
                    end if
									
					rs.MoveNext
				loop
				rs.Close
                        %>
                    </select>
                    <input type="hidden" id="tagPhase" name="tagPhase" value="<%=strProductStatus%>">
                    <input type="hidden" id="txtRegulatoryModel" name="txtRegulatoryModel" value="<%=strRegulatoryModel%>">
                </td>
                <td id="ToolPMRow" style="display: <%=DisplayToolsProject2%>" width="160" style="vertical-align: top">
                    <strong><font size="2">Project&nbsp;Manager:</font><font color="red" size="1">&nbsp;*</font></strong>
                </td>
                <td id="ToolPMRow2" style="display: <%=DisplayToolsProject2%>">
                    <select id="cboToolsPM" name="cboToolsPM" style="width: 160px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdToolsPMAdd" name="cmdToolsPMAdd"
                        language="javascript" onclick="return cmdToolsPMAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Factories:</font></strong><font color="red" size="1"></font>
                </td>
                <td>
                    <select id="cboFactory" name="cboFactory" multiple="true" style="width: 260px;height:100px;" language="javascript">
                        
                    <% 
                         if trim(request("ID")) = "" then
                           rs.open "spListProductFactories 0",cn,adOpenForwardOnly
                        else
                            rs.open "spListProductFactories " & clng( request("ID")),cn,adOpenForwardOnly
                        end if
				       do while not rs.EOF
					      if rs("Selected") ="1" or rs("Selected") ="true"   then
                            ' bhasSelected = 1
						     Response.Write "<OPTION selected value=" & rs("ManufacturingSiteId") & ">" & rs("Name") & " (" & rs("Code") & ")</OPTION>"
					      else
						     Response.Write "<OPTION value=" & rs("ManufacturingSiteId") & ">" & rs("Name") & " (" & rs("Code") & ")</OPTION>"
					      end if
				          rs.MoveNext
				       loop
				       rs.Close
          

                    %>
                    </select>
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Stable & Consistent:</font></strong><font color="red" size="1"></font>
                </td>
                <td width="160" style="vertical-align: top">
                    <input type="checkbox" <%=StableConsistent%> id="chkStableConsistent" name="chkStableConsistent">
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
            <tr>
               <td width="120" style="vertical-align: top">
                    <strong><font size="2">RCTO&nbsp;Sites:</font></strong>
                </td>
                <td nowrap colspan="3">
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
            <tr id="CycleRow" style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Product Group:</font></strong>
                </td>
                <td colspan="3">
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
            </tr>
            <!--Jason's code for create simple av, WILL UNCCOMMENT LATER-->
                <tr style="display: <%=DisplayToolsProject%>">
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Create Simple AV:</font></strong>
                </td>
                <td colspan="3">
                    <table>
                        <tr>
                            <td>
                                <input <%=CreateSimpleAvTypeAutoChecked%> id="optCreateSimpleAvTypeAuto" name="optCreateSimpleAvType" type="radio" value="1" /> <font face="verdana" size="2">Automatic flow of Features</font><br />
                                <input <%=CreateSimpleAvTypeManualChecked%> id="optCreateSimpleAvTypeManual" name="optCreateSimpleAvType" type="radio" value="0" /> <font face="verdana" size="2">Manual Export to SCM</font>

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
                                <input type="checkbox" disabled <%=chkMKT%>
                                    id="chkEnableFollowMarketingName" name="chkEnableFollowMarketingName">
                                <font face="verdana" size="2">This Product will follow the marketing name consistency requirements.<br>
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
            <%if (showBrandsRow = "") then%>
            <tr id="BrandRow" style="display: <%=showBrandsRow%>">
                <td width="160" style="vertical-align: top">
                    <div id="">
                        <strong><font size="2">Brands:</font></strong>
                        Use SCM column to combine multiple brands into 1 SCM. Blanks create a unique SCM. 
                        <i>Note: Brands must share the same Software and Product Drop to be combined into 1 SCM.</i>
                    </div>
                </td>
                <td colspan="3">
                    <%

            response.write "<input id=""txtBrandFrom"" name=""txtBrandFrom"" type=""hidden"" value"""">"
            response.write "<input id=""txtBrandTo"" name=""txtBrandTo"" type=""hidden"" value"""">"
 
                    %>
                    <br />
                    Filter Brands by:&nbsp;&nbsp;
                    <select id="cboBSBrandFilter" name="cboBSBrandFilter" style="width: 200px;">
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
                    </select><br /><br />
                   
                    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
                        border-left: steelblue 1px solid; border-bottom: steelblue 1px solid; height: 160px; width:100%; 
                        background-color: white" id="DIV3">
                        <table width="900px" id="TableBrand">
                            <thead>
                                <tr style="position: relative; top: expression(document.getElementById('DIV3').scrollTop-2);">
                                    <td width="10" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;
                                    </td>
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Brand
                                    </td>
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;SCMs
                                    </td>
                                    <td width="220" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Series
                                    </td>
                                    <td width="90" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Model Number
                                    </td>
                                     <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Screen Size
                                    </td>
                                     <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Generation
                                    </td>
                                     <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Form Factor
                                    </td>
                                    <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Suffix&nbsp;
                                    </td>
                                </tr>
                            </thead>
                            <%

		                    if request("ID") <> "" then
			                    rs.open  "spListBrands4Product_Pulsar " & clng(request("ID")) & ",0",cn,adOpenForwardOnly
		                    else
			                    rs.open  "spListBrands_pulsar ",cn,adOpenForwardOnly
		                    end if

		dim strRow
		dim strRowWait
		strRowWait = ""
        dim strDisableBrandInfo
        dim SCMEnabled
		do while not rs.eof
    		strRow = ""
			if rs("Active") or not isnull(rs("ProductBrandID")) then
					
				if isnull(rs("ProductBrandID")) then
					strChecked = ""
				else
					strChecked = "checked"
					strBrandsLoaded = strBrandsLoaded & "," & trim(rs("ID"))
				end if
                                    		
				if isnull(rs("LastPublishDt")) then
					strDisableBrandInfo = ""
                    SCMEnabled ="1"
				else
					strDisableBrandInfo = " disabled "
                    SCMEnabled ="0"
				end if 
				strAbbr = rs("name") & "" 
                if request("ID") ="" then
                    strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""checkbox"" " & strChecked & " BrandName=""" & rs("Name") & """ BNWOFormula=""" & rs("BrandsWOFormula") & """ title=" & rs("ID") & " id=chkBrands name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & "0" & ")"" value=""" & rs("ID") & """></TD>"
                else
		            strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""checkbox"" " & strChecked & " BrandName=""" & rs("Name") & """ BNWOFormula=""" & rs("BrandsWOFormula") & """ title=" & rs("ID") & " id=chkBrands name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & request("ID") & ")"" value=""" & rs("ID") & """></TD>"

                end if
                if strChecked <> "" then 
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

                'add a dropdown for common SCM between brands
                dim iSCMNum
                iSCMNum=0
                dim iSelectedSCMNumber
                iSelectedSCMNumber =clng(rs("SCMNumber"))
                
                strRow = strRow &  "<TD nowrap><select id=SelectSCM" & trim(rs("ID")) & " name=SelectSCM" & trim(rs("ID")) & " style='width: 70px;' " 
                strRow = strRow & strDisableBrandInfo
                strRow = strRow & ">"
                strRow = strRow & " <option "
                   
                strRow = strRow & " value='0'></option>"
               
                iSCMNum=1
                do while  iSCMNum<21
				    strRow = strRow & "  <option value='" & cstr(iSCMNum) & "'"
                    if iSCMNum = iSelectedSCMNumber then
                        strRow = strRow & " selected "
                    end if  
                    strRow = strRow &             " >SCM " & cstr(iSCMNum) &  "</option>"
			        iSCMNum = iSCMNum +1 
                loop    
                strRow = strRow &     "</select>"
                strRow = strRow &     "<INPUT type=""text"" id=txtSelectedSCM" & trim(rs("ID")) & " name=txtSelectedSCM" & trim(rs("ID")) & " style=""Display:none"" value=""" & iSelectedSCMNumber & """>"
                strRow = strRow &     "<INPUT type=""text"" id=txtSCMEnabled" & trim(rs("ID")) & " name=txtSCMEnabled" & trim(rs("ID")) & " style=""Display:none"" value=""" & SCMEnabled & """>"
                strRow = strRow &     "</TD>"  
                'end of SCM number
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
				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesA" & trim(rs("ID")) & " name=txtSeriesA" & trim(rs("ID")) & " value=""" & SeriesNameArray(0) & """"
                            
                strRow = strRow &  ">"
				strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesA" & trim(rs("ID")) & " name=tagSeriesA" & trim(rs("ID")) & " value=""" & SeriesNameArray(0) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDA" & trim(rs("ID")) & " name=txtSeriesIDA" & trim(rs("ID")) & " value=""" & SeriesIDArray(0) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesB" & trim(rs("ID")) & " name=txtSeriesB" & trim(rs("ID")) & " value=""" & SeriesNameArray(1) & """"
          
                strRow = strRow &  ">"
                strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesB" & trim(rs("ID")) & " name=tagSeriesB" & trim(rs("ID")) & " value=""" & SeriesNameArray(1) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDB" & trim(rs("ID")) & " name=txtSeriesIDB" & trim(rs("ID")) & " value=""" & SeriesIDArray(1) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesC" & trim(rs("ID")) & " name=txtSeriesC" & trim(rs("ID")) & " value=""" & SeriesNameArray(2) & """"
                    
                strRow = strRow &  ">"
                strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesC" & trim(rs("ID")) & " name=tagSeriesC" & trim(rs("ID")) & " value=""" & SeriesNameArray(2) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDC" & trim(rs("ID")) & " name=txtSeriesIDC" & trim(rs("ID")) & " value=""" & SeriesIDArray(2) & """>,"

				strRow = strRow &  "<INPUT style=""width:50"" type=""text"" id=txtSeriesD" & trim(rs("ID")) & " name=txtSeriesD" & trim(rs("ID")) & " value=""" & SeriesNameArray(3) & """"
		               
                strRow = strRow &  ">"
                strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesD" & trim(rs("ID")) & " name=tagSeriesD" & trim(rs("ID")) & " value=""" & SeriesNameArray(3) & """>"
				strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDD" & trim(rs("ID")) & " name=txtSeriesIDD" & trim(rs("ID")) & " value=""" & SeriesIDArray(3) & """>"
                
				strRow = strRow &  "</DIV></TD>"
                
                'Sruthi Changes for adding Model Number to ProductProperties and remove 5th, 6th series
                'Model number cells
                 strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:90"" class=""ModelNumber"" maxlength=""10"" type=""text"" id=txtModelNumber" & trim(rs("ID")) & " name=txtModelNumber" & trim(rs("ID")) & " value=""" & rs("ProductModelNumber") & """"
                 strRow = strRow &  ">" & "</TD>"
                'Dean Changes for adding Screen Size to ProductProperties
                'Screen size cells
                 strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:90"" maxlength=""10"" type=""text"" id=txtScreenSize" & trim(rs("ID")) & " name=txtScreenSize" & trim(rs("ID")) & " value=""" 
                 IF Not IsNull(rs("ScreenSize")) then
					strRow = strRow &  CStr(rs("ScreenSize")) & """"
				 else
					strRow = strRow & """"
				 end if
				if rs("Suffix") <> "AiO" and rs("Suffix") <> "All-in-One"  then
					strRow = strRow & " disabled "
				end if
                strRow = strRow &  ">" & "</TD>"

                ' generation and form factor cells
                strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:40"" maxlength=""5"" type=""text"" id=txtGeneration" & trim(rs("ID")) & " name=txtGeneration" & trim(rs("ID")) & " value=""" & rs("Generation") & """"

                strRow = strRow &  ">" & "</TD>"
                strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:70"" type=""text"" id=txtFormFactor" & trim(rs("ID")) & " name=txtFormFactor" & trim(rs("ID")) & " value=""" & rs("FormFactor") & """"
             
                strRow = strRow &  ">" & "</TD>"
            

				if len(rs("Suffix") & "") > 13 then
                    strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & left(rs("Suffix"),13) & "...&nbsp;</font></TD></TR>"
                else
				    strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & rs("Suffix") & "&nbsp;</font></TD></TR>"
                end if




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
            strBrandsChecked = strBrandsLoaded
		end if
                            %>
                        </table>
                    </div>
                </td>
            </tr>
            <%end if %> 
            <tr style="display:none">
                <td id="Td3" width="160" style="vertical-align: top">
                    <strong><font size="2">Service&nbsp;Tag:</font></strong>
                </td>
                <td id="Td4" style="display: <%=DisplayToolsProject%>" colspan="3">
                    <input type="text" id="txtServiceTag" name="txtServiceTag" style="width: 720px;"
                        value="<%=strServiceTag%>" maxlength="100">
                </td>
            </tr>
            <tr style="display:none">
                <td id="Td5" width="160" style="vertical-align: top" >
                    <strong><font size="2">BIOS&nbsp;Branding:</font></strong>
                </td>
                <td id="Td6" style="display: <%=DisplayToolsProject%>" colspan="3">
                    <input type="text" id="txtBIOSBranding" name="txtBIOSBranding" style="width: 720px;"
                        value="<%=strBIOSBranding%>" maxlength="100">
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
                            response.write "<option value=2>Program Office Program Manager</option>"
                        else
                            response.write "<option value=1>Configuration Manager</option>"
                            response.write "<option selected value=2>Program Office Program Manager</option>"
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
                                <input <%=strDcrToCm%> type="radio" id="chkDCRAutoOpen" name="chkDCRAutoOpen" value="1" />
                            </td>
                            <td>
                                DCR is assigned to the CM/POPM for review.
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDCRAutoOpen%> type="radio" id="Radio1" name="chkDCRAutoOpen"  value="2" />
                            </td>
                            <td>
                                Automatically assign these Primary System Team members as approvers and set status to "Investigating" for new DCRs."
                            </td>
                        </tr>
                        <tr id="TeamRoster2" class="DCRApprover">
                            <td></td>
                            <td>
                                <%  ' get team rosters
                                    dim cyclecount, ODMIncluded
		                              cyclecount = 0
                                      ODMIncluded = 0

                                    set cm2 = server.CreateObject("ADODB.Command")
	                                Set cm2.ActiveConnection = cn
                                    set rs2 = server.CreateObject("ADODB.recordset")
                                        rs2.open "usp_ProductSystemTeamRoster_GetAll " & clng(ODMIncluded), cn, adOpenForwardOnly
                                    while Not rs2.EOF
                                        cyclecount = cyclecount + 1
                                 %>                                    
                                    <span style="width:300px;float:left;">    
                                        <input type="checkbox" <%=sInitialDCR%>  name="ckTeamRosterAndODM" ID="ckTeamRosterAndODM" class="clsTeamRoster"  value="<%=rs2("TeamRosterId")%>" <%=isChecked(rs2("TeamRosterId"))%> /> <%=rs2("TeamRosterName")%>
                                    </span>                       
                                <% if cyclecount = 2 then%>
                                    <br /> 
                                <% cyclecount = 0
                                    end if %>                             
                                <%                
                                        rs2.movenext
                                    Wend
                                        rs2.Close
                                    set rs2 = Nothing   
                                %> 
                            </td>
 
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDCRNoOdm%>  type="radio" id="Radio3" name="chkDCRAutoOpen" value="4"/>
                            </td>
                            <td>
                                Automatically assign these Primary System Team members <span style="color: Red"> (Excluding ODM) </span> as approvers and set status to "Investigating" for new DCRs.
                            </td>
                        </tr>
                        <tr id="TeamRoster4" class="DCRApprover">
                            <td></td>
                            <td>
                                <%  ' get team rosters
                                    dim cyclecount2, ODMExcluded
		                                cyclecount2 = 0
                                        ODMExcluded = 1

                                    set cm2 = server.CreateObject("ADODB.Command")
	                                Set cm2.ActiveConnection = cn
                                    set rs2 = server.CreateObject("ADODB.recordset")
                                        rs2.open "usp_ProductSystemTeamRoster_GetAll " & clng(ODMExcluded), cn, adOpenForwardOnly
                                    while Not rs2.EOF
                                        cyclecount2 = cyclecount2 + 1
                                 %>                                    
                                    <span style="width:300px;float:left;">     
                                        <input type="checkbox" name="ckTeamRosterNoODM" ID="ckTeamRosterNoODM" class="clsTeamRosterNoODM"  value="<%=rs2("TeamRosterId")%>" <%=isCheckedNoODM(rs2("TeamRosterId"))%> /> <%=rs2("TeamRosterName")%>
                                    </span>                       
                                <% if cyclecount2 = 2 then%>
                                    <br /> 
                                <% cyclecount2 = 0
                                    end if %>                             
                                <%                
                                        rs2.movenext
                                    Wend
                                        rs2.Close
                                    set rs2 = Nothing   
                                %> 
                            </td>
                        </tr>
                        <tr>
                            <td style="vertical-align: top;">
                                <input <%=strDcrToList%> type="radio" id="Radio2" name="chkDCRAutoOpen" 
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
        <table id="tabPlatforms" style="display:none ; width:100%; height:80%; border-collapse: collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
            <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Base Unit Groups:</font></strong>
                </td>
                <td>
                    <iframe id="PlatformFrame" frameBorder="0"marginheight="0px" marginwidth="0px" style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; border-left: steelblue 1px solid; border-bottom: steelblue 1px solid;margin-top:4px;height: 100%;width:100%" src="PlatformList.asp?ID=<%=request("ID")%>&FollowMKTName=<%=followMarketingName%>&isCM=<%=isCMPermission%>"></iframe>
                 
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
                        <strong><font size="2">ID Information<br>
                            Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtIDInformationPath" name="txtIDInformationPath"
                            maxlength="256" value="">
                        <input type="hidden" id="tagIDInformationPath" name="tagIDInformationPath" value="<%=strIDInformationPath%>"
                            maxlength="256">
                    </td>
                </tr>
                <tr>
                    <td nowrap width="160" style="vertical-align: top">
                        <strong><font size="2">MSPEKS(Execution) Path:</font></strong>
                    </td>
                    <td>
                        <input type="text" style="width: 100%" id="txtMSPEKSExecutionPath" name="txtMSPEKSExecutionPath"
                            maxlength="256" value="">
                        <input type="hidden" id="tagMSPEKSExecutionPath" name="tagMSPEKSExecutionPath" value="<%=strMSPEKSExecutionPath%>"
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
                        <font size="1" face="verdana">
                            <a href="javascript:TestPath(1);">PDD</a> | 
                            <a href="javascript:TestPath(2);">SCM</a> | 
                            <a href="javascript:TestPath(3);">STL Status</a> | 
                            <a href="javascript:TestPath(4);">Product Data Matrices</a> | 
                            <a href="javascript:TestPath(5);">Accessory Documents</a> | 
                            <a href="javascript:TestPath(6);">ID Information</a> | 
                            <a href="javascript:TestPath(7);">MSPEKS(Execution)</a>
                        </font>
                    </td>
                </tr>
            </table>
        </span>
        <span style="display: none" id="tabOTS">
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
                 <option selected value="0"></option>
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

    <div id="tabsystemTeamMsg" style="display:none;">The System Team Roster positions below identify the lead position for this product, not permissions. Permissions are handled in <a href="/IPulsar/Admin/System Admin/UsersAndRoles_Main.aspx" target="_blank">Users and Roles</a>.</div>
    <table id="tabSystemTeam" style="display: none; width:930px; border-collapse:collapse;" border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan">
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
                  
                </select>&nbsp;<input type="button" value="Add" id="cmdSMAdd" name="cmdSMAdd" language="javascript"
                    onclick="return cmdSMAdd_onclick()">
            </td>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Marketing/Product&nbsp;Mgmt:</font></strong>
            </td>
            <td>
                <select id="cboComMarketing" name="cboComMarketing" style="width: 140px;" language="javascript"
                    onkeypress="return combo_onkeypress()" 
                    onfocus="return combo_onfocus()" 
                    onclick="return combo_onclick()"
                    onkeydown="return combo_onkeydown()">
                   
                </select>&nbsp;<input type="button" value="Add" id="cmdComMarketingAdd" name="cmdComMarketingAdd"
                    language="javascript" onclick="return cmdComMarketingAdd_onclick()">
            </td>
        </tr>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Configuration&nbsp;Manager:</font></strong><label id="lblConfigurationManager"><font color="red" size="1">&nbsp;*</font></label>
            </td>
            <td>
                <select id="cboPM" name="cboPM" style="width: 140px;" language="javascript" 
                    onkeypress="return combo_onkeypress()"
                    onfocus="return combo_onfocus()" 
                    onclick="return combo_onclick()" 
                    onkeydown="return combo_onkeydown()">
                    
                </select>&nbsp;<input type="button" value="Add" id="cmdPMAdd" name="cmdPMAdd" language="javascript"
                    onclick="return cmdPMAdd_onclick()">
            </td>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Supply&nbsp;Chain:</font></strong>
            </td>
            <td>
                <select id="cboSupplyChain" name="cboSupplyChain" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdSupplyChainAdd" name="cmdSupplyChainAdd"
                        language="javascript" onclick="return cmdSupplyChainAdd_onclick()">
            </td>
        </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Program&nbsp;Office&nbsp;Program&nbsp;Manager:</font></strong><label id="lblProgramOfficeManager"><font color="red" size="1">&nbsp;*</font></label>
                </td>
                <td valign="top">
                    <span style="display: none" id="POPMConsOnly"><font face="verdana" size="2" color="green">&nbsp;Consumer Products
                        Only.</font></span>
                    <select id="cboTDCCM" name="cboTDCCM" style="width: 140px;"
                        language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()"
                        onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                    
                    </select>&nbsp;<input type="button" value="Add" id="cmdTDCCMAdd"
                        name="cmdTDCCMAdd" language="javascript" onclick="return cmdTDCCMAdd_onclick()">
                </td>

                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Service:</font></strong>
                </td>
                <td>
                    <select id="cboService" name="cboService" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        
                    </select>&nbsp;<input type="button" value="Add" id="cmdServiceAdd" name="cmdServiceAdd"
                        language="javascript" onclick="return cmdServiceAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Systems&nbsp;Engineering&nbsp;PM:</font></strong><font color="red"
                        size="1">&nbsp;*</font>
                </td>
                <td>
                    <select id="cboSEPM" name="cboSEPM" style="width: 140px;" language="javascript" 
                        onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" 
                        onclick="return combo_onclick()" 
                        onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdSEPMAdd" name="cmdSEPMAdd"
                        language="javascript" onclick="return cmdSEPMAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Quality:</font></strong>
                </td>
                <td>                    
                    <select id="cboQuality" name="cboQuality" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                     
                    </select>&nbsp;<input type="button" value="Add" id="cmdQualityAdd" name="cmdQuaalityAdd"
                        language="javascript" onclick="return cmdQualityAdd_onclick()">
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
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdPlatformDevelopmentAdd" name="cmdPlatformDevelopmentAdd"
                        language="javascript" onclick="return cmdPlatformDevelopmentAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Procurement&nbsp;PM:</font></strong>
                </td>
                <td>                    
                    <select id="cboProcurementPM" name="cboProcurementPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdProcurementPMAdd" name="cmdProcurementPMAdd"
                        language="javascript" onclick="return cmdProcurementPMAdd_onclick()">
                </td>
            </tr>     
            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Commodity&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPDE" name="cboPDE" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdPDEAdd" name="cmdPDEAdd" language="javascript"
                        onclick="return cmdPDEAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">ODM&nbsp;System&nbsp;Engineering&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboODMSEPM" name="cboODMSEPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdODMSEPMAdd" name="cmdODMSEPMAdd" language="javascript"
                        onclick="return cmdODMSEPMAdd_onclick()">
                </td>
            </tr>
            <tr>
            </tr>

            <tr>
                <td width="160" style="vertical-align: top">
                    <strong><font size="2">Planning&nbsp;PM:</font></strong>
                </td>
                <td>
                    <select id="cboPlanningPM" name="cboPlanningPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdPlanningPMAdd" name="cmdPlanningPMAdd" language="javascript"
                        onclick="return cmdPlanningPMAdd_onclick()">
                </td>
                <td width="160" style="vertical-align: top"><strong><font size="2">ODM&nbsp;HW&nbsp;PM:</font></strong></td>
                <td>
                    <select id="cboODMHWPM" name="cboODMHWPM" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdODMHWPMAdd" name="cmdODMHWPMAdd" language="javascript"
                        onclick="return cmdODMHWPMAdd_onclick()">
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
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdBIOSLeadAdd" name="cmdBIOSLeadAdd"
                        language="javascript" onclick="return cmdBIOSLeadAdd_onclick()">
                </td>

                <td width="160" style="vertical-align: top">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
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
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdSETestLeadAdd" name="cmdSETestLeadAdd"
                        language="javascript" onclick="return cmdSETestLeadAdd_onclick()">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top"><strong><font size="2">HW&nbsp;PC:</font></strong></td>
                <td>
                    <select id="cboHWPC" name="cboHWPC" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()"
                        onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdHWPCAdd" name="cmdHWPCAdd" language="javascript"
                        onclick="return cmdHWPCAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">SE&nbsp;Test&nbsp;Lead&nbsp;(Sec):</font></strong>
                </td>
                <td>
                    <select id="cboSETest" name="cboSETest" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                       
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
                      <option value="0"></option>
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
                        <option value="0"></option>
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
                    
                    </select>&nbsp;<input type="button" value="Add" id="cmdSysEngrProgramCoordinatorAdd" name="cmdSysEngrProgramCoordinatorAdd"
                        language="javascript" onclick="return cmdSysEngrProgramCoordinatorAdd_onclick()">
                </td>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Program&nbsp;Business&nbsp;Manager:&nbsp;</font></strong>
                </td>
                <td>
                    <select id="cboProgramBusinessManager" name="cboProgramBusinessManager" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        
                    </select>&nbsp;<input type="button" value="Add" id="Button5" name="cmdProgramBusinessManagerAdd"
                        language="javascript" onclick="return cmdProgramBusinessManagerAdd_onclick()">
                </td>
            </tr>
         <tr>
                <td width="120" style="vertical-align: top">
                    <strong><font size="2">Shared&nbsp;AV&nbsp;Marketing:</font></strong>
                </td>
                <td>
                    <select id="cboSharedAvMarketing" name="cboSharedAvMarketing" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdSharedAvMarketing" name="cmdSharedAvMarketing"
                        language="javascript" onclick="return cmdSharedAvMarketing_onclick()">
                </td>
                <td width="120" style="vertical-align: top"><strong><font size="2">Shared&nbsp;AV&nbsp;Program&nbsp;Coordinator:</font></strong></td>
                <td>
                    <select id="cboSharedAVPC" name="cboSharedAVPC" style="width: 140px;"
                        language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()"
                        onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdSharedAVPC" name="cmdSharedAVPC"
                        language="javascript" onclick="return cmdSharedAVPC_onclick()">
                </td>                                
            </tr>
            <tr>
                <%
                    rs.open "spPULSAR_Product_ListBusinessSegments", cn 
                    do while not rs.eof
                        if trim(rs("BusinessSegmentID")) = trim(strBusinessSegmentID) then
                            if rs("Operation") <> "0" then 
                                strIsDesktop = "YES"
                            end if
                            if trim(rs("BusinessId")) = "1" then
                                strIsCommercial = "YES"
                            end if
                        end if

                        if rs("Operation") = "0" then 
                            strBusinessSegmentList = strBusinessSegmentList + trim(rs("BusinessSegmentID")) + ",0" + "," + trim(rs("BusinessId")) + ";"
                        else
                            strBusinessSegmentList = strBusinessSegmentList + trim(rs("BusinessSegmentID")) + ",1" + "," + trim(rs("BusinessId")) + ";"
                        end if
                    rs.movenext
                    loop
                    rs.close
                %>
                <td id="tdSCMOwner1" width="270" style="vertical-align: top"><strong><font size="2">SCM&nbsp;Owner:</font></strong><label id="lblSCMOwner"><font color="red" size="1">&nbsp;*</font></label></td>
                <td id="tdSCMOwner2">
                    <select id="cboSCMOwner" name="cboSCMOwner" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                       
                    </select>&nbsp;<input type="button" value="Add" id="cmdSCMOwner" name="cmdSCMOwner" language="javascript" onclick="return cmdSCMOwner_onclick()">
                </td>
                <td width="160" style="vertical-align: top">
                        <strong><font size="2">ODM&nbsp;PIN&nbsp;PM:</font></strong>
                </td>
                <td>                    
                    <select id="cboODMPIMPM" name="cboODMPIMPM" style="width: 140px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input id="cmdODMPIMPMAdd"
                        language="javascript"
                        name="cmdODMPIMPMAdd"
                        onclick="return cmdODMPIMPMAdd_onclick()"
                        scmowner="" type="button" value="Add">
                </td>
            </tr>
            <tr>
                <td width="120" style="vertical-align: top"><strong><font size="2">Engineering&nbsp;Data&nbsp;Management:</font></strong><label id="lblEngineeringDataManagement"><font color="red" size="1">&nbsp;*</font></label></td>
                <td>
                    <select id="cboEngineeringDataManagement" name="cboEngineeringDataManagement" style="width: 140px;" language="javascript" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
                      
                    </select>&nbsp;<input type="button" value="Add" id="cmdEngineeringDataManagement" name="cmdEngineeringDataManagement" language="javascript" onclick="return cmdEngineeringDataManagement_onclick()">
                </td>               
                <td width="120" style="vertical-align: top">
                    &nbsp;
                </td>
                <td>
                  &nbsp;
                </td>
            </tr>            
    </table>

    <input style="display: none" type="text" id="txtID" name="txtID" value="<%=request("ID")%>">
    <input type="hidden" id="txtProductName" name="txtProductName">
    <input type="hidden" id="tagProductFullName" name="tagProductFullName" value="<%=strOTSName%>">
    <input type="hidden" id="txtServiceLifeDate" name="txtServiceLifeDate" value="<%=strServiceLifeDate%>">

    
    <input type="hidden" id="txtProductLine" name="txtProductLine" value="<%=strProductLine%>">

    
    <input type="hidden" id="txtProductFamily" name="txtProductFamily" value="<%=strFamily%>">
    <input type="hidden" id="txtOSListChanged" name="txtOSListChanged" value="0">
    <input type="hidden" id="txtFullOSList" name="txtFullOSList" value="">
    <input type="hidden" id="txtBrandsLoaded" name="txtBrandsLoaded" value="<%=strBrandsLoaded%>">
    <input type="hidden" id="txtReleasesLoaded" name="txtReleasesLoaded" value="<%=strReleasesLoaded%>">
    <input type="hidden" id="txtBrands" name="txtBrands" value="">
    <input type="hidden" id="txtInitialSystemBoardID" name="txtInitialSystemBoardID"
        value="<%=strSystemboardIDs%>">
    <input type="hidden" id="txtInitialMachinePnPID" name="txtInitialMachinePnPID" value="<%=strMachinePnPID%>">
    <input type="hidden" id="txtInitialAffectedProduct" name="txtInitialAffectedProduct" value="<%=iAffectedProduct%>">
    <input type="hidden" id="txtIsSEPM" name="txtIsSEPM" value="<%=strIsSEPM%>">
    <input type="hidden" id="hdnIsClone" name="isClone" value="<% if isClone = true then Response.Write("1") else Response.Write("0") end if %>" />
    <input type="hidden" id="txtBrandsAdded" name="txtBrandsAdded" value="<%=strBrandsAdded%>">
    <input type="hidden" id="hdnIsDesktop" name="hdnIsDesktop" value="<%=strIsDesktop%>" />
    <input type="hidden" id="hdnIsCommercial" name="hdnIsCommercial" value="<%=strIsCommercial%>" />    
    <input type="hidden" id="hdnBusinessSegmentList" name="hdnBusinessSegmentList" value="<%=strBusinessSegmentList%>" />
    <input type="hidden" id="hdnRTPandEMDatePASS" name="hdnRTPandEMDatePASS"/>
    <input type="hidden" id="hdnCurrentUser" name="hdnCurrentUser" value="<%=CurrentUser%>" />
    <input type="hidden" id="hdnProductPartner" name="hdnProductPartner" value="<%=strPartner%>" />
    <input type="hidden" id="hdnBusinessSegmentID" name="hdnBusinessSegmentID" value="<%=strBusinessSegmentID%>" />
    <input type="hidden" id="hdnEnableFollowMarketingName" name="hdnEnableFollowMarketingName" value="<%=followMarketingName%>" />

    <div id="Dialog1" style="display:none" title="Pulsar - Dialog">
        <iframe frameborder="0" name="DialogIframe" id="DialogIframe" style="width:100%; height:100%"></iframe>
    </div>

    <div id="Dialog2" style="display:none" title="Pulsar - Dialog">
        <iframe frameborder="0" name="DialogIframe2" id="DialogIframe2" style="width:100%; height:100%"></iframe>
    </div>

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
<script type="text/javascript">
    $(window).load(function () {
        ValidatePagePermission("ProgramMain", "Pulsar Product");
    });
</script>