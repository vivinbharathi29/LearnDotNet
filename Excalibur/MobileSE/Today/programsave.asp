<%@  language="VBScript" %>
<% Response.CodePage = 65001
   Response.Charset="UTF-8" %>
<!-- #include file="../../includes/emailwrapper.asp" -->
<!-- #include file="../../includes/emailqueue.asp" -->
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script src="../../Scripts/shared_functions.js"></script>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script src="../../Scripts/Pulsar2.js"></script>

<script id="clientEventHandlersJS" language="javascript">
<!--
    function window_onload() {
        var sDialogView = globalVariable.get('product_prop_view');
        if (txtSuccess.value != "0") {
            if (document.getElementById("txtAddingProduct").value == '0') { //close the jquery pop up when editing
                if (isFromPulsar2()) {
                    closePulsar2Popup(true);
                }
                else if (IsFromPulsarPlus()) {
                    if (GetQueryStringValue('method') == "missingsystemboard") {
                        window.parent.parent.parent.MissingSystemBoardCallBack(txtSuccess.value);
                    }
                    else if (GetQueryStringValue('method') == "missingsetest") {
                        window.parent.parent.parent.MissingSETestCallBack(txtSuccess.value);
                    }
                    ClosePulsarPlusPopup();
                }
                else {
                    if (CheckOpener() === false) {
                        parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
                    } else {
                        window.returnValue = txtSuccess.value;
                        window.close();
                    }
                }
            } else {
                if (document.getElementById("preferredLayout").value == 'pulsar2') {
                    alert('Product Added Successfully');
                    parent.parent.window.parent.location = "../../../Excalibur/Excalibur.asp?path=pmview.asp%3FClass%3D1%26ID%3D" + txtSuccess.value;
                }
                else if (IsFromPulsarPlus()) {
                    if (GetQueryStringValue('method') == "missingsystemboard") {
                        window.parent.parent.parent.MissingSystemBoardCallBack(txtSuccess.value);
                    }
                    else if (GetQueryStringValue('method') == "missingsetest") {
                        window.parent.parent.parent.MissingSETestCallBack(txtSuccess.value);
                    }
                    ClosePulsarPlusPopup();
                }
                else {
                    if (CheckOpener() === false && sDialogView == 'add') {//close the jquery pop up when adding new product
                        //the ClosePropertiesDialog is initiated from leftmenu's Add New link
                        parent.parent.window.parent.ClosePropertiesDialog(txtSuccess.value, true, null);
                    } else {
                        window.returnValue = txtSuccess.value;
                        window.close();
                    }
                }
            }

        }
    }

    function CheckOpener() {
        //If True, page opened with showModalDialog
        //If False, page opened with JQuery Modal Dialog
        var oWindow = window.dialogArguments;
        return (oWindow == null) ? false : true;
    }

//-->
</script>

</HEAD>
<body language="javascript" onload="return window_onload()">
    Saving Program.&nbsp; Please Wait...
    <%



	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("%", "select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
	

	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	dim strBrands
	dim blnFailed
	dim rs
	dim p
	dim BrandArray
	dim ReleaseArray
	dim strPath
	dim i
	dim strOutput
	dim strSeriesString
	dim bAddingProduct
    dim isCloning

	strSeriesString = ""
	blnFailed = false
	strBrands = request("txtBrands")
	
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open
	
	
	dim CurrentUser
	dim CurrentDomain
	dim CurrentUserEmail
	dim CurrentUserFirstName
	
    '--Bug 18879/Task 18966 - Declare Current User's FullName variable
    dim FullName
	
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if
	

	set rs = server.CreateObject("ADODB.Recordset")
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
		
	if (rs.EOF and rs.BOF) then
		CurrentUserEmail  = "max.yu@hp.com"
		CurrentUserFirstName = "Efren"
	else
        '--Bug 18879/Task 18966 - Define Current User's FullName variable
        FullName = rs("Name")

		CurrentUserEmail = rs("Email") & ""
		if instr(rs("Name") & "",",")> 0 then
			CurrentUserFirstName = mid(rs("Name") & "",instr(rs("Name") & "",",")+1)
		else
			CurrentUserFirstName = ""
		end if
	end if	
	rs.close

	
	cn.BeginTrans

	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
    if request("txtID") = "" then
		cm.CommandText = "spAddProductVersion"
		bAddingProduct = True
    elseif request("isClone") = 1 then
        cm.CommandText = "spAddProductVersion"
        isCloning = True
	else
		cm.CommandText = "spUpdateProductVersion"
		bAddingProduct = False
        isCloning = False
	end if

	cm.CommandType =  &H0004

    'Add the following hidden field to tell if it's a new product as the pop-up is opened 
    'differently and need to be closed differently
    if bAddingProduct then
        Response.Write "<INPUT style='Display:none' type='text' id='txtAddingProduct' name='txtAddingProduct' value='1'>"
    else
        Response.Write "<INPUT style='Display:none' type='text' id='txtAddingProduct' name='txtAddingProduct' value='0'>"
    end if

	if bAddingProduct = false and isCloning = false then
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value = clng(request("txtID"))
		cm.Parameters.Append p
	end if
	
	set p =  cm.CreateParameter("@FamilyID", 3, &H0001)
	p.value = clng(request("cboFamily"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Version", 200, &H0001, 20)
	p.Value = left(request("txtVersion"),30)
	cm.Parameters.Append p

   'START CHANGE - ADD PRODUCT LINE @ BUSINESS SEGMENT   
    set p =  cm.CreateParameter("@BusinessSegmentID", 3, &H0001)
	p.value = clng(request("cboBusinessSegmentID"))
	cm.Parameters.Append p 
  
    set p =  cm.CreateParameter("@ProductLineID", 3, &H0001)
    if trim(request("cboType")) = "2" then
		p.value = 0
	else
		p.value = clng(request("cboProductLine"))
	end if
	
	cm.Parameters.Append p	
   'END CHANGE - ADD PRODUCT LINE @ BUSINESS SEGMENT
        
	set p =  cm.CreateParameter("@PMID", 3, &H0001)
	if trim(request("cboType")) = "2" then
		p.value = clng(request("cboToolsPM"))
	else
		p.value = clng(request("cboPM"))
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@TDCCMID", 3, &H0001)
	if request("cboTDCCM") <> "" then
		p.value = clng(request("cboTDCCM"))
	else
		p.value = null
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SMID", 3, &H0001)
	if trim(request("cboType")) = "2" then
		p.value = clng(request("cboToolsPM"))
	else
		p.value = clng(request("cboSM"))
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SEPMID", 3, &H0001)
	if trim(request("cboType")) = "2" then
		p.value = clng(request("cboToolsPM"))
	else
		p.value = clng(request("cboSEPM"))
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PDEID", 3, &H0001)
	p.value = clng(request("cboPDE"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SCFactoryEngineerID", 3, &H0001)
	p.value = clng(request("cboFactoryEngineer"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@AccessoryPMID", 3, &H0001)
	p.value = clng(request("cboAccessoryPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ComMarketingID", 3, &H0001)
	p.value = clng(request("cboComMarketing"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ConsMarketingID", 3, &H0001)
	p.value = clng(request("cboConsMarketing"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SMBMarketingID", 3, &H0001)
	p.value = clng(request("cboSMBMarketing"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PlatformDevelopmentID", 3, &H0001)
	p.value = clng(request("cboPlatformDevelopment"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SupplyChainID", 3, &H0001)
	p.value = clng(request("cboSupplyChain"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ServiceID", 3, &H0001)
	p.value = clng(request("cboService"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@QualityID", 3, &H0001)
	p.value = clng(request("cboQuality"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UpdateReleases", 16, &H0001)
	if (request("tagPhase") = 1 or bAddingProduct = true or isCloning = true) and request("cboPhase") <> 1 then
		p.value = 1 'Update Dates	
	elseif request("tagPhase") <> 1 and request("cboPhase") = 1 then
		p.value = 2 'Empty Dates	
	else
		p.value = 0 'Do Nothing
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Active", 16, &H0001)
	if request("cboPhase")	 = 1 or request("cboPhase")	 = 2 then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Sustaining", 16, &H0001)
	if request("cboPhase")	 = 3 then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProductStatusID", 3, &H0001)
	if request("cboPhase") = "" then
		p.Value = 1 'Default
	else
		p.Value = clng(request("cboPhase") )
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@DivisionID", 3, &H0001)
	p.value = clng(request("cboDivision"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@TypeID", 16, &H0001)
	p.value = clng(request("cboType"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DevCenter", 16, &H0001)
    if cint(request("cboDevCenter")) = 0 then
	    p.Value = 1
    else 
        p.Value = cint(request("cboDevCenter"))
    end if
	
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@DCRDefaultOwner", 3, &H0001)
	if trim(request("cboDCRDefaultOwner")) = "1" then
        p.value = 1  'Configuration Manager
    else
        p.value = 0  'Program Office Manager
    end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ReferenceID", 3, &H0001)
	if request("cboReference") = "0" or request("cboReference") = "" then
		p.value = null
	else
		p.value = clng(request("cboReference"))
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PartnerID", 3, &H0001)
	if clng(request("cboType")) = 2 then
        p.value = 1
    else
        p.value = clng(request("cboPartner"))
	end if
    cm.Parameters.Append p

	set p =  cm.CreateParameter("@PreinstallTeam", 3, &H0001)
	p.value = clng(request("cboPreinstall"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ReleaseTeam", 3, &H0001)
	p.value = clng(request("cboReleaseTeam"))
	cm.Parameters.Append p
	

	Set p = cm.CreateParameter("@Distribution", 200, &H0001, 1000)
	p.Value = left(request("txtDistribution"),1000)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@ConveyorBuildDistribution", 200, &H0001, 1000)
	p.Value = Left(Request("txtCvrBuildDist"), 1000)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@ConveyorReleaseDistribution", 200, &H0001, 1000)
	p.Value = Left(Request("txtCvrReleaseDist"), 1000)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@ActionNotifyList", 200, &H0001, 1000)
	p.Value = left(request("txtActionNotifyList"),1000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ToolAccessList", 200, &H0001, 1000)
	p.Value = left(replace(request("chkToolAccessID")," ",""),1000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Brands", 200, &H0001, 255)
	p.Value = left(strBrands,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DOTSName", 200, &H0001, 30)
	p.value = left(request("txtProductName"),30)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@RCTOSites", 200, &H0001, 50)
	p.value = left(request("txtRCTOSites"),50)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@RegulatoryModel", 200, &H0001, 15)
	p.value = left(request("txtRegulatoryModel"),15)
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@Emailactive", 16, &H0001)
	if request("chkEmail")	= "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Fusion",11, &H0001)
	if request("chkFusion") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@AllowSMR",11, &H0001)
	if request("chkEnableSMR") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@AllowDeliverableReleases",11, &H0001)
	if request("chkEnableDeliverables") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@AllowImageBuilds",11, &H0001)
	if request("chkEnableImages") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
	

	Set p = cm.CreateParameter("@AllowDCR",11, &H0001)
	if request("chkEnableDCR") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@OnCommodityMatrix",11, &H0001)
	if request("chkCommodities") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OnlineReports", 16, &H0001)
	if request("chkReports")	= "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@DCRAutoOpen", 16, &H0001)
    p.Value = request("chkDCRAutoOpen")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@BaseUnit", 200, &H0001, 2000)
	p.Value = left(request("txtBaseUnit"),2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CurrentROM", 200, &H0001, 200)
	p.Value = left(request("txtCurrentROM"),200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CurrentWebROM", 200, &H0001, 200)
	p.Value = left(request("txtCurrentWebROM"),200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OSSupport", 200, &H0001, 2000)
	p.Value = left(request("txtOSSupport"),2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ImagePO", 200, &H0001, 8000)
	p.Value = left(request("txtImagePO"),8000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ImageChanges", 200, &H0001, 2000)
	p.Value = left(request("txtImageChanges"),2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SystemBoardID", 200, &H0001, 200)
	p.Value = left(request("txtSystemBoardID"),200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SystemBoardComments", 200, &H0001, 1200)
	p.Value = left(request("txtSystemBoardComments"),1200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@MachinePnPID", 200, &H0001, 200)
	p.Value = left(request("txtMachinePnPID"),200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@MachinePnPComments", 200, &H0001, 1200)
	p.Value = left(request("txtMachinePNPComments"),1200)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CommonImages", 200, &H0001, 300)
	p.Value = left(request("txtCommonIMages"),300)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CertificationStatus", 200, &H0001, 3000)
	p.Value = left(request("txtCertificationStatus"),3000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SWQAStatus", 200, &H0001, 8000)
	p.Value = left(request("txtSWQAStatus"),8000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PlatformStatus", 200, &H0001, 8000)
	p.Value = left(request("txtPlatformStatus"),8000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PDDPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtPDDPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SCMPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtSCMPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@AccessoryPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtAccessoryPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@STLStatusPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtSTLPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProgramMatrixPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtProgramMatrixPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Description", 201, &H0001, 2147483647)
	p.Value = request("txtDescription")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Objectives", 201, &H0001, 2147483647)
	p.Value = request("txtObjective")
	cm.Parameters.Append p
	
	set p =  cm.CreateParameter("@SEPE", 3, &H0001)
	p.value = clng(request("cboSEPE"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PINPM", 3, &H0001)
	p.value = clng(request("cboPINPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SETestLead", 3, &H0001)
	p.value = clng(request("cboSETestLead"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SETestID", 3, &H0001)
	p.value = clng(request("cboSETest"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ODMTestLeadID", 3, &H0001)
	p.value = clng(request("cboODMTestLead"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@WWANTestLeadID", 3, &H0001)
	p.value = clng(request("cboWWANTestLead"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@BIOSLead", 3, &H0001)
	p.value = clng(request("cboBIOSLead"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@CommHWPM", 3, &H0001)
	p.value = clng(request("cboCommHWPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@VideoMemoryPM", 3, &H0001)
	p.value = clng(request("cboVideoMemoryPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@GraphicsControllerPM", 3, &H0001)
	p.value = clng(request("cboGraphicsControllerPM"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ProcessorPM", 3, &H0001)
	p.value = clng(request("cboProcessorPM"))
	cm.Parameters.Append p
	
	set p =  cm.CreateParameter("@SustainingMgrID", 3, &H0001)
	p.value = clng(request("cboSustainingMgr"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SustainingSEPMID", 3, &H0001)
	p.value = clng(request("cboSustainingSEPM"))
	cm.Parameters.Append p

  'LY BEGINNING OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB
	set p =  cm.CreateParameter("@SysEngrProgramCoordinatorID", 3, &H0001)
	p.value = clng(request("cboSysEngrProgramCoordinator"))
	cm.Parameters.Append p
  'LY END OF CHANGE - ADD SYSTEM ENGINEERING PROGRAM COORDINATOR FIELD TO SYSTEM TEAM TAB

	set p =  cm.CreateParameter("@PreinstallCutoff", 200, &H0001, 15)
	p.value = trim(request("cboPinCutoff"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PCID", 3, &H0001)
	p.value = trim(request("cboPC"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@MarketingOpsID", 3, &H0001)
	p.value = trim(request("cboMarketingOps"))
	cm.Parameters.Append p
	
    Set p = cm.CreateParameter("@ShowOnWhql",11, &H0001)
	if request("chkMdaCompliance") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@DCRApproverList", 200, &H0001, 1000)
    p.value = replace(trim(request("txtDCRApproverList")), ",", ";")
	cm.Parameters.Append p
	
	set p = cm.CreateParameter("@GPLM", 3, &H0001)
	p.value = trim(request("cboGplm"))
	cm.Parameters.Append p

	set p = cm.CreateParameter("@SPDM", 3, &H0001)
	p.value = trim(request("cboSpdm"))
	cm.Parameters.Append p

	set p = cm.CreateParameter("@SBA", 3, &H0001)
	p.value = trim(request("cboSBA"))
	cm.Parameters.Append p
	
	set p = cm.CreateParameter("@DocPM", 3, &H0001)
	p.value = trim(request("cboDocPM"))
	cm.Parameters.Append p

    set p = cm.CreateParameter("@DKCID", 3, &H0001)
	p.value = trim(request("cboDKC"))
	cm.Parameters.Append p
	
	set p = cm.CreateParameter("@MinRoHSLevel", 3, &H0001)
	p.value = trim(request("cboMinRoHS"))
	cm.Parameters.Append p
	
    Set p = cm.CreateParameter("@BSAMFlag",11, &H0001)
	if request("chkBSAMFlag") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
	
	set p = cm.CreateParameter("@AffectedProduct", 3, &H0001)
	if LCase(request("txtIsSEPM")) = "true" then
	    select case request("rblAffectedProduct")
            case 1
	            p.value = 0
            case 2
	            p.value = -1
            case 3
	            p.value = trim(request("cboMilestones"))
        end select
    else
	    p.value = trim(request("txtInitialAffectedProduct"))
	end if
	cm.Parameters.Append p
	
    Set p = cm.CreateParameter("@AddDCRNotificationList",11, &H0001)
	if request("chkAddDCRNotificationList") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
	
	if bAddingProduct = True or isCloning = True then
		set p = cm.CreateParameter("@NewID", 3, &H0002)
		cm.Parameters.Append p
	end if
    
    Set p = cm.CreateParameter("@FinanceID", 3, &H0001)
	p.Value = 0
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@ODMSEPMID", 3, &H0001)
	p.value = clng(request("cboODMSEPM"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@ProcurementPMID", 3, &H0001)
	p.value = clng(request("cboProcurementPM"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@PlanningPMID", 3, &H0001)
	p.value = clng(request("cboPlanningPM"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@ODMPIMPMID", 3, &H0001)
	p.value = clng(request("cboODMPIMPM"))
	 cm.Parameters.Append p

   '--Bug 18879/Task 18966 - If Updating, Add FullName parameter
    If bAddingProduct = False Then
        Set p = cm.CreateParameter("@FullName", adVarchar, adParamInput, 50, FullName)
	    cm.Parameters.Append p
    End If

    Set p = cm.CreateParameter("@WWANProduct",11, &H0001)
	if request("cboWWAN") = "" then
		p.Value = null
	else
		p.Value = clng(request("cboWWAN"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CommodityLock",11, &H0001)
	if request("chkCommodityLock") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
    
    'if textbox is disabled, use original value
    Set p = cm.CreateParameter("@ServiceLifeDate", 135, &H0001)
	if request("txtServiceEndDate") <> "" then
		p.Value = CDate(request("txtServiceEndDate"))
	else
        if request("txtServiceLifeDate") <> "" then
		    p.Value = CDate(request("txtServiceLifeDate"))
        else
            p.Value = null
        end if
	end if
	cm.Parameters.Append p

	cm.Execute RowsEffected
	
	dim strID
	if bAddingProduct = true or isCloning = true then
		strID = cm("@NewID")
	else
		strID = clng(request("txtID"))'cm("@ID")
	end if
	dim ConsumerArray
	dim CommercialArray
	dim SMBArray
	dim ConsumerTagArray
	dim CommercialTagArray
	dim SMBTagArray
	dim strApprovers
	dim strApproversTag
	dim ApproversArray
	dim ApproversTagArray
	dim OSArray
	dim strSeries1
	dim strSeries2
	dim strSeries3
	dim strSeries4
	dim strSeries5
	dim strSeries6
	dim SeriesArray
	dim CycleArray
	dim strCycle

    ' ********************************* Add Product Release *************************************
    if request("txtProductRelease") <> "" then
	    set cm = server.CreateObject("ADODB.Command")
	    cm.CommandType =  &H0004
	    cm.ActiveConnection = cn
		
	    cm.CommandText = "usp_ProductVersion_AssignRelease"	

        Set p = cm.CreateParameter("@ID", 3,  &H0001)
        p.Value = clng(strID)
        cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ReleaseIDs", 200,  &H0001, 30)
	    p.Value = left("",30)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Releases", 200,  &H0001, 30)
	    p.Value = request("txtProductRelease")
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@BusinessSegmentID", 3,  &H0001)
        p.Value = clng(request("cboBusinessSegmentID"))
        cm.Parameters.Append p
					
		cm.Execute RowsEffected
	    set cm=nothing
			
		if cn.Errors.count <> 0 then
			blnFailed = true
		end if	
    end if
    ' *******************************************************************************************

	if (bAddingProduct = true or isCloning = true or trim(request("txtID")) = "0") and strID <> "" and trim(request("txtAddCycle")) <> "" then
        'Add the new product to the selected cycles
        CycleArray = split(replace(request("txtAddCycle")," ",""),",")
        for each strCycle in CycleArray
        	
        	set cm = server.CreateObject("ADODB.command")

        	cm.ActiveConnection = cn
        	cm.CommandText = "spLinkProductToProgram"
        	cm.CommandType =  &H0004

        	Set p = cm.CreateParameter("@Program", 3, &H0001)
        	p.Value = clng(strCycle)
        	cm.Parameters.Append p

        	Set p = cm.CreateParameter("@Product", 3, &H0001)
        	p.Value = clng(strID)
        	cm.Parameters.Append p

        	cm.Execute RowsEffected
        	if RowsEffected <> 1 then
				blnFailed = true
                set cm = nothing
				exit for
            else
                set cm = nothing
        	end if
        next 
	end if	

	if request("txtOSListChanged") = "1" then
		dim strOldOSPreinstallList
		dim strOldOSWebList
		strOldOSPreinstallList=""
		strOldOSWebList=""
		rs.open "spListProductOS " & strID,cn,adOpenForwardOnly
		do while not rs.eof
			if rs("Preinstall") then
			 strOldOSPreinstallList = strOldOSPreinstallList & ", " & rs("shortname")
			end if
			if rs("Web") then
			 strOldOSWebList = strOldOSWebList & ", " & rs("shortname")
			end if
			rs.movenext
		loop
		rs.close
		
		OSArray = split(request("txtFullOSList"),",")
		for i = lbound(OSArray) to ubound(OSArray)
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spUpdateProductOSList"
			
			Set p = cm.CreateParameter("@ProductID", 3, &H0001)
			p.Value = strID
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@OSID", 3, &H0001)
			p.Value = OSArray(i)
			cm.Parameters.Append p
	
			Set p = cm.CreateParameter("@Preinstall", 11, &H0001)
			if instr(", " & request("chkPreinstallOS") & ",",", " & trim(OSArray(i)) & ",")>0 then
				p.Value = 1
			else
				p.Value = 0
			end if
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Web", 11, &H0001)
			if instr(", " & request("chkWebOS") & ",",", " & trim(OSArray(i)) & ",")>0 then
				p.Value = 1
			else
				p.Value = 0
			end if
			cm.Parameters.Append p
	
			cm.execute
			set cm = nothing
			
			if cn.Errors.count <> 0 then
				blnFailed = true
				exit for
			end if		
		next
		


		dim strNewOSPreinstallList
		dim strNewOSWebList
		strNewOSPreinstallList=""
		strNewOSWebList=""
		rs.open "spListProductOS " & strID,cn,adOpenForwardOnly
		do while not rs.eof
			if rs("Preinstall") then
			 strNewOSPreinstallList = strNewOSPreinstallList & ", " & rs("shortname")
			end if
			if rs("Web") then
			 strNewOSWebList = strNewOSWebList & ", " & rs("shortname")
			end if
			rs.movenext
		loop
		rs.close		
		
		if strNewOSPreinstallList <> "" then
			strNewOSPreinstallList = mid(strNewOSPreinstallList,3)
		end if
		if strNewOSWebList <> "" then
			strNewOSWebList = mid(strNewOSWebList,3)
		end if
		if strOldOSPreinstallList <> "" then
			strOldOSPreinstallList = mid(strOldOSPreinstallList,3)
		end if
		if strOldOSWebList <> "" then
			strOldOSWebList = mid(strOldOSWebList,3)
		end if
		
		dim strOSChangedBody
		strOSChangedBody = "<font face=Arial size=2 color=black>The " & request("txtProductFamily") & " " & request("txtVersion") & " OS list has been changed:<BR><BR>"
		if strNewOSPreinstallList <> strOldOSPreinstallList then
			strOSChangedBody=strOSChangedBody & "<b><u>Preinstall</u></b><BR><b>Old:</b> " & strOldOSPreinstallList & "<BR><b>New:</b> " & strNewOSPreinstallList & "<BR><BR>"
		end if
		if strNewOSWebList <> strOldOSWebList then
			strOSChangedBody=strOSChangedBody &  "<b><u>Web</u></b><BR><b>Old:</b> " & strOldOSWebList & "<BR><b>New:</b> " & strNewOSWebList & "<br><br>"
		end if
		
		if strNewOSPreinstallList <> strOldOSPreinstallList or strNewOSWebList <> strOldOSWebList then
			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			if trim(strID) = "100" then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= "houreleasecoordinatorsswrel@hp.com"
			end if
            oMessage.Subject = "Product OS List Updated in Pulsar" 
	
			oMessage.HTMLBody = strOSChangedBody & "</font>"
			
			oMessage.Send 
			Set oMessage = Nothing 			
		end if
	end if
	
	
	'Close all approved DCRs if changing from any status to Sustaining 
	if (bAddingProduct = false and isCloning = false) and trim(request("tagPhase")) <> "3" and trim(request("cboPhase")) = "3" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spCloseApprovedDCRs"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = strID
		cm.Parameters.Append p

		cm.Execute RowsEffected
		Set cm=nothing

		'cn.Execute "spCloseApprovedDCRs " & request("txtID"), RowsEffected
		if cn.Errors.count <> 0 then
			blnFailed = true
		end if		
	end if
	
    If Not blnFailed And bAddingProduct Then
        set cm = server.CreateObject("ADODB.Command")
        Set cm.ActiveConnection = cn
        cm.CommandType = 4
        cm.CommandText = "spAddRelease2Product"
			
        Set p = cm.CreateParameter("@ProductID", 3, &H0001)
        p.Value = strID
        cm.Parameters.Append p
        
        cm.execute

        set cm = nothing

        if cn.Errors.count <> 0 then
            blnFailed = true
        end if		
    End If

	Response.Write "<BR>"
	
	dim strBrandsAdded
	strBrandsAdded = ""
	
	if not blnfailed then
	    'Rename Brand
	    if trim(request("txtBrandTo")) <> "" and trim(request("txtBrandFrom")) <> "" then
            dim BrandFromArray
            dim BrandToArray
            
            BrandToArray = split(request("txtBrandTo"),",")
            BrandFromArray = split(request("txtBrandFrom"),",")
            
            'Update the Brand Assignment
            for i = 0 to ubound(BrandToArray)
                if trim(BrandToArray(i)) <> "" and trim(BrandFromArray(i)) <> "" then
                    'Look for series changes
					strSeriesTag1 = trim(request("tagSeriesA" & trim(clng(BrandFromArray(i)))))
					strSeriesTag2 = trim(request("tagSeriesB" & trim(clng(BrandFromArray(i)))))
					strSeriesTag3 = trim(request("tagSeriesC" & trim(clng(BrandFromArray(i)))))
					strSeriesTag4 = trim(request("tagSeriesD" & trim(clng(BrandFromArray(i)))))
					strSeriesTag5 = trim(request("tagSeriesE" & trim(clng(BrandFromArray(i)))))
					strSeriesTag6 = trim(request("tagSeriesF" & trim(clng(BrandFromArray(i)))))

					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandToArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandToArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandToArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandToArray(i)))))
					strSeries5 = trim(request("txtSeriesE" & trim(clng(BrandToArray(i)))))
					strSeries6 = trim(request("txtSeriesF" & trim(clng(BrandToArray(i)))))

					strSeriesID1 = trim(request("txtSeriesIDA" & trim(clng(BrandFromArray(i)))))
					strSeriesID2 = trim(request("txtSeriesIDB" & trim(clng(BrandFromArray(i)))))
					strSeriesID3 = trim(request("txtSeriesIDC" & trim(clng(BrandFromArray(i)))))	        
					strSeriesID4 = trim(request("txtSeriesIDD" & trim(clng(BrandFromArray(i)))))	        
					strSeriesID5 = trim(request("txtSeriesIDE" & trim(clng(BrandFromArray(i)))))	        
					strSeriesID6 = trim(request("txtSeriesIDF" & trim(clng(BrandFromArray(i)))))	        

					if strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries1
						strSeries1 = strSeriesTemp
					end if

					if strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries2
						strSeries2 = strSeriesTemp
					end if

					if strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries3
						strSeries3 = strSeriesTemp
					end if

					if strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries4
						strSeries4 = strSeriesTemp
					end if


					if strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries5
						strSeries5 = strSeriesTemp
					end if

					if strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries6
						strSeries6 = strSeriesTemp
					end if

					SeriesTagArray = split(strSeriesTag1 & chr(1) & strSeriesTag2 & chr(1) & strSeriesTag3 & chr(1) & strSeriesTag4 & chr(1) & strSeriesTag5 & chr(1) & strSeriesTag6,chr(1))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4 & chr(1) & strSeries5 & chr(1) & strSeries6,chr(1))
					SeriesIDArray = split(strSeriesID1 & chr(1) & strSeriesID2 & chr(1) & strSeriesID3 & chr(1) & strSeriesID4 & chr(1) & strSeriesID5 & chr(1) & strSeriesID6,chr(1))

					'Update Series Summary
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spUpdateSeriesSummary"
					
					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = clng(strID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@BrandID", 3, &H0001)
					p.Value = trim(clng(BrandFromArray(i)))
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@SeriesSummary", 200, &H0001,2000)
					strOutput = ""
					for j = lbound(SeriesArray) to ubound(SeriesArray)
						if trim(SeriesArray(j)) <> "" then
							strOutput = strOutput & "," & SeriesArray(j)
						end if
					next
					if len(strOutput) > 0 then
						strOutput = mid(strOutput,2)
					end if
					p.value = strOutput
					cm.Parameters.Append p
		
					cm.execute
					set cm = nothing
					if cn.Errors.count <> 0 then
						blnFailed = true
						exit for
					end if		

'------------------------------------------
                    strSeriesString = ""
                    if not blnFailed then
		                'Save/Update Series Records
		                for j = 0 to ubound(SeriesArray)
			                if trim(SeriesTagArray(j)) <> trim(SeriesArray(j)) then
				                set cm = server.CreateObject("ADODB.Command")
				                Set cm.ActiveConnection = cn
				                cm.CommandType = 4
				                if trim(SeriesIDArray(j)) = "" then
					                if trim(SeriesArray(j)) <> "" then
						                cm.CommandText = "spAddSeries2Brand2"
		
				                        Set p = cm.CreateParameter("@ID", 3, &H0001)
				                        p.Value = clng(strID)
				                        cm.Parameters.Append p

						                Set p = cm.CreateParameter("@BrandID", 3, &H0001)
						                p.Value = clng(BrandFromArray(i))
						                cm.Parameters.Append p

						                Set p = cm.CreateParameter("@Name", 200, &H0001,50)
						                p.Value = left(SeriesArray(j),50)
						                cm.Parameters.Append p

						                Set p = cm.CreateParameter("@NewID", 3, &H0002)
						                cm.Parameters.Append p
                						
					                end if
				                else
					                if left(SeriesArray(j),50) = "" then
						                'Create Notification Text for series Removed
						                set rs = server.CreateObject("ADODB.Recordset")
						                rs.open "spGetBrandSeries3 " & clng(SeriesIDArray(j)),cn,adOpenForwardOnly
						                strSeriesString = strSeriesString & BuildSeriesEmail (rs, 2)
						                rs.close
						                set	rs = nothing
                					else
						                'Create Notification Text for series Updated
						                set rs = server.CreateObject("ADODB.Recordset")
						                rs.open "spGetBrandSeries4 " & clng(SeriesIDArray(j)) & ",'" & scrubsql(left(SeriesArray(j),50)) & "'," & clng(BrandFromArray(i)) & "," & clng(BrandToArray(i)) ,cn,adOpenForwardOnly
						                strSeriesString = strSeriesString & BuildSeriesEmail (rs, 4)
						                rs.close
						                set	rs = nothing
					                end if
												
					                cm.CommandText = "spUpdateSeries"
                		
					                Set p = cm.CreateParameter("@ID", 3, &H0001)
					                p.Value = clng(SeriesIDArray(j))
					                cm.Parameters.Append p

					                Set p = cm.CreateParameter("@Name", 200, &H0001,50)
					                p.Value = left(SeriesArray(j),50)
					                cm.Parameters.Append p
				                end if

				                cm.execute
				                if cn.Errors.count <> 0 then
					                blnFailed = true
					                exit for
				                end if
				
				                if trim(SeriesIDArray(j)) = "" then
					                'Create Notification Text for series Added
					                set rs = server.CreateObject("ADODB.Recordset")
					                rs.open "spGetBrandSeries4 " & clng(cm("@NewID")) & ",'" & scrubsql(left(SeriesArray(j),50)) & "'," & clng(BrandFromArray(i)) & "," & clng(BrandToArray(i)) ,cn,adOpenForwardOnly
					                strSeriesString = strSeriesString & BuildSeriesEmail (rs, 5)
					                rs.close
					                set	rs = nothing
				                end if
				
                				set cm = nothing
						
						    elseif trim(SeriesIDArray(j)) <> "" and trim(SeriesArray(j)) <> "" then
				                'Create Notification Text for Brand Updated - Series Didn't change but the brand did so it is still an update
    			                set rs = server.CreateObject("ADODB.Recordset")
	    		                rs.open "spGetBrandSeries4 " & clng(SeriesIDArray(j)) & ",'" & scrubsql(left(SeriesArray(j),50)) & "'," & clng(BrandFromArray(i)) & "," & clng(BrandToArray(i)) ,cn,adOpenForwardOnly
				                strSeriesString = strSeriesString & BuildSeriesEmail (rs, 4)
				                rs.close
				                set	rs = nothing
			                end if
		                next
                    end if



'------------------------------------------
                    'Update Brand Link
	                set cm = server.CreateObject("ADODB.Command")
	                Set cm.ActiveConnection = cn
	                cm.CommandType = 4
	                cm.CommandText = " spUpdateProductBrand"
                					
	                Set p = cm.CreateParameter("@ProductID", 3, &H0001)
	                p.Value = clng(strID)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@OldBrandID", 3, &H0001)
	                p.Value = clng(BrandFromArray(i))
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@NewBrandID", 3, &H0001)
	                p.Value = clng(BrandToArray(i))
	                cm.Parameters.Append p

	                cm.execute RowsUpdated
	                set cm = nothing
	                if RowsUpdated <> 1 then
		                blnFailed = true
		                exit for
	                end if	
                end if
	        next
	        
	    end if
	    
		'Add Brands
		Response.Write "<BR>ADD:"
		BrandArray = split(request("chkBrands"),",")
		for i = 0 to ubound(BrandArray)
			if trim(Brandarray(i)) <> "" then
				if instr("," & request("txtBrandsLoaded") & ",", "," & trim(BrandArray(i)) & ",") = 0 and instr("," & replace(request("txtBrandTo")," ","") & ",", "," & trim(BrandArray(i)) & ",") = 0 then
					Response.Write trim(BrandArray(i)) & "<BR>"
					
					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandArray(i)))))
					strSeries5 = trim(request("txtSeriesE" & trim(clng(BrandArray(i)))))
					strSeries6 = trim(request("txtSeriesF" & trim(clng(BrandArray(i)))))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4 & chr(1) & strSeries5 & chr(1) & strSeries6,chr(1))

					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spAddBrand2Product"
			
					Set p = cm.CreateParameter("@ProductID", 3, &H0001)
					p.Value = strID
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@BrandID", 3, &H0001)
					p.Value = clng(BrandArray(i))
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@SeriesSummary", 200, &H0001,2000)
					strOutput = ""
					for j = lbound(SeriesArray) to ubound(SeriesArray)
						if trim(SeriesArray(j)) <> "" then
							strOutput = strOutput & "," & SeriesArray(j)
						end if
					next
					if len(strOutput) > 0 then
						strOutput = mid(strOutput,2)
					end if
					
					p.value = strOutput
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@ProductBrandID", 3, &H0002)
					cm.Parameters.Append p

					cm.execute
					ProductBrandID = cm("@ProductBrandID")
					set cm = nothing
					if cn.Errors.count <> 0 then
						blnFailed = true
						exit for
					end if		

					strBrandsAdded = strBrandsAdded  & "," & ProductBrandID	

					For Each SeriesName in SeriesArray	
						if trim(SeriesName) <> "" then
							set cm = server.CreateObject("ADODB.Command")
							Set cm.ActiveConnection = cn
							cm.CommandType = 4
							cm.CommandText = "spAddSeries2Brand"
				
							Set p = cm.CreateParameter("@BrandID", 3, &H0001)
							p.Value = ProductBrandID
							cm.Parameters.Append p
	
							Set p = cm.CreateParameter("@Name", 200, &H0001,50)
							p.Value = left(trim(SeriesName),50)
							cm.Parameters.Append p
		
							cm.execute
							set cm = nothing
							if cn.Errors.count <> 0 then
								blnFailed = true
								exit for
							end if	
						end if
					next

					
				end if
			end if
		next

        dim errloop
		'Remove/Update Brands
		BrandArray = split(request("txtBrandsLoaded"),",")
		Response.Write "<BR>REMOVE:"
		for i = 0 to ubound(BrandArray)
			if trim(Brandarray(i)) <> "" then
				if instr(", " & request("chkBrands") & ",", ", " & trim(BrandArray(i)) & ",") = 0 and instr("," & replace(request("txtBrandFrom")," ","") & ",", "," & trim(BrandArray(i)) & ",") = 0 then
					Response.Write trim(BrandArray(i)) & "<BR>"
					'Create Notification Text for Whole Brand Removed
					set rs = server.CreateObject("ADODB.Recordset")
					rs.open "spGetBrandSeries2 " & clng(strID) & "," &  clng(BrandArray(i)),cn,adOpenForwardOnly
					strSeriesString = strSeriesString & BuildSeriesEmail (rs, 2)
					rs.close
					set rs = nothing

					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spRemoveBrandFromProduct"
			
					Set p = cm.CreateParameter("@ProductID", 3, &H0001)
					p.Value = strID
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@BrandID", 3, &H0001)
					p.Value = clng(BrandArray(i))
					cm.Parameters.Append p
	
					cm.execute
					set cm = nothing
					if cn.Errors.count <> 0 then
                        for each errloop in cn.Errors
                            response.write errloop.number & "<BR>"
                            response.write errloop.description & "<BR>"
                            response.write errloop.source & "<BR>"
                            response.write errloop.sqlstate & "<BR>"
                            response.write errloop.nativeerror & "<BR>"
                        next
						blnFailed = true
						exit for
					end if		
				else'if instr("," & replace(request("txtBrandFrom")," ","") & ",", "," & trim(BrandArray(i)) & ",") = 0 then
					strSeriesTag1 = trim(request("tagSeriesA" & trim(clng(BrandArray(i)))))
					strSeriesTag2 = trim(request("tagSeriesB" & trim(clng(BrandArray(i)))))
					strSeriesTag3 = trim(request("tagSeriesC" & trim(clng(BrandArray(i)))))
					strSeriesTag4 = trim(request("tagSeriesD" & trim(clng(BrandArray(i)))))
					strSeriesTag5 = trim(request("tagSeriesE" & trim(clng(BrandArray(i)))))
					strSeriesTag6 = trim(request("tagSeriesF" & trim(clng(BrandArray(i)))))

					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandArray(i)))))
					strSeries5 = trim(request("txtSeriesE" & trim(clng(BrandArray(i)))))
					strSeries6 = trim(request("txtSeriesF" & trim(clng(BrandArray(i)))))

					strSeriesID1 = trim(request("txtSeriesIDA" & trim(clng(BrandArray(i)))))
					strSeriesID2 = trim(request("txtSeriesIDB" & trim(clng(BrandArray(i)))))
					strSeriesID3 = trim(request("txtSeriesIDC" & trim(clng(BrandArray(i)))))
					strSeriesID4 = trim(request("txtSeriesIDD" & trim(clng(BrandArray(i)))))
					strSeriesID5 = trim(request("txtSeriesIDE" & trim(clng(BrandArray(i)))))
					strSeriesID6 = trim(request("txtSeriesIDF" & trim(clng(BrandArray(i)))))

					
					if strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries1
						strSeries1 = strSeriesTemp
					elseif strSeriesTag1 <> "" and strSeriesTag1 <> strSeries1 and strSeriesTag1 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries1
						strSeries1 = strSeriesTemp
					end if

					if strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries2
						strSeries2 = strSeriesTemp
					elseif strSeriesTag2 <> "" and strSeriesTag2 <> strSeries2 and strSeriesTag2 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries2
						strSeries2 = strSeriesTemp
					end if

					if strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries3
						strSeries3 = strSeriesTemp
					elseif strSeriesTag3 <> "" and strSeriesTag3 <> strSeries3 and strSeriesTag3 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries3
						strSeries3 = strSeriesTemp
					end if

					if strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries4
						strSeries4 = strSeriesTemp
					elseif strSeriesTag4 <> "" and strSeriesTag4 <> strSeries4 and strSeriesTag4 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries4
						strSeries4 = strSeriesTemp
					end if

					if strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries5
						strSeries5 = strSeriesTemp
					elseif strSeriesTag5 <> "" and strSeriesTag5 <> strSeries5 and strSeriesTag5 = strSeries6 then
						strSeriesTemp = strSeries6
						strSeries6 = strSeries5
						strSeries5 = strSeriesTemp
					end if

					if strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries1 then
						strSeriesTemp = strSeries1
						strSeries1 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries2 then
						strSeriesTemp = strSeries2
						strSeries2 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries3 then
						strSeriesTemp = strSeries3
						strSeries3 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries4 then
						strSeriesTemp = strSeries4
						strSeries4 = strSeries6
						strSeries6 = strSeriesTemp
					elseif strSeriesTag6 <> "" and strSeriesTag6 <> strSeries6 and strSeriesTag6 = strSeries5 then
						strSeriesTemp = strSeries5
						strSeries5 = strSeries6
						strSeries6 = strSeriesTemp
					end if

					SeriesTagArray = split(strSeriesTag1 & chr(1) & strSeriesTag2 & chr(1) & strSeriesTag3 & chr(1) & strSeriesTag4 & chr(1) & strSeriesTag5 & chr(1) & strSeriesTag6,chr(1))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4 & chr(1) & strSeries5 & chr(1) & strSeries6,chr(1))
					SeriesIDArray = split(strSeriesID1 & chr(1) & strSeriesID2 & chr(1) & strSeriesID3 & chr(1) & strSeriesID4 & chr(1) & strSeriesID5 & chr(1) & strSeriesID6,chr(1))

					'Update Series Summary
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spUpdateSeriesSummary"
					
					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = clng(strID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = trim(clng(BrandArray(i)))
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@SeriesSummary", 200, &H0001,2000)
					strOutput = ""
					for j = lbound(SeriesArray) to ubound(SeriesArray)
						if trim(SeriesArray(j)) <> "" then
							strOutput = strOutput & "," & SeriesArray(j)
						end if
					next
					if len(strOutput) > 0 then
						strOutput = mid(strOutput,2)
					end if
					p.value = strOutput
					cm.Parameters.Append p
		
					cm.execute
					set cm = nothing
					if cn.Errors.count <> 0 then
						blnFailed = true
						exit for
					end if		


					'Save/Update Series Records
					for j = 0 to ubound(SeriesArray)
						if trim(SeriesTagArray(j)) <> trim(SeriesArray(j)) then
							set cm = server.CreateObject("ADODB.Command")
							Set cm.ActiveConnection = cn
							cm.CommandType = 4
							if trim(SeriesIDArray(j)) = "" then
								if trim(SeriesArray(j)) <> "" then
									cm.CommandText = "spAddSeries2Brand2"
					
									Set p = cm.CreateParameter("@ID", 3, &H0001)
									p.Value = clng(strID)
									cm.Parameters.Append p
	
									Set p = cm.CreateParameter("@ID", 3, &H0001)
									p.Value = trim(clng(BrandArray(i)))
									cm.Parameters.Append p
		
									Set p = cm.CreateParameter("@Name", 200, &H0001,50)
									p.Value = left(SeriesArray(j),50)
									cm.Parameters.Append p

									Set p = cm.CreateParameter("@NewID", 3, &H0002)
									cm.Parameters.Append p
									
								end if
							else
								if left(SeriesArray(j),50) = "" then
									'Create Notification Text for series Removed
									set rs = server.CreateObject("ADODB.Recordset")
									rs.open "spGetBrandSeries3 " & clng(SeriesIDArray(j)),cn,adOpenForwardOnly
									strSeriesString = strSeriesString & BuildSeriesEmail (rs, 2)
									rs.close
									set	rs = nothing
								else
									'Create Notification Text for series Updated
									set rs = server.CreateObject("ADODB.Recordset")
									rs.open "spGetBrandSeries3 " & clng(SeriesIDArray(j)) & ",'" & scrubsql(left(SeriesArray(j),50)) & "'" ,cn,adOpenForwardOnly
									strSeriesString = strSeriesString & BuildSeriesEmail (rs, 3)
									rs.close
									set	rs = nothing
								end if
															
								cm.CommandText = "spUpdateSeries"
					
								Set p = cm.CreateParameter("@ID", 3, &H0001)
								p.Value = clng(SeriesIDArray(j))
								cm.Parameters.Append p
	
								Set p = cm.CreateParameter("@Name", 200, &H0001,50)
								p.Value = left(SeriesArray(j),50)
								cm.Parameters.Append p
							end if
		
							cm.execute
							if cn.Errors.count <> 0 then
								blnFailed = true
								exit for
							end if
							
							if trim(SeriesIDArray(j)) = "" then
								'Create Notification Text for series Added
								set rs = server.CreateObject("ADODB.Recordset")
								rs.open "spGetBrandSeries3 " & clng(cm("@NewID")),cn,adOpenForwardOnly
								strSeriesString = strSeriesString & BuildSeriesEmail (rs, 1)
								rs.close
								set	rs = nothing
							end if
							
							set cm = nothing
									
						end if
					next
					
				end if
			end if
			
		next
		
        'this is the end of updating all the brand related info, 
        'call the sp to update the new brandnames based on formula define as that sp assumes the series, product_brand records are already there
        ' also update  generation, form factor , scmnumber
        'Yong changes for SCM, will uncomment out later
        BrandArray = split(request("chkBrands"),",")
		for i = 0 to ubound(BrandArray)
       
                dim strmodelnumber
                dim strgeneration
                dim strFormFactor
                dim SCMNumber
                set cm = server.CreateObject("ADODB.Command")
			    Set cm.ActiveConnection = cn
			    cm.CommandType = 4
			    cm.CommandText = "usp_Product_UpdateBrandNames"				
			
	            Set p = cm.CreateParameter("@ProductversionID", 3, &H0001)
				p.Value = strID
				cm.Parameters.Append p
        
			    Set p = cm.CreateParameter("@BrandID", 3, &H0001)
			    p.Value = clng(BrandArray(i))
			    cm.Parameters.Append p

                strmodelnumber = ""
			    Set p = cm.CreateParameter("@ModelNumber", 200, &H0001, 10)
			    p.Value = strmodelnumber
			    cm.Parameters.Append p
                
                Set p = cm.CreateParameter("@ScreenSize",131, &H0001)
				p.Precision = 5
				p.NumericScale = 2
				cm.Parameters.Append p

                strgeneration = trim(request("txtGeneration" & trim(clng(Brandarray(i)))))
                  response.Write (strgeneration)
			    Set p = cm.CreateParameter("@Generation", 200, &H0001, 5)
			    p.Value = strgeneration
			    cm.Parameters.Append p
			
                strFormFactor = trim(request("txtFormFactor" & trim(clng(Brandarray(i)))))
			    Set p = cm.CreateParameter("@FormFactor", 200, &H0001,15)
			    p.Value = strFormFactor
			    cm.Parameters.Append p    

			    Set p = cm.CreateParameter("@SCMNumber", 3, &H0001) 
			    p.Value = 0 ' no SCM for legacy products
			    cm.Parameters.Append p                
         
		
			    cm.execute
			    set cm = nothing
			    if cn.Errors.count <> 0 then
				    blnFailed = true
				    exit for
			    end if	
        next	

	end if
	Response.Write "<BR>"
	
	ConsumerArray = Split(request("chkSystemTeamConsumer"),",")
	CommercialArray = Split(request("chkSystemTeamCommercial"),",")
	SMBArray = Split(request("chkSystemTeamSMB"),",")
	for i = lbound(ConsumerArray) to ubound(ConsumerArray)
		if instr(strApprovers & ",","," & trim(ConsumerArray(i)) & ",") = 0 then
			strApprovers = strApprovers & "," & trim(ConsumerArray(i))
		end if
	next

	for i = lbound(CommercialArray) to ubound(CommercialArray)
		if instr(strApprovers & ",","," & trim(CommercialArray(i)) & ",") = 0 then
			strApprovers = strApprovers & "," & trim(CommercialArray(i))
		end if
	next

	for i = lbound(SMBArray) to ubound(SMBArray)
		if instr(strApprovers & ",","," & trim(SMBArray(i)) & ",") = 0 then
			strApprovers = strApprovers & "," & trim(SMBArray(i))
		end if
	next

	ConsumerTagArray = Split(request("tagSystemTeamConsumer"),",")
	CommercialTagArray = Split(request("tagSystemTeamCommercial"),",")
	SMBTagArray = Split(request("tagSystemTeamSMB"),",")
	
	for i = lbound(ConsumerTagArray) to ubound(ConsumerTagArray)
		if instr(strApproversTag & ",","," & trim(ConsumerTagArray(i)) & ",") = 0 then
			strApproversTag = strApproversTag & "," & trim(ConsumerTagArray(i))
		end if
	next

	for i = lbound(CommercialTagArray) to ubound(CommercialTagArray)
		if instr(strApproversTag & ",","," & trim(CommercialTagArray(i)) & ",") = 0 then
			strApproversTag = strApproversTag & "," & trim(CommercialTagArray(i))
		end if
	next

	for i = lbound(SMBTagArray) to ubound(SMBTagArray)
		if instr(strApproversTag & ",","," & trim(SMBTagArray(i)) & ",") = 0 then
			strApproversTag = strApproversTag & "," & trim(SMBTagArray(i))
		end if
	next
	
	ApproversArray =Split(strApprovers,",") 
	ApproversTagArray =Split(strApproversTag,",") 
	
	'Add System Team Members
if false then
	for i = lbound(ApproversArray) to ubound(ApproversArray)
		if trim(ApproversArray(i)) <> "" and instr(strApproversTag & ",","," & trim(ApproversArray(i)) & ",")= 0 then

			set cm = nothing
			set cm = server.CreateObject("ADODB.command")

			cm.ActiveConnection = cn
			cm.CommandText = "spAddSystemTeamMember"
			cm.CommandType =  &H0004


			set p =  cm.CreateParameter("@ProdID", 3, &H0001)
			p.value = clng(strID)
			cm.Parameters.Append p
			set p =  cm.CreateParameter("@EmployeeID", 3, &H0001)
			p.value = clng(trim(ApproversArray(i)))
			cm.Parameters.Append p
	
			set p =  cm.CreateParameter("@Consumer", 11, &H0001)
			if instr(", " & request("chkSystemTeamConsumer") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p

			set p =  cm.CreateParameter("@Commercial", 11, &H0001)
			if instr(", " & request("chkSystemTeamCommercial") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p

			set p =  cm.CreateParameter("@SMB", 11, &H0001)
			if instr(", " & request("chkSystemTeamSMB") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p
	
			cm.Execute RowsEffected

			if rowseffected <> 1 then
				blnFailed = true
				exit for
			end if
			Response.Write "Add " & ApproversArray(i) & "<BR>"
		end if
	next

	
	'Remove System Team Members
	for i = lbound(ApproversTagArray) to ubound(ApproversTagArray)
		if trim(ApproversTagArray(i)) <> "" and instr(strApprovers & ",","," & trim(ApproversTagArray(i)) & ",")= 0 then

			set cm = nothing
			set cm = server.CreateObject("ADODB.command")

			cm.ActiveConnection = cn
			cm.CommandText = "spRemoveSystemTeamMember"
			cm.CommandType =  &H0004


			set p =  cm.CreateParameter("@ProdID", 3, &H0001)
			p.value = clng(strID)
			cm.Parameters.Append p
			
			set p =  cm.CreateParameter("@EmployeeID", 3, &H0001)
			p.value = clng(trim(ApproversTagArray(i)))
			cm.Parameters.Append p
	
	
			cm.Execute RowsEffected

			if rowseffected <> 1 then
				blnFailed = true
				exit for
			end if
			Response.Write "Remove " & ApproversTagArray(i) & "<BR>"
		end if
	next
	
	

	'Update Member Properties
	for i = lbound(ApproversArray) to ubound(ApproversArray)
		if trim(ApproversArray(i)) <> "" and instr(strApproversTag & ",","," & trim(ApproversArray(i)) & ",")> 0 then

			set cm = nothing
			set cm = server.CreateObject("ADODB.command")

			cm.ActiveConnection = cn
			cm.CommandText = "spUpdateSystemTeamMember"
			cm.CommandType =  &H0004


			set p =  cm.CreateParameter("@ProdID", 3, &H0001)
			p.value = clng(strID)
			cm.Parameters.Append p
			
			set p =  cm.CreateParameter("@EmployeeID", 3, &H0001)
			p.value = clng(trim(ApproversArray(i)))
			cm.Parameters.Append p
	
			set p =  cm.CreateParameter("@Consumer", 11, &H0001)
			if instr(", " & request("chkSystemTeamConsumer") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p

			set p =  cm.CreateParameter("@Commercial", 11, &H0001)
			if instr(", " & request("chkSystemTeamCommercial") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p

			set p =  cm.CreateParameter("@SMB", 11, &H0001)
			if instr(", " & request("chkSystemTeamSMB") & ",",", " & trim(ApproversArray(i)) & "," ) > 0 then
				p.value = 1
			else
				p.value = 0
			end if
			cm.Parameters.Append p
	
			cm.Execute RowsEffected

			if rowseffected <> 1 then
				blnFailed = true
				exit for
			end if
			Response.Write "Update " & ApproversArray(i) & "<BR>"
		end if
	next
	
end if	
	if cn.Errors.Count > 1 or blnFailed then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		Response.Write "<font size=2 face=verdana><b>Unable to save this product.</b></font>"
		cn.RollbackTrans
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""" & strID & """>"
		cn.CommitTrans
	end if

	if not blnFailed then
		'Notify people that the product name was changed if needed
		if (bAddingProduct = false and isCloning = false) and trim(request("txtID")) <> "100" and request("tagVersion") <> request("txtVersion") then			
			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			oMessage.To= "mobileexcalnewproducts@hp.com"
			oMessage.Subject = "Product Renamed in Pulsar" 

			oMessage.HTMLBody = "<font face=Arial size=2 color=black>Renamed By: " & CurrentUser & "<BR>Old Product Name: " & request("txtProductFamily") & " " & request("tagVersion") & "<BR>New Product Name: " & request("txtProductName") & "</font>"
		
			oMessage.Send 
			Set oMessage = Nothing 			
			
			
		end if
		
		if trim(request("cboPhase")) <> trim(request("tagPhase")) then
			dim strFromPhase
			dim strToPhase
					
			select case trim(request("cboPhase"))
				case "1"
					strToPhase="Definition"
				case "2"
					strToPhase="Development"
				case "3"
					strToPhase="Production"
				case "4"
					strToPhase="Post-Production"
				case "5"
					strToPhase="Inactive"
				case else
					strToPhase=""
			end select
		
			select case trim(request("tagPhase"))
				case "1"
					strFromPhase="Definition"
				case "2"
					strFromPhase="Development"
				case "3"
					strFromPhase="Production"
				case "4"
					strFromPhase="Post-Production"
				case "5"
					strFromPhase="Inactive"
				case else
					strFromPhase=""
			end select

			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			oMessage.To= "max.yu@hp.com"
			if trim(request("txtID")) <> "100" then
				oMessage.CC = "MobileExcalNotification-ProductNames@hp.com"
			end if
			if CurrentUserFirstName <> "" then
				CurrentUserFirstName = "<BR><BR>Thanks,<BR>" & CurrentUserFirstName 
			end if

			if trim(strFromPhase) = "" and trim(strToPhase) = "" then
				oMessage.Subject = request("txtProductFamily") & " " & request("txtVersion") & " status has been changed" 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductFamily") & " " & request("txtVersion") & " status has been changed in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			elseif trim(strFromPhase) = "" then
				oMessage.Subject = request("txtProductFamily") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductFamily") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) & " in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			else	
				oMessage.Subject = request("txtProductFamily") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductFamily") & " " & request("txtVersion") & " status has been changed from " & trim(strFromPhase) & " to " & trim(strToPhase) & " in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			end if
			oMessage.Importance = cdoHigh
		
			oMessage.Send 
			Set oMessage = Nothing 			
		
		end if
		
	
		if (bAddingProduct = false and isCloning = false) and trim(request("tagPhase")) <> trim(request("cboPhase")) and trim(request("tagPhase"))= "5" and trim(request("cboPhase"))<> "5" then			
			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			oMessage.To= "max.yu@hp.com;"
			if trim(request("txtID")) <> "100" then
				oMessage.CC = "tammy.schapiro@hp.com;rick.rostonski@hp.com" 
			end if
			oMessage.Subject = request("txtProductFamily") & " " & request("txtVersion") & " Reactivated in Pulsar" 

			oMessage.HTMLBody = "<font face=Arial size=2 color=black>Reactivated By: " & CurrentUser & "</font>"
			oMessage.Send 
			Set oMessage = Nothing 			
		end if
		
		
		
	end if

    dim strPartnerName
    if trim(request("tagPartnerID")) <> "0" and trim(request("tagPartnerID")) <> trim(request("cboPartner")) and trim(request("cboPartner")) <> "" and trim(request("cboPartner")) <> "0" and trim(request("tagPartnerID")) <> "" and request("txtID") <> "" then
    
	    
		Response.Write "<BR><BR>ODM Updated.  Sending notifications."
	    
	    'Lookup ODM
	    rs.open "spGetPartnerName " & clng(request("cboPartner")),cn,adOpenForwardOnly
	    if rs.eof and rs.bof then
	        strPartnerName = "to Not Specified"
	    elseif trim(rs("Name") & "") = "" then
	        strPartnerName = "to Not Specified"
	    else
	        strPartnerName = "to " & trim(rs("Name") & "")
	    end if
        rs.close

	    rs.open "spGetPartnerName " & clng(request("tagPartnerID")),cn,adOpenForwardOnly
	    if rs.eof and rs.bof then
	        strPartnerName = "<BR>ODM:  Changed from Not Specified " & strPartnerName & "."
	    elseif trim(rs("Name") & "") = "" then
	        strPartnerName = "<BR>ODM:  Changed from Not Specified " & strPartnerName & "."
	    else
	        strPartnerName = "<BR>ODM:  Changed from " & trim(rs("Name") & "") & " " & strPartnerName & "."
	    end if
        rs.close
        
        if lcase(Currentuseremail) = "max.yu@hp.com" then
            strPartnerName  = strPartnerName & "<BR><BR>Updated By: " & CurrentUser
        end if
        	
       
		'ODM Updated
		Set oMessage = New EmailWrapper
		oMessage.From = Currentuseremail
		oMessage.To= "max.yu@hp.com"
		oMessage.Subject = "Product ODM Updated" 

		oMessage.HTMLBody = "<font face=Arial size=2 color=black>Product Name: " & request("txtProductName") & strpartnername & "</font>"
		
		oMessage.Send 
		Set oMessage = Nothing 	
    
    end if

	if bAddingProduct = true or isCloning = true then
	    dim strSEPM
	    
		Response.Write "<BR><BR>Product Created.  Sending notifications."
	    
	    'Lookup ODM
	    rs.open "spGetPartnerName " & clng(request("cboPartner")),cn,adOpenForwardOnly
	    if rs.eof and rs.bof then
	        strPartnerName = "<BR>ODM: Not Specified"
	    elseif trim(rs("Name") & "") = "" then
	        strPartnerName = "<BR>ODM: Not Specified"
	    else
	        strPartnerName = "<BR>ODM: " & trim(rs("Name") & "")
	    end if
        rs.close
        
        if lcase(Currentuseremail) = "max.yu@hp.com" then
            strPartnerName  = strPartnerName & "<BR><BR>Added By: " & CurrentUser
        end if
        	
        'Lookup SEPM
        if trim(request("cboSEPM")) <> "" and isnumeric(request("cboSEPM")) then
    	    rs.open "spGetEmployeeByID " & clng(request("cboSEPM")),cn,adOpenForwardOnly
	        if rs.eof and rs.bof then
	            strSEPM = "<BR>SE PM: Not Specified"
	        elseif trim(rs("Name") & "") = "" then
	            strSEPM = "<BR>SE PM: Not Specified"
	        else
	            strSEPM = "<BR>SE PM: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	        end if
            rs.close
        else
            strSEPM = "<BR>SE PM: Not Specified"
        end if

        'Lookup SM
        dim strSM
        if trim(request("cboSM")) <> "" and isnumeric(request("cboSM")) then
    	    rs.open "spGetEmployeeByID " & clng(request("cboSM")),cn,adOpenForwardOnly
	        if rs.eof and rs.bof then
	            strSM = "<BR>SM: Not Specified"
	        elseif trim(rs("Name") & "") = "" then
	            strSM = "<BR>SM: Not Specified"
	        else
	            strSM = "<BR>SM: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	        end if
            rs.close
        else
            strSM = "<BR>SM: Not Specified"
        end if


        'Lookup PDM
        dim strPDM
        if trim(request("cboPlatformDevelopment")) <> "" and isnumeric(request("cboPlatformDevelopment")) then
    	    rs.open "spGetEmployeeByID " & clng(request("cboPlatformDevelopment")),cn,adOpenForwardOnly
	        if rs.eof and rs.bof then
	            strPDM = "<BR>PDM: Not Specified"
	        elseif trim(rs("Name") & "") = "" then
	            strPDM = "<BR>PDM: Not Specified"
	        else
	            strPDM = "<BR>PDM: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	        end if
            rs.close
        else
            strPDM = "<BR>PDM: Not Specified"
        end if

        'Lookup CM 
        dim strCM
        if cint(request("cboDevCenter")) = 2 then
            if trim(request("cboTDCCM")) <> "" and isnumeric(request("cboTDCCM")) then
    	        rs.open "spGetEmployeeByID " & clng(request("cboTDCCM")),cn,adOpenForwardOnly
	            if rs.eof and rs.bof then
	                strCM = "<BR>CM: Not Specified"
	            elseif trim(rs("Name") & "") = "" then
	                strCM = "<BR>CM: Not Specified"
	            else
	                strCM = "<BR>CM: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	            end if
                rs.close
            else
                strCM = "<BR>CM: Not Specified"
            end if
        else
            if trim(request("cboPM")) <> "" and isnumeric(request("cboPM")) then
    	        rs.open "spGetEmployeeByID " & clng(request("cboPM")),cn,adOpenForwardOnly
	            if rs.eof and rs.bof then
	                strCM = "<BR>CM: Not Specified"
	            elseif trim(rs("Name") & "") = "" then
	                strCM = "<BR>CM: Not Specified"
	            else
	                strCM = "<BR>CM: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	            end if
                rs.close
            else
                strCM = "<BR>CM: Not Specified"
            end if
        end if

        
        'Lookup DevCenter
        dim strDevCenter
        rs.open "spGetDevCenterName " & cint(request("cboDevCenter")),cn,adOpenForwardOnly
        if rs.eof and rs.bof then
            strDevCenter = "<BR>Development Center: Not Specified"
        else
            strDevCenter = "<BR>Development Center: " & rs("Name")
        end if
        rs.close
        
		'Product Added
		Set oMessage = New EmailWrapper
		oMessage.From = Currentuseremail
		oMessage.To= "mobileexcalnewproducts@hp.com"
		oMessage.Subject = "New product created in Pulsar" 

		oMessage.HTMLBody = "<font face=Arial size=2 color=black>Product Name: " & request("txtProductName") & strDevCenter & strSM & strPDM & strSEPM & strCM & strpartnername & "</font>"
		
		oMessage.Send 
		Set oMessage = Nothing 	
		
	end if


		dim MailBrandArray
		dim MailBrandItem
		dim MailBrandItemArray

		'Brand/Series Added
		if strBrandsAdded <> "" then
			strBrandsAdded = mid(strBrandsAdded,2)
			MailBrandArray = split(strBrandsAdded,",")
			set rs = server.CreateObject("ADODB.Recordset")
			for each MailBrandItem in MailBrandArray
				if trim(MailBrandItem) <> "" then
					rs.open "spGetBrandSeries " &  MailBrandItem,cn,adOpenForwardOnly
					strSeriesString = strSeriesString & BuildSeriesEmail (rs, 1)
					rs.close
				end if
				
			next
			set rs = nothing
		end if
		
		
		if strSeriesString <> "" and isnumeric(trim(request("txtVersion"))) then
			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			if trim(request("txtID")) <> "100" then
				oMessage.To= "MobileExcalNotification-ProductNames@hp.com;" & Currentuseremail
			else
				oMessage.To= "max.yu@hp.com;"
			end if
			oMessage.Subject = request("txtProductName") & " series definitions updated in Pulsar" 

			oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
							    "<font face=Verdana size=2 color=black>Product series definitions updated.<BR><BR></font>" & _
							    strSeriesString  & _
							    "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
			
			oMessage.Send 
			Set oMessage = Nothing 	

		end if
		
		'Notify Service PMs when the product is first transitioned to Post-Production
		if request("txtServiceLifeDate")="" and request("tagPhase") <> request("cboPhase") and trim(request("cboPhase")) = "4" then
			rs.open "spListServicePMs",cn,adOpenForwardOnly
			ServicePMEmailList=""
			do while not rs.eof
				ServicePMEmailList = ServicePMEmailList & ";" & rs("email")
				rs.movenext	
			loop
			rs.close
			
			if ServicePMEmailList <> "" then
				ServicePMEmailList = mid(ServicePMEmailList,2)
			else
				ServicePMEmailList = "max.yu@hp.com"
			end if
			
			Set oMessage = New EmailWrapper
			oMessage.From = Currentuseremail
			if trim(request("txtID")) <> "100" then
				oMessage.To= ServicePMEmailList
			else
				oMessage.To= "max.yu@hp.com;"
			end if
			oMessage.Subject = request("txtProductName") & " has transitioned to Post-Production in Pulsar" 

			oMessage.HTMLBody = "<font size=2 face=verdana>" & request("txtProductName") & " has transitioned to Post-Production in Pulsar.<BR><BR><a target=_blank href=""http://16.81.19.70/mobilese/today/programs.asp?Commodity=1&ID=" & request("txtID") & """>Click here</a> to assign the Commodity Manager for Service and to set the anticipated End of Service Life date." & "</font>"  
			
			oMessage.Send 
			Set oMessage = Nothing 	
		end if


		'Notify with Email when System Board ID or PnP ID changed
		if trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" or trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
			'Set oMessage = New EmailWrapper
			'oMessage.From = Currentuseremail
			
			set	oMessage = New EmailQueue 		
			oMessage.From = "pulsar.support@hp.com"
						
			if trim(request("txtID")) = "100" then
    			oMessage.To= "max.yu@hp.com;"
            else
    			oMessage.To= "MobileExcalNotification-SysID@hp.com;"
            end if			
			
			
			if trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
				'Both SID and PnP ID were change
				if trim(request("txtInitialSystemBoardID")) & "" <> "" and trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " System Board ID and PnP ID changed" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" and trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " System Board ID added, PnP ID changed" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>New System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" <> "" and trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID changed, PnP ID added" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" and trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID and PnP ID added" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			elseif trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" = trim(request("txtInitialMachinePnPID")) & "" then
				'Only SID changed
				if trim(request("txtInitialSystemBoardID")) & "" <> "" then
					oMessage.To =oMessage.To  & "psgsoftpaqsupport@hp.com;TWN.PDC.NB-ReleaseLab@hp.com;"
					oMessage.Subject = request("txtProductName") & " System Board ID changed" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID has been added" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			elseif trim(request("txtSystemBoardID")) & "" = trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
				'Only PnP ID changed
				if trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " Machine PnP ID changed" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " Machine PnP ID has been added" 
					oMessage.HtmlBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			end if
					
			oMessage.SendWithOutCopy
			Set oMessage = Nothing 	

		end if
		
		
	on error resume next
	
	'Populate/Update Sudden Impact Product/ODM/GenericComponents
    if trim(request("txtID")) <> "100" then
	    cn.execute "spUpdateSuddenImpactProduct " & strID
	end if
	
	set rs = nothing	
	set cm = nothing
	set cn = nothing
	

	function FormatSystemID(strValue)
		dim RowArray
		dim Row
		dim strOutput
		dim IDArray
		
		if instr(strValue,"^")=0 and instr(strValue,"|")=0 then
			FormatSystemID = strValue
		else
			strOutput = ""
			RowArray = split(strValue,"|")
			for each Row in Rowarray
				if instr(Row,"^")=0 then
					strOutput = strOutput & ", " & Row
				else
					IDArray = split(Row,"^")
					strOutput = strOutput & ", " & IDArray(0)
					if Ubound(IDArray) > 0 then
						if trim(IDArray(1)) <> "" and trim(IDArray(1)) <> "&nbsp;" then
							strOutput = strOutput & "&nbsp;(" & replace(IDArray(1)," ","&nbsp;") & ")"
						end if
					end if
				end if
			next
			if strOutput = "" then
				FormatSystemID = "&nbsp;"
			else
				FormatSystemID = mid(strOutput,3)
			end if
			
		end if
	end function
	
Function BuildSeriesEmail(rs,TypeID)
	dim strOutput
	dim strLogoBadge
	dim strOldLogoBadge
	dim strNewLogoBadge

	dim strBrandName
	dim strOldBrandName
	dim strNewBrandName

	dim strRASFamily
	dim strOldRASFamily
	dim strNewRASFamily
	
	strOutput = ""
	do while not rs.eof
        if trim(TypeID) = "4" or  trim(TypeID) = "5" then
		    strOldlogoBadge = rs("OldLogoBadge") & "" 
		    strNewlogoBadge = rs("NewLogoBadge") & "" 
		    strOldBrandName = rs("OldBrandName") & "" 
		    strNewBrandName = rs("NewBrandName") & "" 
		    strOldRASFamily = rs("Family") & " " & rs("OldRASSegment") & " " & left(rs("Version") & "",len(rs("Version")&"")-1) & "X - " & rs("OldStreetName") & " " & rs("OldSeries")  'rs("MarketingShortName")
		    strNewRASFamily = rs("Family") & " " & rs("NewRASSegment") & " " & left(rs("Version") & "",len(rs("Version")&"")-1) & "X - " & rs("NewStreetName") & " " & rs("NewSeries")  'rs("MarketingShortName")
		    if trim(TypeID)="5" then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		    elseif trim(rs("OldSeries") & "") <> trim(rs("NewSeries") & "") then
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand and Series Updated</b></font>"
		    else
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand Updated</b></font>"
		    end if
		    if trim(TypeID) ="5" then 'Added
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("NewMarketingLongName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("NewMarketingShortName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHWeb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    else ' Updated
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & rs("OldMarketingLongName") & "<BR><b>NEW:&nbsp;</b>" & rs("NewMarketingLongName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & rs("OldMarketingShortName") & "<BR><b>NEW:&nbsp;</b>" & rs("NewMarketingShortName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strOldLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHWeb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    end if
        
        else
		    strlogoBadge = rs("LogoBadge") & "" 
		    if TypeID =3 then
			    strNewLogoBadge =  rs("NewLogoBadge") & "" 
		    end if
		    strbrandName = rs("BrandName") & "" 
		    if TypeID =3 then
			    strNewBrandName =  rs("NewBrandName") & "" 
		    end if
		    strRASFamily = rs("Family") & " " & rs("RASSegment") & " " & left(rs("Version") & "",len(rs("Version")&"")-1) & "X - " & rs("StreetName") & " " & rs("Series")  'rs("MarketingShortName")
		    if TypeID =3 then
			    strNewRASFamily = rs("Family") & " " & rs("RASSegment") & " " & left(rs("Version") & "",len(rs("Version")&"")-1) & "X - " & rs("StreetName") & " " & rs("NewSeries") 'rs("NewMarketingShortName")
		    end if
    			
		    if TypeID=1 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		    elseif TypeID=2 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Removed</b></font>"
		    elseif TypeID=3 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Updated</b></font>"
		    end if
		    if TypeID=1 or TypeID=2 then
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("MarketingLongName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("MarketingShortName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHWeb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    elseif TypeID=3 then 
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & rs("MarketingLongName") & "<BR><b>NEW:&nbsp;</b>" & rs("NewMarketingLongName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & rs("MarketingShortName") & "<BR><b>NEW:&nbsp;</b>" & rs("NewMarketingShortName") & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHWeb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewbrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    end if
        end if
		rs.movenext
	loop
	BuildSeriesEmail = strOutput
	
end function
    %>
    <input type="hidden" id="preferredLayout" value="<%=Request.Cookies("PreferredLayout2")%>" />
</body>
</html>
