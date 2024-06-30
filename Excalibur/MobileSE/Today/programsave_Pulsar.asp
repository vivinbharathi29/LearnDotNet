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
                    closePulsar2Popup(false);
                }
                else if (IsFromPulsarPlus()) {
                    ClosePulsarPlusPopup();
                }
                else if (CheckOpener() === false) {
                    parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
                } else {
                    window.returnValue = txtSuccess.value;
                    window.close();
                }
            } else if (document.getElementById("txtAddingProduct").value == '2') {//close the jquery pop up when cloning
                if (isFromPulsar2()) {
                    closePulsar2Popup(false);
                }
                else if (IsFromPulsarPlus()) {
                    ClosePulsarPlusPopup();
                }
                else if (CheckOpener() === false) {
                    parent.window.parent.ClosePropertiesDialog_fromClone(txtSuccess.value);
                } else {
                    window.returnValue = txtSuccess.value;
                    window.close();
                }
            } else { //txtAddingProduct = 0; close the jquery pop up when adding new product
                if (document.getElementById("preferredLayout").value == 'pulsar2') {
                    alert('Product Added Successfully');
                    parent.parent.window.parent.location = "../../../Excalibur/Excalibur.asp?path=pmview.asp%3FClass%3D1%26ID%3D" + txtSuccess.value;
                }
                else if (IsFromPulsarPlus()) {
                    ClosePulsarPlusPopup();
                    parent.parent.window.parent.location = "../../../Excalibur/Excalibur.asp?path=pmview.asp%3FClass%3D1%26ID%3D" + txtSuccess.value;
                }
                else if (CheckOpener() === false && sDialogView == 'add') {
                    //the ClosePropertiesDialog is initiated from leftmenu's Add New link
                    parent.parent.window.parent.ClosePropertiesDialog(txtSuccess.value, true, null);
                } else {
                    window.returnValue = txtSuccess.value;
                    window.close();
                }
            }
        }
    }

    function CheckOpener() {
        //If True, page opened with showModalDialog
        //if False, page opened with JQuery Modal Dialog
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
    dim OrgNames
    dim strSeriesString_Param


	strSeriesString = ""
    strSeriesString_Param = ""
    	
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
	    FullName = rs("Name")
		CurrentUserEmail = rs("Email") & ""
		if instr(rs("Name") & "",",")> 0 then
			CurrentUserFirstName = mid(rs("Name") & "",instr(rs("Name") & "",",")+1)
		else
			CurrentUserFirstName = ""
		end if
	end if	
	rs.close
	
    ' get the current Product's name properties Bug 14425-Task 14436<Generation not included in email notifications>
    ' record the old names
    if request("txtID") <> "" then

        'dim strProdID = clng(request("txtID"))

        set cm = server.CreateObject("ADODB.Command")
	    cm.CommandType =  &H0004
	    cm.ActiveConnection = cn
		
	    cm.CommandText = "usp_GetBrands4Product"	

        Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
        p.Value = clng(request("txtID"))
        cm.Parameters.Append p

	    Set p = cm.CreateParameter("@SelectedOnly", 16,  &H0001)
	    p.Value = 1
	    cm.Parameters.Append p

        rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
		Set rs = cm.Execute
	    set cm=nothing			

        OrgNames = ""
        do while not rs.EOF	
            ' BrandID, SeriesID, SeriesName, LongName, ShortName, FamilyName, BrandName, LogoBadge
            OrgNames = OrgNames & rs("BrandID") & "," & rs("SeriesID") & "," & rs("SeriesName") & "," & rs("LongName") & "," & rs("ShortName") & "," & rs("FamilyName") & "," & rs("BrandName") & "," & rs("LogoBadge") & "|"
            rs.movenext           
        loop
        rs.Close
        'Bug 15278 / Task 15290 - Error adding brands after Product created; check OrgName isn't empty
        If OrgNames <> "" and Right(OrgNames,1) = "|" Then
            OrgNames = left(OrgNames, len(OrgNames)-1)
            'Response.Write("<BR/>OrgNames: " & OrgNames & "<BR/>")
        End If
    end if

          

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	if request("txtID") = "" then
		cm.CommandText = "spAddProductVersion_Pulsar"
		bAddingProduct = True
    elseif request("isClone") = 1 then
        cm.CommandText = "spAddProductVersion_Pulsar"
        isCloning = True
	else
		cm.CommandText = "spUpdateProductVersion_Pulsar"
		bAddingProduct = False
        isCloning = False
	end if
	cm.CommandType =  &H0004


    'Add the following hidden field to tell if it's a new product as the pop-up is opened 
    'differently and need to be closed diffreently
    if bAddingProduct then
        Response.Write "<INPUT style='Display:none' type='text' id='txtAddingProduct' name='txtAddingProduct' value='1'>"
    elseif isCloning then
         Response.Write "<INPUT style='Display:none' type='text' id='txtAddingProduct' name='txtAddingProduct' value='2'>"
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
	p.Value = left(request("txtVersion"),20)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@ProductName", 200, &H0001, 30)
	p.Value = left(request("txtProductNameBase"),30)
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ProductLineID", 3, &H0001)
    if trim(request("cboType")) = "2" then
		p.value = 0
	else
		p.value = clng(request("cboProductLine"))
	end if
	
	cm.Parameters.Append p
	
	set p =  cm.CreateParameter("@BusinessSegmentID", 3, &H0001)
	p.value = clng(request("cboBusinessSegmentID"))
	cm.Parameters.Append p

    set p = cm.CreateParameter("@FactoryIds", 200, &H0001, 250)
	p.value = trim(request("cboFactory"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PMID", 3, &H0001)
	if trim(request("cboPM")) = "" then
		p.value = 0
	else
		p.value = clng(request("cboPM"))
	end if
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@TDCCMID", 3, &H0001)
	if trim(request("cboTDCCM")) = "" then
		p.value = 0
	else
		p.value = clng(request("cboTDCCM"))
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
	p.value = clng(request("cboComMarketing"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@SMBMarketingID", 3, &H0001)
	p.value = clng(request("cboComMarketing"))
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
	if (request("tagPhase") = 1 or bAddingProduct = true or isCloning = true) and request("cboPhase") <> 1 then 'request("txtID") = ""
		p.value = 1'Update Dates	
	elseif request("tagPhase") <> 1 and request("cboPhase") = 1 then
		p.value = 2'Empty Dates	
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
	if trim(request("cboDCRDefaultOwner")) = "1"  then
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
	'if trim(request("txtOTS")) = "" then
		p.value = left(request("txtProductName"),30)
	'else
	'	p.Value = left(request("txtOTS"),30)
	'end if
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
	'need the flag always on
    'if request("chkFusion") = "on" then
		p.Value = 1
	'else
	'	p.Value = 0
	'end if
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@FusionRequirements",11, &H0001)
	'if request("optRequirmentType") = "1" then
		p.Value = 1
	'else
	'	p.Value = 0
	'end if
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@CreateSimpleAvTypeAuto",11, &H0001)
	if request("optCreateSimpleAvType") = "1" then
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
        
    if request("txtID") = "" or isCloning then
        Set p = cm.CreateParameter("@AllowFollowMarketingName",11, &H0001)
        p.Value = 1
        cm.Parameters.Append p
	end if

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
'	if request("chkDCRAutoOpen") = "on" then
'		p.Value = 1
'	else
'		p.Value = 0
'	end if
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
        
	Set p = cm.CreateParameter("@IDInformationPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtIDInformationPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
	If len(trim(strPath)) = 0 Then
		p.Value = NULL
	Else
		p.Value = replace(strPath, "/", "\")
	End If
	cm.Parameters.Append p
                
	Set p = cm.CreateParameter("@MSPEKSExecutionPath", 200, &H0001, 256)
	strPath = left(replace(replace(lcase(request("txtMSPEKSExecutionPath")),"\\houhpqexcal01\","\\houhpqexcal01.auth.hpicorp.net\"),"\\tpopsgdev01\","\\16.159.144.31\"),256)
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

	set p =  cm.CreateParameter("@SysEngrProgramCoordinatorID", 3, &H0001)
	p.value = clng(request("cboSysEngrProgramCoordinator"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@ProgramBusinessManagerID", 3, &H0001)
	p.value = clng(request("cboProgramBusinessManager"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PreinstallCutoff", 200, &H0001, 15)
	p.value = trim(request("cboPinCutoff"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@PCID", 3, &H0001)
	p.value = clng(request("cboPC"))
	cm.Parameters.Append p

	set p =  cm.CreateParameter("@MarketingOpsID", 3, &H0001)
	p.value = clng(request("cboMarketingOps"))
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
	p.value = clng(request("cboGplm"))
	cm.Parameters.Append p

	set p = cm.CreateParameter("@SPDM", 3, &H0001)
	p.value = clng(request("cboSpdm"))
	cm.Parameters.Append p

	set p = cm.CreateParameter("@SBA", 3, &H0001)
	p.value = clng(request("cboSBA"))
	cm.Parameters.Append p
	
	set p = cm.CreateParameter("@DocPM", 3, &H0001)
	p.value = clng(request("cboDocPM"))
	cm.Parameters.Append p

    set p = cm.CreateParameter("@DKCID", 3, &H0001)
	p.value = clng(request("cboDKC"))
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
	
    if bAddingProduct = False and isCloning = False then
       Set p = cm.CreateParameter("@UserName", 200, &H0001, 30)
	   p.Value = ""
	   cm.Parameters.Append p
    end if

	if bAddingProduct = True or isCloning = True then
		set p = cm.CreateParameter("@NewID", 3, &H0002)
		cm.Parameters.Append p
	end if

    Set p = cm.CreateParameter("@FinanceID", 3, &H0001)
	p.Value = 0
	cm.Parameters.Append p

    'Set p = cm.CreateParameter("@FullName")
	'p.value = FullName
	'cm.Parameters.Append p

    Set p = cm.CreateParameter("@FullName", adVarchar, adParamInput, 50, FullName)
	cm.Parameters.Append p
        
    set p =  cm.CreateParameter("@SharedAVMarketingPMID", 3, &H0001)
	p.value = trim(request("cboSharedAvMarketing"))
	cm.Parameters.Append p
    
    set p =  cm.CreateParameter("@SharedAVPCID", 3, &H0001)
	p.value = trim(request("cboSharedAVPC"))
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

    set p =  cm.CreateParameter("@SCMOwnerId", 3, &H0001)
	p.value = clng(request("cboSCMOwner"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@EngineeringDataManagementId", 3, &H0001)
	p.value = clng(request("cboEngineeringDataManagement"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@ODMHWPMId", 3, &H0001)
	p.value = clng(request("cboODMHWPM"))
	cm.Parameters.Append p

    set p =  cm.CreateParameter("@HWPCId", 3, &H0001)
	p.value = clng(request("cboHWPC"))
	cm.Parameters.Append p

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
	if cn.Errors.count <> 0 then
			blnFailed = true
	end if	
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

	    Set p = cm.CreateParameter("@ReleaseIDs", 200,  &H0001, 128)
	    p.Value = request("txtProductReleaseIDs")
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
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
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

					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandToArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandToArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandToArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandToArray(i)))))

					strSeriesID1 = trim(request("txtSeriesIDA" & trim(clng(BrandFromArray(i)))))
					strSeriesID2 = trim(request("txtSeriesIDB" & trim(clng(BrandFromArray(i)))))
					strSeriesID3 = trim(request("txtSeriesIDC" & trim(clng(BrandFromArray(i)))))	        
					strSeriesID4 = trim(request("txtSeriesIDD" & trim(clng(BrandFromArray(i)))))	        

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
					end if					

					SeriesTagArray = split(strSeriesTag1 & chr(1) & strSeriesTag2 & chr(1) & strSeriesTag3 & chr(1) & strSeriesTag4,chr(1))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4,chr(1))
					SeriesIDArray = split(strSeriesID1 & chr(1) & strSeriesID2 & chr(1) & strSeriesID3 & chr(1) & strSeriesID4,chr(1))

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
                    strSeriesString_Param = ""
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
						        strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(BrandFromArray(i)) & "," & clng(SeriesIDArray(j)) & ",0,2|"
						   
                					else
						                'Create Notification Text for series Updated
						                strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(BrandToArray(i)) & "," & clng(SeriesIDArray(j)) & ",0,4," & clng(BrandFromArray(i)) & "|"
						   
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
		                                    strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(BrandFromArray(i)) & "," & clng(cm("@NewID")) & ",0,5|"
					   				                end if
				
                				set cm = nothing
						
						    elseif trim(SeriesIDArray(j)) <> "" and trim(SeriesArray(j)) <> "" then
				                'Create Notification Text for Brand Updated - Series Didn't change but the brand did so it is still an update
                                  strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(BrandToArray(i)) & "," & clng(SeriesIDArray(j)) & ",0,4," & clng(BrandFromArray(i)) & "|"
				      
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
		'If AllowFollowMarketingName = 1, don't Add new Product_Brand
       	for i = 0 to ubound(BrandArray)
			if trim(Brandarray(i)) <> "" and  request("hdnEnableFollowMarketingName") <> 1  then
				if (instr("," & request("txtBrandsLoaded") & ",", "," & trim(BrandArray(i)) & ",") = 0 and instr("," & replace(request("txtBrandTo")," ","") & ",", "," & trim(BrandArray(i)) & ",") = 0) or isCloning=true then
					Response.Write trim(BrandArray(i)) & "<BR>"
					
					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandArray(i)))))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4,chr(1))

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
'					if strSeries2 = "" and strSeries1 = "" then
'						p.Value = ""
'					elseif strSeries2 = "" then
'						p.Value = left(strSeries1,2000)
'					elseif strSeries1 = "" then
'						p.Value = left(strSeries2,2000)
'					else
'						p.Value = left(strSeries1 & "," & strSeries2,2000)
'					end if
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
							
							'strBrandsAdded = strBrandsAdded & ProductBrandID & ":" & left(trim(SeriesName),50) & ";"	
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
					'set rs = server.CreateObject("ADODB.Recordset")
					'rs.open "spGetBrandSeries2 " & clng(strID) & "," &  clng(BrandArray(i)),cn,adOpenForwardOnly
                    'rs.open "usp_GetBrandSeries_Pulsar " & clng(strID) & "," & clng(Brandarray(i)) & ", 0",cn,adOpenForwardOnly
					'strSeriesString = strSeriesString & BuildSeriesEmail (rs, 2, GetOrginalNames(trim(Brandarray(i)), 0))
                    strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(Brandarray(i)) & ",0,0,2|"
					'rs.close
					'set rs = nothing

					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spRemoveBrandFromProduct"
			        cm.CommandTimeout = 0
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

					strSeries1 = trim(request("txtSeriesA" & trim(clng(BrandArray(i)))))
					strSeries2 = trim(request("txtSeriesB" & trim(clng(BrandArray(i)))))
					strSeries3 = trim(request("txtSeriesC" & trim(clng(BrandArray(i)))))
					strSeries4 = trim(request("txtSeriesD" & trim(clng(BrandArray(i)))))

					strSeriesID1 = trim(request("txtSeriesIDA" & trim(clng(BrandArray(i)))))
					strSeriesID2 = trim(request("txtSeriesIDB" & trim(clng(BrandArray(i)))))
					strSeriesID3 = trim(request("txtSeriesIDC" & trim(clng(BrandArray(i)))))
					strSeriesID4 = trim(request("txtSeriesIDD" & trim(clng(BrandArray(i)))))

					'if  (strSeries2 = strSeriesTag1 and (strSeriesTag2="" or strSeriesTag2=strSeries1 )) or (strSeriesTag2 = strSeries1 and strSeries1 <> "" and strSeries2 = "" and strSeries1 <> strSeriesTag1) then
					'	strSeriesTemp = strSeries2
					'	strSeries2 = strSeries1
					'	strSeries1 = strSeriesTemp
					'end if
					
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
					end if
        
                    SeriesTagArray = split(strSeriesTag1 & chr(1) & strSeriesTag2 & chr(1) & strSeriesTag3 & chr(1) & strSeriesTag4,chr(1))
					SeriesArray = split(strSeries1 & chr(1) & strSeries2 & chr(1) & strSeries3 & chr(1) & strSeries4,chr(1))
					SeriesIDArray = split(strSeriesID1 & chr(1) & strSeriesID2 & chr(1) & strSeriesID3 & chr(1) & strSeriesID4,chr(1))

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
					'if strSeries2 = "" and strSeries1 = "" then
					'	p.Value = ""
					'elseif strSeries2 = "" then
					'	p.Value = left(strSeries1,2000)
					'elseif strSeries1 = "" then
					'	p.Value = left(strSeries2,2000)
					'else
					'	p.Value = left(strSeries1 & "," & strSeries2,2000)
					'end if
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
									'set rs = server.CreateObject("ADODB.Recordset")
									'rs.open "spGetBrandSeries3 " & clng(SeriesIDArray(j)),cn,adOpenForwardOnly
                                    'rs.open "usp_GetBrandSeries_Pulsar " & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(SeriesIDArray(j)),cn,adOpenForwardOnly
									'strSeriesString = strSeriesString & BuildSeriesEmail (rs, 2, GetOrginalNames(trim(Brandarray(i)), SeriesIDArray(j)))
                                    strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(SeriesIDArray(j)) & ",0,2|"
									'rs.close
									'set	rs = nothing
								else
									'Create Notification Text for series Updated
									'set rs = server.CreateObject("ADODB.Recordset")
									'rs.open "spGetBrandSeries3 " & clng(SeriesIDArray(j)) & ",'" & scrubsql(left(SeriesArray(j),50)) & "'" ,cn,adOpenForwardOnly
                                    'rs.open "usp_GetBrandSeries_Pulsar " & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(SeriesIDArray(j)),cn,adOpenForwardOnly
									'strSeriesString = strSeriesString & BuildSeriesEmail (rs, 3, GetOrginalNames(trim(Brandarray(i)), SeriesIDArray(j)) )
                                    strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(SeriesIDArray(j)) & ",0,3|"
									'rs.close
									'set	rs = nothing
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
								'set rs = server.CreateObject("ADODB.Recordset")
								'rs.open "spGetBrandSeries3 " & clng(cm("@NewID")),cn,adOpenForwardOnly
                                'rs.open "usp_GetBrandSeries_Pulsar " & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(cm("@NewID")),cn,adOpenForwardOnly
								'strSeriesString = strSeriesString & BuildSeriesEmail (rs, 1, "")
                                strSeriesString_Param = strSeriesString_Param & clng(strID) & "," & clng(Brandarray(i)) & "," & clng(cm("@NewID")) & ",0,1|"
								'rs.close
								'set	rs = nothing
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

                strmodelnumber = trim(request("txtModelNumber" & trim(clng(Brandarray(i)))))
                  response.Write (strmodelnumber)
			    Set p = cm.CreateParameter("@ModelNumber", 200, &H0001, 10)
			    p.Value = strmodelnumber
			    cm.Parameters.Append p
                      
                strscreenSize = trim(request("txtScreenSize" & trim(clng(Brandarray(i)))))
                  response.Write (strscreenSize)
			    Set p = cm.CreateParameter("@ScreenSize",131, &H0001)
				p.Precision = 5
				p.NumericScale = 2
				if (strscreenSize) <> "" then
					p.Value = strscreenSize
				end if
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

                if request("txtSCMEnabled" & trim(clng(Brandarray(i)))) ="1" then
                    SCMNumber = trim(request("SelectSCM" & trim(clng(Brandarray(i)))))
                else 'if disabled, get the value from the hidden txtbox
                     SCMNumber = trim(request("txtSelectedSCM" & trim(clng(Brandarray(i)))))
                end if
			    Set p = cm.CreateParameter("@SCMNumber", 3, &H0001)
			    p.Value = SCMNumber
			    cm.Parameters.Append p                
         
		
			    cm.execute
			    set cm = nothing
			    if cn.Errors.count <> 0 then
				    blnFailed = true
				    exit for
			    end if	
        next	
        

	'		    cm.execute
	'		    set cm = nothing
	'		    if cn.Errors.count <> 0 then
	'			    blnFailed = true
	'			    exit for
	'		    end if	
     '   next	


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
		if (bAddingProduct = false and isCloning = false) and trim(request("txtID")) <> "100" and request("tagProductNameBase") <> request("txtProductNameBase") then 'and request("tagVersion") <> request("txtVersion") then
'			on error resume next
'			dim cnOTS
'			set cnOTS = server.CreateObject("ADODB.Connection")
'			cnOTS.ConnectionString = Application("OTS_ConnectionString") 
 '       	cnOTS.ConnectionTimeout = 10
'			cnOTS.IsolationLevel=256
'			cnOTS.Open
'			
'			dim strSQL
'			
'			strSQL = "Update CyclePlatform " & _
'					 "set platform = '" & trim(scrubsql(request("txtProductName"))) & "' " & _
'					 "where cycle='" & trim(scrubsql(request("txtProductFamily"))) & "' " & _
'					 "and Platform='" & trim(scrubsql(request("txtProductFamily") & " " & request("tagVersion"))) & "' " & _
'				 	 "and organizationid=3 " & _
'				 	 "and Operation='Add' " & _
'					 "and partnumber like 'EXC-%'"
'					 
'			cnOTS.Execute strSQL

'			set cnOTS = nothing
'			on error goto 0
			
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
			oMessage.From = Currentuseremail
			oMessage.To= "mobileexcalnewproducts@hp.com"'"rick.rostonski@hp.com;tammy.schapiro@hp.com;kenneth.m.berntsen@hp.com;jan.clements@hp.com;diana.salas@hp.com;meghan.novak@hp.com"
			oMessage.Subject = "Product Renamed in Pulsar" 

			'oMessage.HTMLBody = "<font face=Arial size=2 color=black>Renamed By: " & CurrentUser & "<BR>Old Product Name: " & request("txtProductFamily") & " " & request("tagVersion") & "<BR>New Product Name: " & request("txtProductName") & "</font>"
            oMessage.HTMLBody = "<font face=Arial size=2 color=black>Renamed By: " & CurrentUser & "<BR>Old Product Name: " & request("tagProductNameBase") & "<BR>New Product Name: " & request("txtProductNameBase") & "</font>"
		
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

			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
			oMessage.From = Currentuseremail
			oMessage.To= "max.yu@hp.com"
			if trim(request("txtID")) <> "100" then
				oMessage.CC = "MobileExcalNotification-ProductNames@hp.com" '"tammy.schapiro@hp.com;rick.rostonski@hp.com;steve.bachmeier@hp.com;ginger.christopher@hp.com" 
'				oMessage.bcc = strdeveloperList
			end if
			if CurrentUserFirstName <> "" then
				CurrentUserFirstName = "<BR><BR>Thanks,<BR>" & CurrentUserFirstName 
			end if

			if trim(strFromPhase) = "" and trim(strToPhase) = "" then
				oMessage.Subject = request("txtProductName") & " " & request("txtVersion") & " status has been changed" 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductName") & " " & request("txtVersion") & " status has been changed in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			elseif trim(strFromPhase) = "" then
				oMessage.Subject = request("txtProductName") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductName") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) & " in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			else	
				oMessage.Subject = request("txtProductName") & " " & request("txtVersion") & " status has been changed to " & trim(strToPhase) 
				oMessage.HTMLBody = "<font face=Arial size=2 color=black>" & request("txtProductName") & " " & request("txtVersion") & " status has been changed from " & trim(strFromPhase) & " to " & trim(strToPhase) & " in Pulsar." & CurrentUserFirstName & "</font>" ' Unless a significant field or factory issue arises, we will no longer release new deliverables for this product. If an issue does arise, the product will be reset to active until the issue is resolved.<BR><BR>If you have any questions or concerns about this status change, do not hesitate to contact me." & CurrentUserFirstName & "</font>"
			end if
			oMessage.Importance = cdoHigh
            
            oMessage.Send 
			Set oMessage = Nothing 			
		
		end if
		


		if (bAddingProduct = false and isCloning = false) and trim(request("tagPhase")) <> trim(request("cboPhase")) and trim(request("tagPhase"))= "5" and trim(request("cboPhase"))<> "5" then
			
			
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
			oMessage.From = Currentuseremail
			oMessage.To= "max.yu@hp.com;"
			if trim(request("txtID")) <> "100" then
				oMessage.CC = "tammy.schapiro@hp.com" 
			end if
			oMessage.Subject = request("txtProductFamily") & " " & request("txtVersion") & " Reactivated in Pulsar" 

			oMessage.HTMLBody = "<font face=Arial size=2 color=black>Reactivated By: " & CurrentUser & "</font>" '& "<BR>Please reactivate all CMT deliverables attached to " & request("txtProductFamily") & " " & request("txtVersion")
		
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
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
		oMessage.From = Currentuseremail
		oMessage.To= "max.yu@hp.com"
		oMessage.Subject = "Product ODM Updated" 

		oMessage.HTMLBody = "<font face=Arial size=2 color=black>Product Name: " & request("txtProductName") & strpartnername & "</font>"
		
		oMessage.Send 
		Set oMessage = Nothing 	
    
    end if

	if bAddingProduct = true or isCloning = true then ' or lcase(Session("LoggedInUser")) = "auth\dwhorton" then
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
        if trim(request("hdnIsDesktop")) = "YES" then
    	    rs.open "spGetEmployeeByID " & clng(request("cboSCMOwner")),cn,adOpenForwardOnly
	        if rs.eof and rs.bof then
	            strCM = "<BR>SCM Owner: Not Specified"
	        elseif trim(rs("Name") & "") = "" then
	            strCM = "<BR>SCM Owner: Not Specified"
	        else
	            strCM = "<BR>SCM Owner: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	        end if
            rs.close
        else
    	    rs.open "spGetEmployeeByID " & clng(request("cboPM")),cn,adOpenForwardOnly
	        if rs.eof and rs.bof then
	            strCM = "<BR>Configuration Manager: Not Specified"
	        elseif trim(rs("Name") & "") = "" then
	            strCM = "<BR>Configuration Manager: Not Specified"
	        else
	            strCM = "<BR>Configuration Manager: " & trim(rs("Name") & "") & " [ " & trim(rs("Email") & "") & " ]"
	        end if
            rs.close
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
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")
		oMessage.From = Currentuseremail
		oMessage.To= "mobileexcalnewproducts@hp.com"'"rick.rostonski@hp.com;tammy.schapiro@hp.com"
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
					'MailBrandItemArray = split(MailBrandItem,":")
					'rs.open "spGetBrandSeries " &  MailBrandItemArray(0),cn,adOpenForwardOnly
					'rs.open "spGetBrandSeries " &  MailBrandItem,cn,adOpenForwardOnly
                    'rs.open "usp_GetBrandSeries_Pulsar 0, 0, 0," & MailBrandItem,cn,adOpenForwardOnly
					'strSeriesString = strSeriesString & BuildSeriesEmail (rs, 1, "")
                    strSeriesString_Param = strSeriesString_Param & "0,0,0," & MailBrandItem & ",1|"
'					do while not rs.eof
'					'if not (rs.eof and rs.bof) then
'					
'						if ucase(left(trim(rs("MarketingShortName") & ""),3))  = "HP " then
'							strLogoBadge = mid(rs("MarketingShortName") & "",4)''
'						else
'							strlogoBadge = rs("MarketingShortName") & ""
'						end if
'					
'					    strSeriesString = strSeriesString & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
'						strSeriesString = strSeriesString & "<TABLE border=1 dellpadding=2 cellspacing=0>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("MarketingLongName") & "</font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("MarketingShortName") & "</font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strLogoBadge & "</font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("Family") & " " & rs("RASSegment") & " " & left(rs("Version") & "",len(rs("Version")&"")-1) & "X - " & rs("MarketingShortName") & "</font></TD></TR>"
'						strSeriesString = strSeriesString & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & rs("MarketingShortName") & "</font></TD></TR>"
'						strSeriesString = strSeriesString & "</TABLE><BR><BR>"
'					'end if
'						rs.movenext
'					loop
					'rs.close
				end if
	'			strSeriesString = strSeriesString & lbound(MailBrandItemArray) & ":" & ubound(MailBrandItemArray) & "<BR>"
				
			next
			set rs = nothing
		end if
		
		'Thien -----
	    ' get the updated brand/series names, preparing for the notification email body. Bug 14425-Task 14436<Generation not included in email notifications>
        if request("txtID") <> "" and strSeriesString_Param <> "" then
           dim arrParams
           dim strParamFields

           if Right(strSeriesString_Param,1) = "|" then
                strSeriesString_Param = Left(strSeriesString_Param,Len(strSeriesString_Param)-1)
           end if

           arrParams = split(strSeriesString_Param, "|")
           for each strParamFields in arrParams
               'Response.Write("<br/>strParamFields: " & strParamFields & "<br/>")
               strSeriesString = strSeriesString & BuildSeriesEmail_Pulsar (strParamFields)
           next

        end if
        '-------

		if strSeriesString <> "" then 'and isnumeric(trim(request("txtVersion"))) then
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")
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

			oMessage.HTMLBody = "<font size=2 face=verdana>" & request("txtProductName") & " has transitioned to Post-Production in Pulsar.<BR><BR><a target=_blank href=""http://" & Application("Excalibur_ServerName") & "/excalibur/mobilese/today/programs.asp?Commodity=1&ID=" & request("txtID") & """>Click here</a> to assign the Commodity Manager for Service and to set the anticipated End of Service Life date." & "</font>"  
			
			oMessage.Send 
			Set oMessage = Nothing 	
		end if


		'Notify with Email when System Board ID or PnP ID changed
		if trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" or trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
			set	oMessage = New EmailQueue 		
			oMessage.From = "pulsar.support@hp.com"
			oMessage.To = "MobileExcalNotification-SysID@hp.com;" & Currentuseremail

			
			if trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
				if trim(request("txtInitialSystemBoardID")) & "" <> "" and trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " System Board ID and PnP ID changed" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" and trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " System Board ID added, PnP ID changed" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>New System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" <> "" and trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID changed, PnP ID added" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" and trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID and PnP ID addedd" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			elseif trim(request("txtSystemBoardID")) & "" <> trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" = trim(request("txtInitialMachinePnPID")) & "" then
				if trim(request("txtInitialSystemBoardID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " System Board ID changed" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID changed from: " & request("txtInitialSystemBoardID") & " TO: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialSystemBoardID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " System Board ID has been added" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>System Board ID has been added: " & formatsystemid(request("txtSystemBoardComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			elseif trim(request("txtSystemBoardID")) & "" = trim(request("txtInitialSystemBoardID")) & "" and trim(request("txtMachinePnPID")) & "" <> trim(request("txtInitialMachinePnPID")) & "" then
				if trim(request("txtInitialMachinePnPID")) & "" <> "" then
					oMessage.Subject = request("txtProductName") & " Machine PnP ID changed" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID changed from: " & request("txtInitialMachinePnPID") & " TO: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				elseif trim(request("txtInitialMachinePnPID")) & "" = "" then
					oMessage.Subject = request("txtProductName") & " Machine PnP ID has been added" 
					oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & request("txtProductName") & "</b><BR><BR></font>" & _
						   "<font face=Verdana size=2 color=black>Machine PnP ID has been added: " & formatsystemid(request("txtMachinePnPComments")) & "<BR><BR></font>" & _
						   "<BR><font face=Arial size=2 color=black>Updates By: " & CurrentUser & " - " & now() 
				end if
			end if
					
			oMessage.Send
			Set oMessage = Nothing 	

		end if
		
		
	on error resume next
	
	'Populate/Update Sudden Impact Product/ODM/GenericComponents
    if trim(request("txtID")) <> "100" then
	    cn.execute "spUpdateSuddenImpactProduct " & strID
	end if

    'GET Approvers Team Roster selection
    dim teamRosterSel
    dim chkTeamRoster
    dim ArrTeamRoster
    dim strSelectedTeamRoster
    dim k

    teamRosterSel = request("chkDCRAutoOpen") 
             
    if (teamRosterSel = "2") then 
            chkTeamRoster = request("ckTeamRosterAndODM")
                ArrTeamRoster  = split(chkTeamRoster,",")
                strSelectedTeamRoster = ""

                for k=LBound(ArrTeamRoster) to UBound(ArrTeamRoster)
                    strSelectedTeamRoster = strSelectedTeamRoster + ArrTeamRoster(k) + ","
                Next
    
                if(LEN(strSelectedTeamRoster) > 0) then
                        strSelectedTeamRoster = LEFT(strSelectedTeamRoster, LEN(strSelectedTeamRoster) -1)
                end if
        end if 

    if (teamRosterSel = "4") then
            chkTeamRoster = request("ckTeamRosterNoODM")
                ArrTeamRoster  = split(chkTeamRoster,",")
                strSelectedTeamRoster = ""

                for k=LBound(ArrTeamRoster) to UBound(ArrTeamRoster)
                    strSelectedTeamRoster = strSelectedTeamRoster + ArrTeamRoster(k) + ","
                Next
    
                if(LEN(strSelectedTeamRoster) > 0) then
                        strSelectedTeamRoster = LEFT(strSelectedTeamRoster, LEN(strSelectedTeamRoster) -1)
                end if
    end if 


		if strSelectedTeamRoster <> "" then

			set cm = nothing
			set cm = server.CreateObject("ADODB.command")

			cm.ActiveConnection = cn
			cm.CommandText = "usp_ProductSystemTeamRoster_Update"
			cm.CommandType =  &H0004


			set p =  cm.CreateParameter("@p_ProductVersionId", 3, &H0001)
			p.value = clng(strID)
			cm.Parameters.Append p
			
			set p = cm.CreateParameter("@p_chrPrimaryTeamRosterIds", 200, &H0001, 64)
	            p.Value = strSelectedTeamRoster
	            cm.Parameters.Append p

        		cm.Execute RowsEffected
		    
            set cm=nothing

		    if cn.Errors.count <> 0 then
			    blnFailed = true
		    end if		
		end if
    'End save Team Roster Approvers
        	
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
	
Function BuildSeriesEmail(rs,TypeID, oldNames)
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
   ' dim strpdversion
    dim strOldLongName
    dim strOldShortName
    dim strNewLongName
    dim strNewShortName

    dim arrOldNames

    'Response.Write("TypeID: " & TypeID & "</br>")
    'Response.Write("oldNames: [" & oldNames & "]</br>")

    if oldNames <> "" then

        arrOldNames = split(oldNames, ",")

        strOldLongName = arrOldNames(3)
        strOldShortName = arrOldNames(4)
        strOldRASFamily = arrOldNames(5)
        strOldBrandName = arrOldNames(6)
	    strOldLogoBadge = arrOldNames(7)
    else
        strOldLongName = ""
        strOldShortName = ""
        strOldRASFamily = ""
        strOldBrandName = ""
	    strOldLogoBadge = ""
    end if
    if not rs.eof then  
        strNewLongName =  rs("LongName") & "" 
        strNewShortName = rs("ShortName") & "" 
        strNewLogoBadge = rs("LogoBadge") & "" 
	    strNewBrandName = rs("BrandName") & "" 
        strNewRASFamily = rs("FamilyName") & "" 
    else
        strNewLongName =  "" 
        strNewShortName = "" 
        strNewLogoBadge = ""
	    strNewBrandName = ""
        strNewRASFamily = ""
    end if

	strOutput = ""
	'do while not rs.eof
        if trim(TypeID) = "4" or  trim(TypeID) = "5" then
		    'strOldlogoBadge = rs("OldLogoBadge") & "" 
		    'strNewlogoBadge = rs("NewLogoBadge") & "" 
		    'strOldBrandName = rs("OldBrandName") & "" 
		    'strNewBrandName = rs("NewBrandName") & "" 
            'if len(rs("Version")) > 0 then
            '    strpdversion = left(rs("Version") & "",len(rs("Version")&"")-1)
            'else
            '    strpdversion = ""
            'end if
            
		    'strOldRASFamily = rs("Family") & " " & rs("OldRASSegment") & " " & strpdversion & "X - " & rs("OldStreetName") & " " & rs("OldSeries")  'rs("MarketingShortName")
		    'strNewRASFamily = rs("Family") & " " & rs("NewRASSegment") & " " & strpdversion & "X - " & rs("NewStreetName") & " " & rs("NewSeries")  'rs("MarketingShortName")
		    if trim(TypeID)="5" then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		    elseif trim(rs("OldSeries") & "") <> trim(rs("NewSeries") & "") then
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand and Series Updated</b></font>"
		    else
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand Updated</b></font>"
		    end if
		    if trim(TypeID) ="5" then 'Added
			    strOutput = strOutput & "<TABLE border=1 cellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLongName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewShortName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    else ' Updated
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldLongName & "<BR><b>NEW:&nbsp;</b>" & strNewLongName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldShortName & "<BR><b>NEW:&nbsp;</b>" & strNewShortName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strOldLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    end if
        
        else
		    'strlogoBadge = rs("LogoBadge") & "" 
		    'if TypeID =3 then
			'    strNewLogoBadge =  rs("NewLogoBadge") & "" 
		    'end if
		    'strbrandName = rs("BrandName") & "" 
		    'if TypeID =3 then
			'    strNewBrandName =  rs("NewBrandName") & "" 
		    'end if
            ' if len(rs("Version")) > 0 then
            '    strpdversion = left(rs("Version") & "",len(rs("Version")&"")-1)
           ' else
           '     strpdversion = ""
           ' end if
		   ' strRASFamily = rs("Family") & " " & rs("RASSegment") & " " & strpdversion & "X - " & rs("StreetName") & " " & rs("Series")  'rs("MarketingShortName")
		   ' if TypeID =3 then
		'	    strNewRASFamily = rs("Family") & " " & rs("RASSegment") & " " & strpdversion & "X - " & rs("StreetName") & " " & rs("NewSeries") 'rs("NewMarketingShortName")
		'    end if
    			
		    if TypeID=1 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		    elseif TypeID=2 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Removed</b></font>"
		    elseif TypeID=3 then	
			    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Updated</b></font>"
		    end if
		    if TypeID=1 then
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLongName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewShortName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    elseif TypeID=2 then
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLongName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldShortName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    elseif TypeID=3 then 
			    strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldLongName & "<BR><b>NEW:&nbsp;</b>" & strNewLongName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldShortName & "<BR><b>NEW:&nbsp;</b>" & strNewShortName & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strOldLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewBrandName & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
		    end if
        end if
		'rs.movenext
	'loop
	BuildSeriesEmail = strOutput
	
end function

function GetOrginalNames(BrandID, SeriesID, TypeID)
    dim theCurNames
    dim arrNameFields
    dim arrNames

    arrNames = split(OrgNames, "|")
    theCurNames = ""

    if TypeID =2 and SeriesID=0 then
        for each arrName in arrNames
            arrNameFields = split(arrName, ",")
            if arrNameFields(0) = BrandID then
                theCurNames = theCurNames & arrName & ";"
            end if
        next
    else
        for each arrName in arrNames
            arrNameFields = split(arrName, ",")
            if arrNameFields(0) = BrandID and arrNameFields(1) = SeriesID then
                theCurNames = arrName
                exit for
            end if
        next
    end if
    GetOrginalNames = theCurNames
end function

function BuildSeriesEmail_Pulsar(Params)
    dim strOutput
	dim strOldLogoBadge
	dim strOldBrandName
	dim strOldRASFamily
    dim strOldLongName
    dim strOldShortName
    dim strOldSeriesName

	dim strNewLogoBadge
	dim strNewBrandName
	dim strNewRASFamily
    dim strNewLongName
    dim strNewShortName
    dim strNewSeriesName

    dim oldNames: oldNames = ""
    dim arrOldNames
    dim arrParam
    dim ProductVersionID: ProductVersionID = 0
    dim BrandID: BrandID = 0
    dim SeriesID: SeriesID = 0
    dim ProductBrandID: ProductBrandID = 0
    dim TypeID: TypeID = 0
    dim OldBrandID: OldBrandID = 0
'Response.Write("<br/>BuildSeriesEmail_Pulsar:" & Params & "<br/>")
    arrParam = Split(Params, ",")

    ProductVersionID = arrParam(0)
    BrandID = arrParam(1)
    SeriesID = arrParam(2)
    ProductBrandID = arrParam(3)
    TypeID = arrParam(4)

    if trim(TypeID) = "2" then  'and SeriesID = 0 then ' removing a whole brand; separate it here so we can display all removed series
        strOutput = BuildSeriesEmail_Pulsar_BrandRemoved(Params)
    else
        ' get old names
        If trim(TypeID) = "4" Then
            OldBrandID = arrParam(5)
            oldNames = GetOrginalNames(OldBrandID, SeriesID, TypeID)
        Else
            oldNames = GetOrginalNames(BrandID, SeriesID, TypeID)
        End IF

        
        'Response.Write("oldNames: [" & oldNames & "]</br>")
        if oldNames <> "" then
            arrOldNames = split(oldNames, ",")
            strOldLongName = arrOldNames(3)
            strOldShortName = arrOldNames(4)
            strOldRASFamily = arrOldNames(5)
            strOldBrandName = arrOldNames(6)
	        strOldLogoBadge = arrOldNames(7)
        else
            strOldLongName = ""
            strOldShortName = ""
            strOldRASFamily = ""
            strOldBrandName = ""
	        strOldLogoBadge = ""
        end if

        ' get updated names
        set rs = server.CreateObject("ADODB.Recordset")
        rs.open "usp_GetBrandSeries_Pulsar " & clng(ProductVersionID) & "," & clng(BrandID) & "," & clng(SeriesID) & "," & clng(ProductBrandID),cn,adOpenForwardOnly

        strOutput = ""
        do while not rs.eof

            strNewLongName =  rs("LongName") & "" 
            strNewShortName = rs("ShortName") & "" 
            strNewLogoBadge = rs("LogoBadge") & "" 
	        strNewBrandName = rs("BrandName") & "" 
            strNewRASFamily = rs("FamilyName") & "" 
            strNewSeriesName = rs("SeriesName") & "" 

            if trim(TypeID) = "4" or  trim(TypeID) = "5" then
		        if trim(TypeID)="5" then	
			        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		        elseif trim(strOldSeriesName) <> trim(strNewSeriesName) then
			        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand and Series Updated</b></font>"
		        else
			        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Brand Updated</b></font>"
		        end if
		        if trim(TypeID) ="5" then 'Added
			        strOutput = strOutput & "<TABLE border=1 cellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLongName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewShortName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLogoBadge & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewRASFamily & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewBrandName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
		        else ' Updated
			        strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldLongName & "<BR><b>NEW:&nbsp;</b>" & strNewLongName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldShortName & "<BR><b>NEW:&nbsp;</b>" & strNewShortName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strOldLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewBrandName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
		        end if
        
            else
		        if TypeID=1 then	
			        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Added</b></font>"
		        elseif TypeID=2 then
                        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Removed</b></font>"
		        elseif TypeID=3 then	
			        strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Updated</b></font>"
		        end if
		        if TypeID=1 then
			        strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLongName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewShortName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewLogoBadge & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewRASFamily & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strNewBrandName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
		        elseif TypeID=2 then
			        strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLongName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldShortName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLogoBadge & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldRASFamily & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldBrandName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
		        elseif TypeID=3 then 
			        strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldLongName & "<BR><b>NEW:&nbsp;</b>" & strNewLongName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldShortName & "<BR><b>NEW:&nbsp;</b>" & strNewShortName & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:</b>" & strOldLogoBadge & "<BR><b>NEW:&nbsp;</b>" & strNewLogoBadge & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldRASFamily & "<BR><b>NEW:&nbsp;</b>" & strNewRASFamily & "</font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana><b>OLD:&nbsp;</b>" & strOldBrandName & "<BR><b>NEW:&nbsp;</b>" & strNewBrandName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
		        end if
            end if
            rs.movenext
        loop

        rs.close
        set	rs = nothing
    end if

    BuildSeriesEmail_Pulsar = strOutput
end function

function BuildSeriesEmail_Pulsar_BrandRemoved(Params) 'TypeID=2 'and SeriesID=0 - removing a whole brand

    dim strOutput
    dim strOldLogoBadge
	dim strOldBrandName
	dim strOldRASFamily
    dim strOldLongName
    dim strOldShortName
    dim strOldSeriesName

    dim oldNames: oldNames = ""
    dim arrOldNames
    dim arrParam
    dim entry
    dim arrFields

    dim ProductVersionID: ProductVersionID = 0
    dim BrandID: BrandID = 0
    dim SeriesID: SeriesID = 0
    dim ProductBrandID: ProductBrandID = 0
    dim TypeID: TypeID = 0
    dim OldBrandID: OldBrandID = 0

    strOldLongName = ""
    strOldShortName = ""
    strOldRASFamily = ""
    strOldBrandName = ""
	strOldLogoBadge = ""

    arrParam = Split(Params, ",")

    ProductVersionID = arrParam(0)
    BrandID = arrParam(1)
    SeriesID = arrParam(2)
    ProductBrandID = arrParam(3)
    TypeID = arrParam(4)

    ' get old names
    oldNames = GetOrginalNames(BrandID, SeriesID, TypeID)
 
    ' possible multiple series belonging to the removed brand, need to split it (; as delimiter)
    If oldNames <> "" and Right(oldNames,1) = ";" Then
         oldNames = left(oldNames, len(oldNames)-1)
    End If

    arrOldNames = Split(oldNames, ";") ' array of series names

    strOutput = ""
    for each entry in arrOldNames

        arrFields = split(entry, ",")

        strOldLongName = arrFields(3)
        strOldShortName = arrFields(4)
        strOldRASFamily = arrFields(5)
        strOldBrandName = arrFields(6)
	    strOldLogoBadge = arrFields(7)

		' TypeID=2 and SeriesID = 0 then
	    strOutput = strOutput & "<font face=Arial size=3 color=black><b>Series Removed</b></font>"
		strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
		strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Long Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLongName & "</font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Short Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldShortName & "</font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo&nbsp;Badge&nbsp;C&nbsp;Cover:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldLogoBadge & "</font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldRASFamily & "</font></TD></TR>"
		strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Brand Name:</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & strOldBrandName & "</font></TD></TR>"
		strOutput = strOutput & "</TABLE><BR><BR>"
    next
    BuildSeriesEmail_Pulsar_BrandRemoved = strOutput
end function
    %>
    <input type="hidden" id="preferredLayout" value="<%=Request.Cookies("PreferredLayout2")%>" />
</body>
</html>