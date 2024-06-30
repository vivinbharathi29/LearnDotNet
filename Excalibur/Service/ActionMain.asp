<%@  language="VBScript" %>
<%
  Option Explicit
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

' Issue Types
' 1 = Issue
' 2 = Action Item
' 3 = Change Request
' 4 = Status Note
' 5 = Improvement Opportunity  
' 6 = Test Request
	
	Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim TypeID : TypeID = regEx.Replace(Request.QueryString("Type"), "")
    Dim ProdID : ProdID = regEx.Replace(Request.QueryString("ProdID"), "")
    Dim IssueID : IssueID = regEx.Replace(Request.QueryString("ID"), "")
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")  
%>
<html>
<head>
    <title>ActionMain</title>

    <script type="text/javascript" language="javascript" src="../_ScriptLibrary/jsrsClient.js"></script>

    <meta http-equiv="Pragma" content="no-cache" />
    <meta http-equiv="Expires" content="-1" />
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE" />
    <meta name="VI60_defaultClientScript" content="JavaScript" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <link rel="stylesheet" type="text/css" href="../style/programoffice.css" />

    <script id="clientEventHandlersJS" type="text/javascript" language="javascript">
<!--


        function cmdDate_onclick(target) {
            var strID;
            var txtDateField = document.getElementById(target);
            strID = window.showModalDialog("/mobilese/today/calDraw1.asp", txtDateField.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                txtDateField.value = strID;
            }
        }

        function cmdOwnerAdd_onclick() {
            ChooseEmployee(ProgramInput.cboOwner);
        }

        function ChooseEmployee(myControl) {
            var ResultArray;

            ResultArray = window.showModalDialog("/MobileSE/Today/ChooseEmployee.asp", "", "dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")

            if (typeof (ResultArray) != "undefined") {
                if (ResultArray[0] != 0) {
                    myControl.options[myControl.length] = new Option(ResultArray[1], ResultArray[0]);
                    myControl.selectedIndex = myControl.length - 1;
                }
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
            var cboEmployee = document.getElementById("cboEmployee");
            
            DeleteCell.style.display = "";
            for (i = 0; i < ProgramInput.txtApproversLoaded.value; i++)
                document.all("Del" + i).style.display = "";

            NewRow = ApproverTable.insertRow(ApproverTable.rows.length - 1);
            NewRow.name = "Row" + (ApproverTable.rows.length - 2);
            NewRow.id = "Row" + (ApproverTable.rows.length - 2);
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "<INPUT type=\"checkbox\" id=\"chkDelete" + (ApproverTable.rows.length - 2) + "\" name=\"chkDelete" + (ApproverTable.rows.length - 2) + "\">";
            NewCell = NewRow.insertCell();
            NewCell.innerHTML = "<SELECT size=1 id=cboApprover" + (ApproverTable.rows.length - 2) + " name=cboApprover" + (ApproverTable.rows.length - 2) + "  LANGUAGE=javascript onkeypress=\"return combo_onkeypress()\" onfocus=\"return combo_onfocus()\" onclick=\"return combo_onclick()\" onkeydown=\"return combo_onkeydown()\">" + cboEmployee.innerHTML + "</SELECT>&nbsp;<INPUT type=\"button\" value=\"Add\" id=button1 name=button1 LANGUAGE=javascript onclick=\"return ChooseEmployee(ProgramInput." + "cboApprover" + (ApproverTable.rows.length - 2) + ");\">";
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
            var rc;
            if (window.confirm("Are you sure?")) {
                rc = window.showModalDialog("../MobileSe/Today/DeleteAction.asp?xxI1Iu4uT9Tg6gR2R=" + strID, "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                if (typeof (rc) != "undefined") {
                    if (rc == "1") {
                        window.returnValue = 1;
                        window.close();
                    }
                }
            }
        }

        function ChangeStatus(strID) {

            var rc;

            if (document.all("Status" + strID).innerHTML == "Cancelled") {
                if (confirm("Are you sure you want to reset this status to Requested?")) {
                    rc = window.showModalDialog("../MobileSe/Today/ApproverStatus.asp?ActionID=" + ProgramInput.txtID.value + "&ID=" + strID + "&Status=1", "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
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
                    rc = window.showModalDialog("../MobileSe/Today/ApproverStatus.asp?ActionID=" + ProgramInput.txtID.value + "&ID=" + strID + "&Status=4", "", "dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
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

        //        function Left(str, n) {
        //            if (n <= 0)     // Invalid bound, return blank string
        //                return "";
        //            else if (n > String(str).length)   // Invalid bound, return
        //                return str;                // entire string
        //            else // Valid bound, return appropriate substring
        //                return String(str).substring(0, n);
        //        }

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
            if (txtRecordLocked.value == "1") {
                if (typeof (window.parent.frames["LowerWindow"].cmdSubmit) != "undefined")
                    window.parent.frames["LowerWindow"].cmdSubmit.disabled = true;
            }
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
                alert("Automatically selected the product defined for this cycle.\r\rPlease verify the product list was updated correctly.");
            else if (UpdateCount == ProdCount)
                alert("Automatically selected all " + UpdateCount + " products defined for this cycle.\r\rPlease verify the product list was updated correctly.");
            else if (ProdCount == 1)
                alert("Unable to find the only product defined for this cycle.  No products have been selected.");
            else
                alert("Automatically selected only " + UpdateCount + " of the " + ProdCount + " products defined for this cycle.\r\rPlease verify the product list was updated correctly.");
        }

        function cmdAdd_onclick() {
            var strResult;
            strResult = window.showModalDialog("../Email/AddressBook.asp?AddressList=" + ProgramInput.txtNotify.value, "", "dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No");
            if (typeof (strResult) != "undefined")
                ProgramInput.txtNotify.value = strResult;

        }

        function AddAttachment() {
            var strID;
            strID = window.showModalDialog("../PMR/SoftpaqFrame.asp?Title=Upload Test Report&Page=../common/fileupload.aspx", "", "dialogWidth:600px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
            if (typeof (strID) != "undefined") {
                AddAttachmentLinks.style.display = "none";
                UpdateAttachmentLinks.style.display = "";
                document.getElementById("txtAttachment").value = strID;
            }
        }

        function RemoveAttachment() {
            AddAttachmentLinks.style.display = "";
            UpdateAttachmentLinks.style.display = "none";
            document.getElementById("txtAttachment").value = "";
        }

        function ViewTestReport() {
            window.open("../DownloadZip.aspx?file=" + encodeURIComponent( document.getElementById("txtAttachment").value) );
        }
//-->
    </script>

    <style type="text/css">
        A:visited
        {
            color: blue;
        }
        A:hover
        {
            color: red;
        }
        TABLE
        {
            font-size: x-small;
            font-family: Verdana;
        }
    </style>
    <link rel="stylesheet" type="text/css" href="../Style/programoffice.css" />
</head>
<%
	if trim(TypeID) = "" then
		Response.Write "<body bgcolor=""Ivory""><INPUT type=""hidden"" id=txtRecordLocked name=txtRecordLocked value=""""><FONT face=verdana size=2>Not enough information supplied to display this page.</font>"
		Response.End()
	end if

	
	dim rs 
	dim cn
	dim strProducts
	dim strOwners
	dim DisplayForAdd
	dim DisplayForChangeOnly
	dim JustificationTemplate
	dim strID
	dim strPMID
	dim strPCID
	dim strSummary
	dim strRep
	dim strReps
	dim strSubmitter
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
	dim strAmericas
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
	dim NoApprovals
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
	dim  blnPreinstallApprover
	dim PreinstallOwnerID
	dim strPreinstallOwnerList
	dim blnProdFound
	dim strAvailableForTest
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
	dim IsPM
	dim strNetAffect
	dim strOSList
	dim strCommodityManagerID
	dim strDetails
	dim strStatusValue
	dim strInitiator
	dim strInitiatorText
	dim strSpareKitPn
	dim strSubAssemblyPn
	dim strInventoryDispositionId
    dim strInventoryDisposition	
    dim strQSpecDt
    dim strCompEcoDt
    dim strCompEcoNo
    dim strSaEcoDt
    Dim strSaEcoNo
    dim strSpsEcoDt
    dim strSpsEcoNo
    dim strAttachment
    dim strBomAnalystComments
	
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
		CurrentUserID = rs("ID")
		CurrentUserSysAdmin = rs("SystemAdmin")
		CurrentUserGroup = rs("WorkgroupID")
	end if
	rs.Close

	
	rs.Open "spListPMsActive 4",cn,adOpenForwardOnly
	isPM=false
	do while not rs.EOF
		if trim(rs("ID")) = trim(CurrentUserID) then
			isPM = true
		end if
		rs.MoveNext
	loop
	rs.Close

%>
<body bgcolor="Ivory" onload="return window_onload()">
<form action="ActionSave.asp" method="post" name="ProgramInput">
<% 

	
	blnProdFound = false	
	strPMID = ""
	strPCID = ""
	strSMID = ""
	blnPreinstallApprover = false
	PreinstallOwnerID = ""
	strPreinstallOwnerList = ""
	strID = ""
	strSummary=""
	strDescription = ""
	strRep = ""
	strSubmitter = ""
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
	strAmericas = ""
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
	strAvailableNotes = ""
	strProdApprovalList = ""
	strApproverListLink = ""
	strProductname= ""
	strReqChange = ""
	SustainingProduct = false
	strNetAffect = ""
	strDetails = ""
	strStatusValue = ""
	strInitiator = ""
	strInitiatorText = ""
	strSpareKitPn = ""
	strSubAssemblyPn = ""
    strInventoryDispositionId = ""
    strInventoryDisposition	= ""
    strAttachment = ""
    strQSpecDt = ""
    strCompEcoDt = ""
    strCompEcoNo = ""
    strSaEcoDt = ""
    strSaEcoNo = ""
    strSpsEcoDt = ""
    strSpsEcoNo = ""
    strAttachment = ""
	strBomAnalystComments = ""
	
	if CategoryID = "1" then
		strSKUChange = "checked"
	else
		strSKUChange = ""
	end if
	strImageChange = ""
	strCommodityChange = ""
	
	if IssueID <> "" then

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetActionProperties"

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = IssueID
		cm.Parameters.Append p

		rs.CursorType = adOpenStatic
		Set rs = cm.Execute 
		Set cm=nothing

		if rs.eof and rs.bof then
			strID=0
		else
			strID = IssueID
			strAction  = rs("Actions") & ""
			strResolution  = rs("Resolution") & ""
			strSummary  = replace(rs("Summary") & "","""","&QUOT;")
			strAttachment = rs("Description") & ""
			strRep = rs("CoreTeamRep") & ""
			strSubmitter = rs("Submitter") & ""
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
			strAmericas = rs("Americas") & ""
			strAPJ = rs("APJ") & ""
			strEMEA = rs("EMEA") & ""
			strReqChange = rs("ReqChange") & ""
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
			strAvailableNotes = rs("AvailableNotes") & ""
			strDetails = rs("Details") & ""
			strSubAssemblyPn = rs("SubAssemblyPn") & ""
			strSpareKitPn = rs("SpareKitPn") & ""
			strInitiator = rs("InitiatedBy") & ""
			strInventoryDispositionId = rs("InventoryDisposition") & ""
			strQSpecDt = rs("QSpecSubmittedDt")
            strCompEcoDt = rs("CompEcoSubmittedDt")
            strCompEcoNo = rs("CompEcoNo")
            strSaEcoDt = rs("SaEcoSubmittedDt")
            strSaEcoNo = rs("SaEcoNo")
            strSpsEcoDt = rs("SpsEcoSubmittedDt") & ""
            strSpsEcoNo = rs("SpsEcoNo")
            strBomAnalystComments = "" & ""
			'if Trim(strDeliverableRootID) = "" Then strDeliverableRootID = "0"

		end if
		rs.Close
		
		rs.Open "SELECT AddDCRNotificationList FROM ProductVersion with (NOLOCK) WHERE ID=" & strProgramId, cn, adOpenStatic
        Dim AddDCRNotificationList
        AddDCRNotificationList = rs("AddDCRNotificationList")
        rs.Close

		Select Case strInventoryDispositionId
            Case "1"
                strInventoryDisposition = "No OSSP Inventory Affected"
            Case "2"
                strInventoryDisposition = "Use current material until depleted, then roll to new material"
            Case "3"
                strInventoryDisposition = "Purge (Quality Issue - OTHER ACTION IS REQUIRED!)"
            Case Else
                strInventoryDisposition = "&nbsp;"
		End Select
		
		Select Case strInitiator
		    Case "1"
		        strInitiatorText = "ODM"
		    Case "2"
		        strInitiatorText = "HP"
		    Case "3"
		        strInitiatorText = "OSSP"
		    Case "4"
		        strInitiatorText = "OTHER"
		    Case Else
		        strInitiatorText = "&nbsp;"
		End Select
		
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
		end if
		rs.Close
		
		if strReqChange = "" then
			strReqChange = ""
		elseif strReqChange  then
			strReqChange = "checked"
		else
			strReqChange = ""
		end if
    
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
			
			if trim(rs("ApproverID")) = "646" then
				blnPreinstallApprover = true
			end if
			strApproverComments = rs("Comments")
			if isPM or trim(CurrentUserID) = trim(strPMID) or trim(CurrentUserID) = trim(strSMID) then
				if ((trim(CurrentUserID) = trim(rs("ApproverID"))) or ((trim(rs("ApproverID")) = "646") and CurrentUserGroup = 15)) and  rs("Status") = "1"  and ApprovalCount = 0 then
					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><span style=""Display:""><font face=verdana size=1><INPUT type=""radio"" id=optApproval name=optApproval LANGUAGE=javascript onclick=""return ApprovalChange(0)"">Approve<BR><INPUT type=""radio"" id=optApproval name=optApproval LANGUAGE=javascript onclick=""return ApprovalChange(1)"">Disapprove<BR><INPUT type=""radio"" id=optApproval name=optApproval LANGUAGE=javascript onclick=""return ApprovalChange(2)"">Not Applicable<BR></span><SELECT style=""display:none"" id=cboApproverStatus name=cboApproverStatus><OPTION selected value=1>Approval Requested</OPTION><OPTION value=2>Approved</OPTION><OPTION value=3>Disapproved</OPTION><OPTION value=5>Not Applicable</OPTION></SELECT></td><TD><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
					strSaveApprovals = trim(rs("ID"))
					ApprovalCount = ApprovalCount + 1
					ApproversPending = ApproversPending & rs("ApproverID") & ","
				else
					if  rs("Status") = "1" then
						ApproversPending = ApproversPending & rs("ApproverID") & ","
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & ">" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
    				elseif ((trim(CurrentUserID) = trim(rs("ApproverID"))) or ((trim(rs("ApproverID")) = "646") and CurrentUserGroup = 15)) and  rs("Status") <> "1" then
    					strSaveApprovals = trim(rs("ID"))
    					'strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap>" & strStatusValue & "<font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & "><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & "><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
                    else
                        strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1><a ID=Status" & rs("ID") & " href=""javascript:ChangeStatus(" & rs("ID") & ")"">" & strStatusText & "</a></font></td><TD><font face=verdana size=1 ID=Comments" & rs("ID") & ">" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
                    end if
				end if
			else
				if ((trim(CurrentUserID) = trim(rs("ApproverID"))) or ((trim(rs("ApproverID")) = "646") and CurrentUserGroup = 15)) and rs("Status") = "1" and ApprovalCount = 0 then
					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><SELECT id=cboApproverStatus name=cboApproverStatus><OPTION selected value=1>Approval Requested</OPTION><OPTION value=2>Approved</OPTION><OPTION value=3>Disapproved</OPTION></SELECT></td><TD><INPUT type=""text"" style= ""width=100%"" id=txtApproverComments maxlength=300 name=txtApproverComments value=""" & strApproverComments & """></font></TD></TR>" 
					strSaveApprovals = trim(rs("ID"))
					ApproversPending = ApproversPending & rs("ApproverID") & ","
					ApprovalCount = ApprovalCount + 1
				else
					if  rs("Status") = "1" then
						ApproversPending = ApproversPending & rs("ApproverID") & ","
    					strApprovals = strApprovals & "<TR><TD style=""Display:none"" ID=Del" & i & ">&nbsp;</TD><TD nowrap><font face=verdana size=1>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1>" & strStatusText & "</font></td><TD><font face=verdana size=1>" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
    				elseif ((trim(CurrentUserID) = trim(rs("ApproverID"))) or ((trim(rs("ApproverID")) = "646") and CurrentUserGroup = 15)) and  rs("Status") <> "1" then
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
			if (isPM or trim(currentuserid) = trim(strpmid) or trim(currentuserid) = trim(strsmid) or trim(currentuserid) = trim(strownerid) or CurrentUserGroup = 15) then
				strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=600 bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD nowrap><font face=verdana size=1><strong>Status</strong></font></TD><TD width=""100%""><font face=verdana size=1><strong>Comments</strong></font></TD></TR>" & strApprovals & "<TR><TD colspan=4><font size=1 face=verdana><a href=""javascript:AddApprover();"">Add Approver</a></font></td></tr></TABLE><BR>"
			else
				strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=600 bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD nowrap><font face=verdana size=1><strong>Status</strong></font></TD><TD width=""100%""><font face=verdana size=1><strong>Comments</strong></font></TD></TR>" & strApprovals & "</TABLE><BR>"
			end if
		end if

	else
		strJustification = ""
		JustificationTemplate = ""		
		
		rs.Open "SELECT OptionConfig, Name, MIN(DisplayOrder) AS DisplayOrder FROM Regions with (NOLOCK) WHERE (Active = 1) GROUP BY OptionConfig, Name ORDER BY OptionConfig",cn,adOpenForwardOnly
		LanguageList = ""
		do while not rs.EOF
			LanguageList = LanguageList &  "<OPTION Value=""" & rs("OptionConfig") & """>" & rs("OptionConfig") & " - " & rs("Name") & "</OPTION>"
			rs.movenext
		loop
		rs.Close

		'dim strCycleProductLinksCons
		'dim strCycleProductLinksComm
		'strCycleProductLinksCons = ""
		'strCycleProductLinksComm = ""
		
		'dim strProductIDLinks
		'strProductIDLinks = ""
		'rs.Open "spGetProgramTree",cn,adOpenForwardOnly
		'Dim LastProgram
		'Dim LastProgramID
		'dim LastBusiness
		'do while not rs.EOF
		'	if LastProgram <> rs("Program") and LastProgram <> "" then
	'			if len(strProductIDLinks) > 1 then
'					strProductIDLinks = mid(strProductIDLinks,2)
'				end if
		'		if LastBusiness = "2" then
		'			strCycleProductLinksCons = strCycleProductLinksCons & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
		'		else
		'			strCycleProductLinksComm = strCycleProductLinksComm & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
		'		end if
'				strProductIDLinks = ""
'			end if
'			strProductIDLinks = strProductIDLinks & "," & rs("ID")
'			LastProgram = rs("Program") & ""
'			lastProgramID = rs("ProgramID")
'			LastBusiness = trim(rs("BusinessID"))
'			rs.MoveNext
'		loop
		
'		if len(strProductIDLinks) > 1 then
'			strProductIDLinks = mid(strProductIDLinks,2)
'		end if
'		if LastBusiness = "2" then
'			strCycleProductLinksCons = strCycleProductLinksCons & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
'		else
'			strCycleProductLinksComm = strCycleProductLinksComm & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
'		end if	
'		rs.Close		
		
		
		
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

	strStatuses = ""

		if strStatus <> "2" and strStatus <> "4" and strStatus <> "5" and strStatus <> "1" then
			strStatuses = "<Option value=0 selected>--- Set Status ---</option>"
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
		if strStatus = "2" then
			strStatuses = strStatuses & "<OPTION value=2 selected>Closed</OPTION>"
			 strStatusText  = "Closed"
		else
			strStatuses = strStatuses & "<OPTION value=2>Closed</OPTION>"
		end if 

		rs.Open "spgetproductsall -2",cn,adOpenForwardOnly
		strproducts = ""
		blnProdFound = false
		do while not rs.EOF
			if (trim(TypeID) = "3" and rs("AllowDCR") ) or ( trim(TypeID) <> "3" and rs("ProductStatusID") < 5) then
				if strProgramID = rs("ID") & "" or ProdID = rs("ID") & "" then
					strproducts = strproducts &  "<OPTION selected Value=""" & rs("ID") & """>" & rs("Name") & " " & rs("Version") & "</OPTION>"
					strProductname = rs("Name") & " " & rs("Version")
					blnprodFound = true
				else
				    strproducts = strproducts &  "<OPTION Value=""" & rs("ID") & """>" & rs("Name") & " " & rs("Version") & "</OPTION>"
				end if

			end if
			rs.movenext
		loop
		rs.Close
		
		if not blnProdFound and IssueID <> "" then
			rs.open "spGetProductVersionName " & strProgramID ,cn,adOpenForwardOnly
			if rs.EOF and rs.EOF then
				strproducts = strproducts &  "<OPTION selected Value=""" & strProgramID & """>Product Name Not Found</OPTION>"		
			else
				strproducts = strproducts &  "<OPTION selected Value=""" & strProgramID & """>" & rs("Name") & "</OPTION>"	
				strProductName = rs("Name") & " " & rs("Version")	
			end if
			rs.Close
		end if


'	if strProgramID = "170" then
'		strproducts = strproducts &  "<OPTION selected Value=""" & "170" & """>" & "Not Assigned" & "</OPTION>"
'	end if
	
	
	
	rs.Open "usp_ListSvcBomAnalyst",cn,adOpenForwardOnly
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
	
	rs.Open "usp_ListGplms",cn,adOpenStatic
	do while not rs.Eof
	    If InStr(strEmployee, "<OPTION Value=""" & rs("ID") & """>" & rs("Name")&  "</OPTION>") = 0 Then
	        strEmployee = strEmployee &  "<OPTION Value=""" & rs("ID") & """>" & rs("Name")&  "</OPTION>"
	    End If
	    rs.MoveNext
	loop
	rs.Close
	
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
	
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	'if currentuserid = 547 and ( SustainingProduct or trim(strCoreTeamRep) = "12" )then
		'isPM = true
	'end if
	
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
	case "7"
		strTypeDisplay = "Service ECR"
		DisplayDistribution = "none"
	case else
		strTypeDisplay = "&nbsp;"
		DisplayDistribution = "none"
	end select
	
	if IssueID <> "" then
		if isPM or trim(CurrentUserID) = trim(strPMID) or trim(CurrentUserID) = trim(strSMID) or CurrentUserSysAdmin or trim(CurrentUserID) = 397 then
			Response.write "<TABLE width=""100%""><TR><TR><TD align=left><H3>" & strTypeDisplay & " Properties</H3></TD><TD align=right><font face=verdana size=1><a href=""javascript:DeleteItem(" & IssueID & ");"">Delete " & strTypeDisplay & "</a></font></TD></TR></TABLE>" 
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
	if (isPM or trim(CurrentUserID) = trim(strPMID) or trim(CurrentUserID) = trim(strSMID) or CurrentUserSysAdmin) and strOnlineReports = "1"  then
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
	if not (isPM or trim(currentuserid) = trim(strpmid) or trim(currentuserid) = trim(strsmid)  or CurrentUserSysAdmin) then
		if ((not SustainingProduct) and trim(strCoreTeamRep) <> "12") or (trim(strECNDate) <> "" or trim(strStatus) = "5") then
			Response.Write "<font size=2 face=verdana color=red><b>This DCR is closed and can only be edited by the PM.</b></font><BR><BR>"
			strRecordLocked = "1"
		end if
	end if
end if


	noApprovals = false
	if ((isPM or trim(currentuserid) = trim(strpmid) or trim(currentuserid) = trim(strsmid) or trim(currentuserid) = trim(strownerid)) and strApprovals = "") and strProgramID <> "170"  then
		strApprovals = "<table ID=ApproverTable border=1 cellPadding=2 cellSpacing=0 width=100% bgcolor=ivory bordercolor=tan><TR bgcolor=cornsilk><TD ID=DeleteCell style=""Display:none"" width=10><font size=1 face=verdana><a href=""javascript:DeleteApprover();"">Delete</a></font></TD><TD width=160 nowrap><font face=verdana size=1><strong>Approver</strong></font></TD><TD><font face=verdana size=1><strong>Status</strong></font></TD><TD><font face=verdana size=1><strong>Comments</strong></font></TD></TR><TR><TD colspan=4><font size=1 face=verdana><a href=""javascript:AddApprover();"">Add Approver</a></font>" & strApprovalListLink & "<font face=verdana size=1>&nbsp;</font></td></tr></TABLE><BR>" '|&nbsp;<a href=""javascript:UpdateApproverList();"">Save Product Approver List</a>
		'noApprovals = true
	elseif strApprovals = "" then
		Response.write "<Table style=""Display:none;WIDTH=100%"" ID=ApproverTable></TABLE>"	
	end if	

	if (not noapprovals) and trim(TypeID) <> "4" then
		Response.write strApprovals
	elseif trim(TypeID) = "4" and strID <> "" then
		Response.Write "<Table style=""Display:none"" ID=ApproverTable></TABLE>"
	end if

	
    %>
    <table border="1" cellpadding="2" cellspacing="0" width="100%" bgcolor="cornsilk"
        bordercolor="tan">
        <tr style="display: <%=DisplayForAdd%>">
            <td width="160" style="vertical-align: top">
                <strong><font size="2">ID:</font></strong>
            </td>
            <td>
                <font size="2">
                    <%=strID%>
                </font>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Submitter:</font></strong>
            </td>
            <td>
                <font size="2" face="verdana">
                    <%=strSubmitter%>
                </font>
                <input id="txtSubmitter" name="txtSubmitter" type="hidden" value="<%=strSubmitter%>">
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Date Submitted:</font></strong>
            </td>
            <td>
                <font size="2" face="verdana">
                    <%=strSubmitted%>
                </font>
            </td>
        </tr>
        <%if strStatus = "2" or strStatus = "5" then%>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Date&nbsp;Closed:</font></strong>
            </td>
            <td colspan="2" valign="top">
                <font size="2" face="verdana">
                    <%=strActual%>
                </font>
            </td>
        </tr>
        <%end if%>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Summary:</font><font color="red" size="1"> *</font></strong>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtSummary" name="txtSummary" style="width: 100%;" maxlength="120"
                    value="<%=strsummary%>">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strSummary), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr>
            <td>
                <span style="font: bold x-small">SPS Kit PN:<span style="color: Red; font-size: xx-small">*</span></span>
            </td>
            <td>
                <div style="display: <% If strStatus = "2" Or strStatus = "5" Then Response.Write "none" %>">
                    <input type="text" id="txtSpareKitNo" name="txtSpareKitNo" style="width: 100%;" maxlength="500"
                        value="<%= strSpareKitPn %>" />
                </div>
                <div style="display: <% If Not (strStatus = "2" Or strStatus = "5") Then Response.Write "none" %>;">
                    <%= server.HTMLEncode(strSpareKitPn) %></div>
            </td>
        </tr>
        <tr>
            <td>
                <span style="font: bold x-small">SA PN:<span style="color: Red; font-size: xx-small">*</span></span>
            </td>
            <td>
                <div style="display: <% If strStatus = "2" Or strStatus = "5" Then Response.Write "none" %>">
                    <input type="text" id="txtSaNo" name="txtSaNo" style="width: 100%;" maxlength="500"
                        value="<%= strSubAssemblyPn %>" />
                </div>
                <div style="display: <% If Not (strStatus = "2" Or strStatus = "5") Then Response.Write "none" %>;">
                    <%= server.HTMLEncode(strSubAssemblyPn) %></div>
            </td>
        </tr>
        <tr id="RowProgram">
            <td style="vertical-align: top">
                <strong><font size="2">Program(s):</font><font color="red" size="1"> *</font></strong>
            </td>
            <td>
                <select size="2" id="lstProducts" name="lstProducts" style="width: 180px; height: 121px"
                    multiple="multiple">
                    <%=strproducts%>
                </select>
                <span style="color: Green; font: verdana xx-small;">
                    <br />
                    Use CRTL or SHIFT to multi-select</span>
            </td>
        </tr>
        <%if ((isPM or trim(currentuserid) = trim(strpmid) or trim(currentuserid) = trim(strsmid)  or (trim(currentuserid) = trim(strownerid) and ( trim(strCoreTeamRep) = "12" or SustainingProduct) ) or CurrentUserSysAdmin ) or (TypeID <> "3")) then%>
        <!-- and (NOT (strStatus = "2" Or strStatus = "5")) -->
        <tr style="display: <%=DisplayForAdd%>">
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Status:</font><font color="red" size="1"> *</font></strong>
            </td>
            <td colspan="2">
                <select id="cboStatus" name="cboStatus" style="width: 180px;">
                    <%=strStatuses%>
                </select>
            </td>
        </tr>
        <%else%>
        <tr>
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Status:</font><font color="red" size="1"> *</font></strong>
            </td>
            <td colspan="2">
                <font face="verdana" size="2">
                    <%=strStatusText%>
                </font>
                <select id="cboStatus" name="cboStatus" style="display: none; width: 180px;" language="javascript"
                    onchange="return cboStatus_onchange()">
                    <%=strStatuses%>
                </select>
            </td>
        </tr>
        <%end if%>
        <tr style="display: <%=DisplayForAdd%>">
            <td width="160" style="vertical-align: top">
                <strong><font size="2">Assigned To:</font><font color="red" size="1"> *</font></strong>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <div style="display: none">
                    <%End If %>
                    <select id="cboOwner" name="cboOwner" style="width: 180px;" language="javascript"
                        onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()"
                        onkeydown="return combo_onkeydown()">
                        <%=strowners%>
                    </select>&nbsp;<input type="button" value="Add" id="cmdOwnerAdd" name="cmdOwnerAdd"
                        language="javascript" onclick="return cmdOwnerAdd_onclick()">
                    <% If strStatus = "2" Or strStatus = "5" Then %>
                </div>
                <%= strOwner %>
                <%End If %>
            </td>
        </tr>
        <tr>
            <td width="160" style="vertical-align: top">
                <span style="font-size: x-small; font-weight: bold">Description: <span style="color: red;
                    font-size: xx-small">*</span></span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <div style="display: none">
                    <%End If %>
                    <textarea id="txtDetails" style="width: 100%; height: 120px" name="txtDetails" rows="3"><%= strDetails%></textarea>
                    <% If strStatus = "2" Or strStatus = "5" Then %>
                </div>
                <%= replace(server.HTMLEncode(strDetails), vbcrlf, "<BR>") %>
                &nbsp;<%End If %>
            </td>
        </tr>
        <tr>
            <td width="160" style="vertical-align: top">
                <span style="font-size: x-small; font-weight: bold">Attachment: </span>
            </td>
            <td colspan="2">
                <%if trim(strAttachment)="" then%>
                <span id="AddAttachmentLinks">
                    <%else%>
                    <span id="AddAttachmentLinks" style="display: none">
                        <%end if%>
                        <a href="javascript:AddAttachment();">Add</a>&nbsp;<font size="1" color="green" face="verdana">(One
                            file only)</font></span>
                    <%if trim(strAttachment)="" then%>
                    <span id="UpdateAttachmentLinks" style="display: none">
                        <%else%>
                        <span id="UpdateAttachmentLinks">
                            <%end if%>
                            <a href="javascript: ViewTestReport();">View</a>&nbsp;|&nbsp;<a href="javascript:AddAttachment();">Replace</a>&nbsp;|&nbsp;
                            <a href="javascript:RemoveAttachment();">Remove</a> </span>
                        <input type="hidden" id="txtAttachment" name="txtAttachment" value="<%=strAttachment%>" />
            </td>
        </tr>
        <tr>
            <td>
                <span style="font: bold x-small">Recommended Inventory Disposition:</span>
            </td>
            <td>
                <div style="display: <% If strStatus = "2" Or strStatus = "5" Then Response.Write "none" %>">
                    <select id="selInventoryDisposition" name="selInventorDisposition">
                        <option value="0">--- Select Inventory Disposition ---</option>
                        <option value="1" <% If strInventoryDispositionId = "1" Then Response.Write "Selected" %>>
                            No OSSP Inventory Affected</option>
                        <option value="2" <% If strInventoryDispositionId = "2" Then Response.Write "Selected" %>>
                            Use current material until depleted, then roll to new material</option>
                        <option value="3" <% If strInventoryDispositionId = "3" Then Response.Write "Selected" %>>
                            Purge (Quality Issue - OTHER ACTION IS REQUIRED!)</option>
                    </select>
                </div>
                <div style="display: <% If Not (strStatus = "2" Or strStatus = "5") Then Response.Write "none" %>;">
                    <%= server.HTMLEncode(strInventoryDisposition) %></div>
            </td>
        </tr>
        <tr>
            <td>
                <span style="font: bold x-small">Initiated By:</span>
            </td>
            <td>
                <div style="display: <% If strStatus = "2" Or strStatus = "5" Then Response.Write "none" %>">
                    <select id="selInitiator" name="selInitiator">
                        <option value="0">--- Select Inventory Disposition ---</option>
                        <option value="1" <% If strInitiator = "1" Then Response.Write "Selected" %>>ODM</option>
                        <option value="2" <% If strInitiator = "2" Then Response.Write "Selected" %>>HP</option>
                        <option value="3" <% If strInitiator = "3" Then Response.Write "Selected" %>>OSSP</option>
                        <option value="4" <% If strInitiator = "3" Then Response.Write "Selected" %>>Other</option>
                    </select>
                    <input type="text" id="txtInitiatorOther" name="txtInitiatorOther" style="display: none;" />
                </div>
                <div style="display: <% If Not (strStatus = "2" Or strStatus = "5") Then Response.Write "none" %>;">
                    <%= server.HTMLEncode(strInitiatorText) %></div>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">BOM Analyst Comments:</span>
            </td>
            <td colspan="2">
            <textarea id="txtBomAnalystComments" name="txtBomAnalystComments" style="width: 100%; height: 120px" rows="2" cols="80"><%= strBomAnalystComments %></textarea>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">New QSpecs Submitted:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtQSpecDt" name="txtQSpecDt" style="width: 80%;" maxlength="120"
                    value="<%=strQSpecDt%>">
                                                    <a href="javascript: cmdDate_onclick('txtQSpecDt')">
                                <img id="picZsrpReadyTargetDt" src="/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>

                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strQSpecDt), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">Component ECO Submitted:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtCompEcoDt" name="txtCompEcoDt" style="width: 80%;" maxlength="120"
                    value="<%=strCompEcoDt%>">                                <a href="javascript: cmdDate_onclick('txtCompEcoDt')">
                                <img id="Img1" src="/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>

                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strCompEcoDt), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">Component ECO No.:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtCompEcoNo" name="txtCompEcoNo" style="width: 80%;" maxlength="120"
                    value="<%=strCompEcoNo%>">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strCompEcoNo), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">SA ECO Submitted:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtSaEcoDt" name="txtSaEcoDt" style="width: 80%;" maxlength="120"
                    value="<%=strSaEcoDt%>">                                <a href="javascript: cmdDate_onclick('txtSaEcoDt')">
                                <img id="Img2" src="/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>

                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strSaEcoDt), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">SA ECO No.:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtSaEcoNo" name="txtSaEcoNo" style="width: 80%;" maxlength="120"
                    value="<%=strSaEcoNo%>">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strSaEcoNo), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">SPS ECO Submitted:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtSpsEcoDt" name="txtSpsEcoDt" style="width: 80%;" maxlength="120"
                    value="<%=strSpsEcoDt%>">                                <a href="javascript: cmdDate_onclick('txtSpsEcoDt')">
                                <img id="Img3" src="/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21" /></a>

                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strSpsEcoDt), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
        <tr style="display: <%=DisplayForAdd%>">
            <td>
                <span style="font: bold x-small">SPS ECO No.:</span>
            </td>
            <td colspan="2">
                <% If strStatus = "2" Or strStatus = "5" Then Response.Write "<div style=""display: none"">" End If%>
                <input type="text" id="txtSpsEcoNo" name="txtSpsEcoNo" style="width: 80%;" maxlength="120"
                    value="<%=strSpsEcoNo%>">
                <% If strStatus = "2" Or strStatus = "5" Then %>
                <%= "</div>"%>
                <%= replace(Server.HTMLEncode(strSpsEcoNo), vbcrlf, "<BR>")%>
                <%End If %>
            </td>
        </tr>
    </table>
    <%
	if noapprovals then
		Response.write "<BR>" & strApprovals
	end if
end if
    %>
    <textarea style="display: none" rows="2" cols="20" id="txtJustificationTemplate"
        name="txtJustificationTemplate"><%=JustificationTemplate%></textarea>
    <input style="display: none" type="text" id="txtID" name="txtID" value="<%=IssueID%>" />
    <input type="hidden" id="txtType" name="txtType" value="<%=TypeID%>" />
    <input type="hidden" id="Approvers2Add" name="Approvers2Add" value="" />
    <input style="display: none" type="text" id="txtSaveApproval" name="txtSaveApproval"
        value="<%=strSaveApprovals%>" />
    <input style="display: none" type="text" id="txtCurrentUserID" name="txtCurrentUserID"
        value="<%=CurrentUSerID%>" />
    <input style="display: none" type="text" id="txtApproversLoaded" name="txtApproversLoaded"
        value="<%=ApproversLoaded%>" />
    <input style="display: none" type="text" id="txtApproversPending" name="txtApproversPending"
        value="<%=ApproversPending%>" />
    <input style="display: none" type="text" id="txtCommodityManagerID" name="txtCommodityManagerID"
        value="<%=strCommodityManagerID%>" />
    <select id="cboEmployee" name="cboEmployee" style="display: none; width: 180px;">
        <%=strEmployee%>
    </select>
    <br />
    <div id="divApproverList">
        <select style="display: none" id="cboProdApprovalList" name="cboProdApprovalList">
            <%=strProdApprovalList%>
        </select>
    </div>
    <input type="hidden" id="txtProductName" name="txtProductName" value="<%=strProductname%>" />
    <input type="hidden" id="txtProductID" name="txtProductID" value="<%=strProgramID%>" />
    <input type="hidden" id="hidProdId" name="hidProdId" value="<%=ProdId %>" />
    <input type="hidden" id="hidAddDCRNotificationList" name="hidAddDCRNotificationList" value="<%=AddDCRNotificationList%>" />
</form>
    <input type="hidden" id="txtRecordLocked" name="txtRecordLocked" value="<%=strRecordLocked%>" />

</body>
</html>
