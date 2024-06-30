<%@ Language=VBScript %>
<!-- #include file = "../../includes/noaccess.inc" -->
<HTML>
<head>
    <title>Qual Status Main</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
    <!--
    <!-- #include file = "../../_ScriptLibrary/sort.js" -->

    var KeyString = "";
    var origStatus = "";
    var origDCR = "";
    var origConfigRestriction = "";
    var origSupplyRestricted = "";
    var origComment = "";
    var origSubassemblies = [];
    var origDate = "";
    var origConfidence = "";

    $(document).ready(function () {
        origStatus = $("#cboStatus").val();
        origDCR = $("#cboWhy").val();
        origConfidence = $("#cboConfidence").val();
        origSupplyRestricted = $("#chkSupplyRestricted").val();
        origConfigRestriction = $("#chkConfigurationRestricted").val();
        origComment = $("#txtComments").val();
        origDate = $("#txtTestDate").val();

        $('[name="lstSub"]:checked').each(function (i) {
            origSubassemblies[i] = $(this).val();
        });
        window.parent.frames["LowerWindow"].enableButton();
        $("#txtRedirect").val("QualStatusMainPulsar.asp?ProdID=" + $("#txtProdID").val() + "&VersionID=" + $("#txtVersionID").val() + "&ProductDeliverableReleaseID=" + $("#txtProdDelRelID").val() + "&TodayPageSection=" + $("#txtTodayPageSection").val());
    });

    function combo_onkeypress() {
        if (event.keyCode == 13) {
            KeyString = "";
        }
        else {
            KeyString = KeyString + String.fromCharCode(event.keyCode);
            event.keyCode = 0;
            var i;
            var regularexpression;

            for (i = 0; i < event.srcElement.length; i++) {
                regularexpression = new RegExp("^" + KeyString, "i")
                if (regularexpression.exec(event.srcElement.options[i].text) != null) {
                    event.srcElement.selectedIndex = i;
                    break;
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

    function mouseover_Column() {
        event.srcElement.style.color = "red";
        event.srcElement.style.cursor = "hand";

    }
    function mouseout_Column() {
        event.srcElement.style.color = "black";
    }

    function cmdDate_onclick(FieldID) {
        var strID;

        strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", frmStatus.txtTestDate.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) != "undefined")
            frmStatus.txtTestDate.value = strID;
    }

    function Left(str, n) {
        if (n <= 0)     // Invalid bound, return blank string
            return "";
        else if (n > String(str).length)   // Invalid bound, return
            return str;                // entire string
        else // Valid bound, return appropriate substring
            return String(str).substring(0, n);
    }


    function window_onload() {
        //frmStatus.cboStatus.focus();
    }

    function cboStatus_onchange() {

        var strStatus = frmStatus.cboStatus.options[frmStatus.cboStatus.selectedIndex].text;
        $("#txtStatusText").val(strStatus);
        var strCommentsRequired = txtCommentsRequired.value.indexOf("," + frmStatus.cboStatus.options[frmStatus.cboStatus.selectedIndex].value + ",");
        if (frmStatus.cboStatus.options[frmStatus.cboStatus.selectedIndex].value == "3")
            DateRow.style.display = "";
        else
            DateRow.style.display = "none";

        if (strStatus == "FCS" || strStatus == "OOC" || frmStatus.cboStatus.options[frmStatus.cboStatus.selectedIndex].value == "3")
            ConfidenceRow.style.display = "";
        else
            ConfidenceRow.style.display = "none";


        if (strStatus == "QComplete")
            frmStatus.chkRiskRelease.checked = false;
        else if (strStatus == "Risk Release")
            frmStatus.chkRiskRelease.checked = true;



        if (frmStatus.cboStatus.selectedIndex == 0) {
            frmStatus.cboWhy.selectedIndex = 0;
            frmStatus.cboDCR.selectedIndex = 0;
            DCRRow.style.display = "none";
            SupportRow.style.display = "";
        }
        else {
            frmStatus.chkDelete.checked = false;
            SupportRow.style.display = "none";
        }

        if (strCommentsRequired == -1)
            CommentStar.style.display = "none";
        else
            CommentStar.style.display = "";


    }

    function cboWhy_onchange() {
        if (frmStatus.cboWhy.selectedIndex == 3)
            DCRRow.style.display = "";
        else
            DCRRow.style.display = "none";
    }

    function SwitchRelease(ProdID, VersionID, ProductDeliverableReleaseID) {
        var isModified = false;
        $("#txtKeepItOpen").val(true);

        if (origStatus != $("#cboStatus").val()) {
            isModified = true;
        }
        else if (origDCR != $("#cboWhy").val()) {
            isModified = true;
        }
        else if (origConfidence != $("#cboConfidence").val()) {
            isModified = true;
        }
        else if (origSupplyRestricted != $("#chkSupplyRestricted").val()) {
            isModified = true;
        }
        else if (origConfigRestriction != $("#chkConfigurationRestricted").val()) {
            isModified = true;
        }
        else if (origComment != $("#txtComments").val()) {
            isModified = true;
        }
        else if (origDate != $("#txtTestDate").val()) {
            isModified = true;
        }
        else {
            var count = 0;
            $('[name="lstSub"]:checked').each(function (i) {
                if (jQuery.inArray($(this).val(), origSubassemblies) == -1) {
                    isModified = true;
                    return false;
                }
                count++;
            });

            if (count < origSubassemblies.length)
                isModified = true;
        }

        if (isModified) {
            if (confirm("Do you want to save your changes for this Release?")) {
                $("#txtRedirect").val("QualStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val());
                window.parent.frames["LowerWindow"].cmdOK_onclick();
            }
            else {
                document.location = "QualStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val();
                window.parent.repositionParentWindow();
            }
        }
        else {
            document.location = "QualStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionID=" + VersionID + "&ProductDeliverableReleaseID=" + ProductDeliverableReleaseID + "&TodayPageSection=" + $("#txtTodayPageSection").val() + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val();
            window.parent.repositionParentWindow();
        }
    }
    //-->
    </script>
    <link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">
    <style>
        A:visited {
            COLOR: blue;
        }

        A:hover {
            COLOR: red;
        }

        .DelTable TBODY TD {
            BORDER-TOP: gray thin solid;
        }
    </style>
</head>

<body bgcolor="ivory" language="javascript" onload="return window_onload()">


    <%
	dim cn
	dim rs
	dim cm
	dim p
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim strStatus
	dim strDate
	dim strDeliverable
	dim strID
	dim strPartNumber
	dim strProduct
	dim strVendor
	dim strVendorID
	dim strRootID
	dim strCategory
	dim strCategoryID
	dim blnAdmin
	dim blnFound
	dim strStatusList
	dim strStatusText
	dim strStatusID
	dim strDevStatus
	dim strDCR
	dim strComments
	dim strWhy
	dim strHW
	dim strFW
	dim strRev
	dim strModel
	dim strStatusSelected
	dim strSubsLoaded
	dim CurrentUserEmail
	dim strQCompleteSubject
	dim strQCompleteBody
	dim strFailSubject
	dim strFailBody
	dim strPMs
	dim strDevEmail
	dim strQCompleteCount
	dim strDevCenter
	dim strRows
	dim strCommentsRequired
	dim strSupplyRestricted
	dim strConfigurationRestricted
	dim strRiskRelease
	dim strSupplyRestrictionID
	dim strConfigurationRestrictionID
	dim strRestrictBody
	dim strPartnerID
	dim strDeveloperTestStatus 
	dim strDeveloperTestStatusName 
	dim strIntegrationTestStatus 
	dim strODMTestStatus 
	dim strDeveloperTestStatusCheck 
	dim strIntegrationTestStatusCheck 
	dim strODMTestStatusCheck 
	dim strDeveloperTestNotes 
	dim strIntegrationTestNotes 
	dim strODMTestNotes 
	dim strWWANTestNotes
	dim strTTS
	dim strWWANCell
	dim strDevCell
	dim strMITCell
	dim strODMCell
    dim strTTSCell
	dim strBridgedIDs
	dim TestStatusArray 
	dim blnRequiresWWANSignoff
	dim blnRequiresODMSignoff
	dim blnRequiresMITSignoff
	dim blnRequiresDeveloperSignoff
	dim blnTestingComplete
	dim strTestLeadsTo
	dim blnServicePM
    dim strReleaseLink
    dim pvid
    dim isPulsarProduct	
    dim strSql
    dim ProductDeliverableReleaseID
    dim ProductDeliverableReleaseName
    dim ProductDeliverableID

    ProductDeliverableReleaseID = 0
    ProductDeliverableReleaseName = ""

    strStatusID = 0
  	TestStatusArray = split("TBD,Passed,Failed,Blocked,Watch,N/A",",")
    
    blnTestingComplete = 1
    
	strSupplyRestrictionID=""
	strConfigurationRestrictionID=""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
    'Get Product
    if trim(Request("ProdID") & "") = "" then
        pvid = 0
    else
	    pvid = clng(Request("ProdID"))
    end if

    isPulsarProduct =0
    rs.Open "select COALESCE( FusionRequirements,0) as FusionRequirements from ProductVersion where ID=" & pvid & ";",cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			isPulsarProduct =0	
		else
			isPulsarProduct = cint(rs("FusionRequirements"))
		end if
    rs.Close
	
    dim intDefaultReleaseID
    if Request("ReleaseID") then
        intDefaultReleaseID = clng(Request("ReleaseID"))	    
    end if
    
	if (isPulsarProduct=1) and (intDefaultReleaseID = 0) and (Request("TodayPageSection") = "") then
		strSql = "select top 1 pr.ID, pr.Name from ProductVersion_Release pvr join ProductVersionRelease pr on pr.ID = pvr.ReleaseID where pvr.ProductVersionID= " & clng(Request("ProdID")) & " order by pr.ReleaseYear desc, pr.ReleaseMonth desc;"
		rs.open strSql,cn
		if not (rs.EOF and rs.BOF) then
            intDefaultReleaseID = rs("ID")
		end if	
		rs.close
	end if

    if (isPulsarProduct=1) then

        if Request("ProductDeliverableReleaseID") then
            ProductDeliverableReleaseID = trim(Request("ProductDeliverableReleaseID"))            
        end if

        if Request("TodayPageSection") = "" then
            strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID and pdr.targeted = pd.targeted " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & pvid & " and pd.DeliverableVersionID= " & Request("VersionID") & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
        else 
            strSql = "select pvr.Name, ReleaseID = pvr.ID, pdr.ID " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & pvid & " and pd.DeliverableVersionID= " & Request("VersionID")

            if ProductDeliverableReleaseID > 0 then
                strSql = strSql & " and pdr.id = " & ProductDeliverableReleaseID & " order by pvr.id desc"
            else
                if intDefaultReleaseID > 0 then
                    strSql = strSql & " and pdr.id = " & intDefaultReleaseID & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
                else 
                    strSql = strSql & " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
                end if
            end if
        end if

	    rs.open strSql, cn
        	    
        strReleaseLink = "Switch Releases:&nbsp;"

        Do until rs.EOF            
        
            if strReleaseLink <> "Switch Releases:&nbsp;" then
                strReleaseLink = strReleaseLink & " | " 
            end if

            if ProductDeliverableReleaseID > 0 and ProductDeliverableReleaseID = trim(rs("ID")) then
                strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                ProductDeliverableReleaseName = rs("Name")
            else
                if rs("ReleaseID") = intDefaultReleaseID and ProductDeliverableReleaseID = 0 then
                    strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                    ProductDeliverableReleaseID = rs("ID")
                    ProductDeliverableReleaseName = rs("Name")
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & pvid & "," & Request("VersionID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
                end if
            end if
                
            rs.MoveNext
        Loop

        if ProductDeliverableReleaseID = 0 then
            strReleaseLink = "Switch Releases:&nbsp;"
            dim count
            count = 0
            rs.MoveFirst
            Do until rs.EOF            
        
                if strReleaseLink <> "Switch Releases:&nbsp;" then
                    strReleaseLink = strReleaseLink & " | " 
                end if

                if  count = 0 then
                    strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                    ProductDeliverableReleaseID = rs("ID")
                    ProductDeliverableReleaseName = rs("Name")
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & pvid & "," & Request("VersionID") & "," & rs("ID") & ");"">" & rs("Name") & "</a>"
                end if
                
                count = count + 1
                rs.MoveNext
            Loop
        end if
        rs.Close        
    end if

	'Get User
	dim CurrentDomain
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
	Set rs = cm.Execute 
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("email") & ""
		blnServicePM = rs("ServicePM")
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
		blnServicePM = false
	end if
	rs.Close
	
	if currentuserid=31 or currentuserid=8 then
		blnadmin=true
	else
		blnadmin=false
	end if
	
	if blnServicePM and trim(request("ProdID")) <> ""then
	
		rs.open "spGetHardwareTeamAccessList " & CurrentUserID & "," & clng(request("ProdID")),cn,adOpenStatic
		do while not rs.EOF
			if rs("HWTeam") = "ProgramCoordinator" or blnEngineeringCoordinator > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "PlatformDevelopment" and rs("Products") > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "Processor" and rs("Products") > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "Comm" and rs("Products") > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "GraphicsController" and rs("Products") > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "VideoMemory" and rs("Products") > 0 then
				blnServicePM = false
			elseif rs("HWTeam") = "SuperUser" and rs("Products") > 0 then
				blnServicePM = false
			'elseif rs("HWTeam") = "Commodity" and rs("Products") > 0 then
			'	blnServicePM = false
			end if
			rs.MoveNext			
		loop
		rs.Close
	
	end if
	
    ProductDeliverableID = ""
	strID=""
	strStatus = ""
	strDate = ""
	strVendor=""
	strVendorID = 0
	strRootID=0
	strCategory = ""
	strCategoryID=0
	strPartNumber=""
	strDeliverable = ""
	strProduct= ""
	strStatusList = ""
	strDevStatus = ""
	strDCR = "0"
	strComments = ""
	strHW = "&nbsp;"
	strFW = "&nbsp;"
	strRev = "&nbsp;"
	strModel = "&nbsp;"
	blnFound = false
	strStatusSelected = ""
	strFailSubject = ""
	strFailBody = ""
	strPMs = ""
	strQCompleteCount = ""
	strDevCenter = ""
	strRows = ""
	strRiskRelease = ""
	strSupplyRestricted = ""
	strConfigurationRestricted = ""
	strPartnerID = ""
	strDeveloperTestStatus =  ""
	strDeveloperTestStatusName =  ""
	strDeveloperTestNotes =  ""
	strIntegrationTestStatus =  ""
	strIntegrationTestNotes =  ""
	strODMTestStatus = ""
	strODMTestNotes = ""
	strWWANTestStatus = ""
	strWWANTestNotes = ""
	strDeveloperTestStatusCheck =  ""
	strIntegrationTestStatusCheck =  ""
	strODMTestStatusCheck = ""
	strTTS = ""
	strWWANCell=""
	strDevCell=""
	strMITCell=""
	strWWANCell=""
	strODMCell=""
    strTTSCell=""
    strBridgedIDs = ""
    strTestLeadsTo = ""	
    

	if request("ProdID") = "" or request("VersionID") = "" then
		Response.Write "Not enough information supplied to process your request."
	else
		
        rs.Open "spGetCommodityStatusRelease " & clng(request("ProdID")) & "," & clng(request("VersionID")) & "," & ProductDeliverableReleaseID, cn, adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strNumber = ""
		else
			blnFound = true			            
            ProductDeliverableID = rs("ProductDeliverableID")
			strID = rs("ID") & ""
			if trim(rs("SupplyChainRestriction")) = "1" then
				strSupplyRestricted = "checked"
				strSupplyRestrictionID = "1"
			end if			
			if trim(rs("ConfigurationRestriction")) = "1" then
				strConfigurationRestricted = "checked"
				strConfigurationRestrictionID = "1"
			end if			
			strRiskRelease = rs("RiskRelease")
			strProduct = rs("Product") & ""
			strDevCenter = trim(rs("DevCenter") & "")
			strDeliverable = rs("DeliverableName") & ""
			strStatus = rs("TestStatusID") & ""
			strPartNumber = rs("PartNumber") & ""
			strModel = rs("ModelNumber") & "&nbsp;"
			strDate = rs("TestDate") & ""
			if isdate(strDate) then
				strDate = formatdatetime(strDate,vbshortdate)
			end if
			strVendor=rs("Vendor") & ""
			strCategory = rs("Category")& ""
			strCategoryID = rs("CategoryID") & ""
			strComments = server.HTMLEncode(rs("TargetNotes") & "")
			strDCR = trim(rs("DCRID") & "")
			strHW = rs("Version") & "&nbsp;"
			strFW = rs("Revision") & "&nbsp;"
			strRev = rs("Pass") & "&nbsp;"
			strVendorID = rs("VendorID") & ""
			strRootID = rs("DeliverableRootID") & ""
			strDevStatus = rs("developernotificationstatus") & ""
			strPartnerID = rs("PartnerID") & ""
			strConfidence = trim(rs("TestConfidence") & "")
			strDeveloperTestStatus = rs("developerteststatus") & ""
			strDeveloperTestNotes = rs("developerTestNotes") & ""
			strIntegrationTestStatus = rs("Integrationteststatus") & ""
			strIntegrationTestNotes = rs("IntegrationTestNotes") & ""
			strODMTestStatus = rs("ODMteststatus") & ""
			strODMTestNotes = rs("ODMtestNotes") & ""
			strWWANTestStatus = rs("WWANTestStatus") & ""
			strWWANTestNotes = rs("WWANTestNotes") & ""
			strTTS = rs("TTS") & ""

			if rs("WWANProduct") and rs("requiresWWANtestfinalapproval") then
				blnRequiresWWANSignoff = true
			else
				blnRequiresWWANSignoff = false
			end if
			
			blnRequiresODMSignoff = rs("requiresodmtestfinalapproval")
			blnRequiresMITSignoff = rs("requiresMITtestfinalapproval")
			blnRequiresDeveloperSignoff = rs("requiresdeveloperfinalapproval")

            blnTestingComplete = 1
            if clng(blnRequiresODMSignoff) = -1 and trim(strODMTestStatus) <> "1"  and trim(strODMTestStatus) <> "5" then
               blnTestingComplete = 0
            end if
            if clng(blnRequiresWWANSignoff) = -1 and trim(strWWANTestStatus) <> "1"  and trim(strWWANTestStatus) <> "5" then
               blnTestingComplete = 0
            end if
            if clng(blnRequiresMITSignoff) = -1 and trim(strIntegrationTestStatus) <> "1"  and trim(strIntegrationTestStatus) <> "5" then
               blnTestingComplete = 0
            end if
            if clng(blnRequiresDeveloperSignoff) = -1 and trim(strDeveloperTestStatus) <> "1"  and trim(strDeveloperTestStatus) <> "5" then
               blnTestingComplete = 0
            end if
            
            if blnTestingComplete =1 and lcase(trim(rs("TTS") & "")) = "pending" then
                blnTestingComplete = 2
            elseif blnTestingComplete =0 and lcase(trim(rs("TTS") & "")) = "pending" then
                blnTestingComplete = 3
            end if
            
			if rs("DeveloperNotificationStatus")=0 and trim(strDeveloperTestStatus) = "0" then
				strDevStatus="Awaiting Development Team Approval"
			elseif  rs("DeveloperNotificationStatus")=1 and trim(strDeveloperTestStatus) = "0" then
				strDevStatus="Approved For Testing"
			elseif  rs("DeveloperNotificationStatus")=1 and trim(strDeveloperTestStatus) = "0" then
				strDevStatus="Not Approved For Testing"
			elseif  trim(strDeveloperTestStatus)="1" then
				strDevStatus="Approved For Production"
			elseif  trim(strDeveloperTestStatus)="2" then
				strDevStatus="Not Approved For Production"
			else
				strDevStatus="TBD"
			end if
		end if
		rs.Close

		rs.open "spListTestLeads4ProductDeliverable " & clng(request("ProdID")) & "," & clng(request("VersionID")),cn,adOpenForwardOnly
        strTestLeadsTo = ""
        do while not rs.eof
			if clng(blnRequiresODMSignoff) = -1 and rs("Role") = "ODM" and trim(strODMTestStatus) <> "1"  and trim(strODMTestStatus) <> "5" then
			    strTestLeadsTo = strTestLeadsTo & ";" & rs("email")
			elseif clng(blnRequiresMITSignoff) = -1 and rs("Role") = "MIT" and trim(strIntegrationTestStatus) <> "1"  and trim(strIntegrationTestStatus) <> "5" then
			    strTestLeadsTo = strTestLeadsTo & ";" & rs("email")
			elseif clng(blnRequiresWWANSignoff) = -1 and rs("Role") = "WWAN" and trim(strWWANTestStatus) <> "1"  and trim(strWWANTestStatus) <> "5" then
			    strTestLeadsTo = strTestLeadsTo & ";" & rs("email")
			elseif clng(blnRequiresWWANSignoff) = -1 and rs("Role") = "WWAN" and lcase(trim(strTTS)) = "pending" then
			    strTestLeadsTo = strTestLeadsTo & ";" & rs("email")
			elseif clng(blnRequiresDeveloperSignoff) = -1 and rs("Role") = "DEV"  and trim(strDeveloperTestStatus) <> "1"  and trim(strDeveloperTestStatus) <> "5" then
			    strTestLeadsTo = strTestLeadsTo & ";" & rs("email")
			end if
            rs.movenext 
        loop
        rs.close
        if strTestLeadsTo <> "" then
            strTestLeadsTo = mid(strTestLeadsTo,2)
        end if

			
		if not blnRequiresWWANSignoff then
			strWWANCell = "N/A"
	        strTTSCell = "N/A"
		else
			strWWANCell = TestStatusArray(clng(strWWANTestStatus))
	        strTTSCell = strTTS
		end if
		if not blnRequiresODMSignoff then
			strODMCell = "N/A"
		else
			strODMCell = TestStatusArray(clng(strODMTestStatus))
		end if
		if not blnRequiresDeveloperSignoff then
			strDEVCell = "N/A"
		else
			strDEVCell = TestStatusArray(clng(strDeveloperTestStatus))
		end if
		if not blnRequiresMITSignoff then
			strMITCell = "N/A"
		else
			strMITCell = TestStatusArray(clng(strIntegrationTestStatus))
		end if
	
		if strIntegrationTestNotes <> "" then
			strIntegrationTestNotes = " - " & strIntegrationTestNotes
		end if
		if not blnRequiresMITSignoff then
			strIntegrationTestStatus = "N/A"
		else
			strIntegrationTestStatus = teststatusarray(clng(strIntegrationTestStatus))
		end if

		if strDeveloperTestNotes <> "" then
			strDeveloperTestNotes = " - " & strDeveloperTestNotes
		end if
		if not blnRequiresDeveloperSignoff then
			strDeveloperTestStatusName = "N/A"
		else
			strDeveloperTestStatusName = teststatusarray(clng(strDeveloperTestStatus))
		end if

		if strODMTestNotes <> "" then
			strODMTestNotes = " - " & strODMTestNotes
		end if
		if not blnRequiresODMSignoff then
			strODMtestStatus = "N/A"
		else
			strODMtestStatus = teststatusarray(clng(strODMTestStatus))
		end if

		if strWWANTestNotes <> "" then
			strWWANTestNotes = " - " & strWWANTestNotes
		end if
		if not blnRequiresWWANSignoff then
			strWWANtestStatus = "N/A"
		else
			strWWANtestStatus = teststatusarray(clng(strWWANtestStatus))
		end if

		if trim(strRiskRelease) = "1" then
			strRiskRelease = "checked"
		else
			strRiskRelease = ""
		end if
		
		strStatusSelected = ""
		strCommentsRequired = ""
		rs.Open "spListTestStatus",cn,adOpenForwardOnly
		do while not rs.EOF
			strStatusText = rs("Status") & ""
			if trim(rs("ID")) = trim(strStatus) then
			    strStatusID = trim(rs("ID"))
				strStatusList = strStatusList & "<option selected value=""" & rs("ID") & """>" & strStatusText & "</option>"
				strStatusSelected = strStatusText
			elseif (not blnservicepm) then 'or rs("ID") = 18 then	
				strStatusList = strStatusList & "<option value=""" & rs("ID") & """>" & strStatusText & "</option>"
			end if

			if trim(rs("ID")) = "5" and not blnservicepm then
    			if trim(rs("ID")) = trim(strStatus) and strRiskRelease = "checked" then
	    			strStatusList = strStatusList & "<option selected value=""" & rs("ID") & """>Risk Release</option>"
		    		strStatusSelected = "Risk Release"
			    else	
				    strStatusList = strStatusList & "<option value=""" & rs("ID") & """>Risk Release</option>"
			    end if
			end if

			if rs("CommentsRequired") then
				strCommentsRequired = strCommentsRequired & "," & trim(rs("ID"))
			end if

			rs.movenext
		loop
		rs.Close
		strCommentsRequired = strCommentsRequired & ","

		strDevEmail = ""
		rs.open "spGetDeliverableDeveloper " & clng(request("VersionID")),cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			if trim(rs("DeveloperEmail") & "") <> "" then
				strDevEmail = strDevEmail & ";" & rs("DeveloperEmail")
			end if
			if trim(rs("DevManagerEmail") & "") <> "" then
				strDevEmail = strDevEmail & ";" & rs("DevManagerEmail")
			end if
		end if
		rs.Close

		if trim(strDevEmail) = "" then
			strDevEmail = "max.yu@hp.com"
		else
			strDevEmail = mid(strDevEmail,2)
		end if

		strPMs = ";" & strDevEmail
		
		rs.Open "spListHardwarePMs4Version " & clng(request("VersionID")),cn,adOpenStatic
		do while not rs.EOF
			strPMs = strPMs & ";" & rs("Email")
			rs.MoveNext
		loop
		rs.Close
		if trim(strPMs) = "" then
			strPMs = "max.yu@hp.com"
		else
			strPMs = mid(strPMs,2)
		end if
	
	end if

	if 	blnFound then
    %>


    <font face="verdana" size="2"><b>
    <label ID="lblTitle"><%=strVendor%>&nbsp;<%=strDeliverable%> on <%=strProduct%></label></b></font>
    <br />
    <span style="font-family: Verdana; font-size: 9pt;"><%=strReleaseLink %></span>
    <% 
	dim strVersion
	
	strVersion = ""
	if trim(strHW) <> "&nbsp;" then
		strVersion =  strHW
	end if
	if trim(strFW) <> "&nbsp;" then
		strVersion = strVersion & "," & strFW
	end if
	if trim(strRev) <> "&nbsp;" then
		strVersion = strVersion & "," & strRev
	end if
	
	strQCompleteSubject	= strVendor & " " & strDeliverable & " [" & strVersion & "] set to QComplete on " & strProduct 
	strQCompleteBody = "<font size=2 color=black face=Verdana><b>" & strQCompleteSubject & "</b></font><BR><BR>"
	strQCompleteBody = strQCompleteBody & "<font size=2 color=black face=Verdana>"

	strFailSubject= strVendor & " " & strDeliverable & " [" & strVersion & "] set to ##StatusText.## on " & strProduct 
	strFailBody = "<font size=2 color=black face=Verdana><b>" & strFailSubject & "</b></font>##ExtraNote.##<BR><BR>"
	strFailBody = strFailBody & "<font size=2 color=black face=Verdana>"
	
	strRows = "<TR>"
	if trim(request("VersionID")) <> "" then
		strRows = strRows & "<TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & request("VersionID") & """>" & request("VersionID") & "</a></TD>"
	else
		strRows = strRows & "<TD>&nbsp;</TD>"
	end if
	strRows = strRows & "<TD>" & strProduct & "</TD>"
	

	strRows = strRows & "<TD>" & strVendor & "</TD>"
	strRows = strRows & "<TD>" & strDeliverable & "</TD>"
	strRows = strRows & "<TD>" & strVersion & "</TD>"

	'if trim(strModel) <> "&nbsp;" then
		strRows = strRows & "<TD>" & strModel & "</TD>"
	'end if
	'if trim(strPartNumber) <> "" then
		strRows = strRows & "<TD>" & strPartNumber & "</TD>"
	'end if
		strRows = strRows & "<TD>" & strDevCell & "</TD>"
		strRows = strRows & "<TD>" & strMITCell & "</TD>"
		strRows = strRows & "<TD>" & strODMCell & "</TD>"
		strRows = strRows & "<TD>" & strWWANCell & "</TD>"
		strRows = strRows & "<TD>" & strTTSCell & "</TD>"
	strRows = strRows & "<TD>##ShowComments.##</TD>"

	strFailBody = strFailBody & "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>DEV</b></TD><TD><b>MIT</b></TD><TD><b>ODM</b></TD><TD><b>WWAN</b></TD><TD><b>TTS</b></TD><TD><b>Comments</b></TD></tr>" & strRows & "</table></font>"
	strQCompleteBody = strQCompleteBody & "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>DEV</b></TD><TD><b>MIT</b></TD><TD><b>ODM</b></TD><TD><b>WWAN</b></TD><TD><b>TTS</b></TD><TD><b>Comments</b></TD></tr>" & strRows & "</table></font>"

	strFailSubject = "Hardware Failure Notification"': " & strFailSubject
	strQCompleteSubject = "Hardware QComplete Notification"': " & strQCompleteSubject


	'Restrict body
	strRestrictBody = ""
	strRestrictBody = strRestrictBody & "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & request("VersionID") & """>" & request("VersionID") & "</a></TD>"
	strRestrictBody = strRestrictBody & "<TD>" & strDeliverable & "</TD>"
	strRestrictBody = strRestrictBody & "<TD nowrap>" & strVersion & "&nbsp;</TD>"
	strRestrictBody = strRestrictBody & "<TD>" & strModel & "&nbsp;</TD>"
	strRestrictBody = strRestrictBody & "<TD nowrap>" & strPartnumber & "&nbsp;</TD>"
	strRestrictBody = strRestrictBody & "</TR>"


    %>

    <form id="frmStatus" method="post" action="QualStatusSavePulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">

        <table id="tabGeneral" width="100%" bgcolor="cornsilk" border="1" cellspacing="0" cellpadding="2" bordercolor="tan">
            <tr>
                <td valign="top" width="60" nowrap><b>HW&nbsp;Ver:</b>&nbsp;</td>
                <td width="40%"><%=strHW%></td>
                <td valign="top" width="120" nowrap><b>Part&nbsp;Number:</b>&nbsp;</td>
                <td>
                    <%=strPartNumber%>&nbsp;
                </td>
            </tr>
            <tr>
                <td valign="top" width="60" nowrap><b>FW&nbsp;Ver:</b>&nbsp;</td>
                <td><%=strFW%></td>
                <td valign="top" width="120" nowrap><b>Model:</b>&nbsp;</td>
                <td width="100%">
                    <%=strModel%>
                </td>
            </tr>
            <tr>
                <td valign="top" width="60" nowrap><b>Rev:</b>&nbsp;</td>
                <td colspan="3"><%=strRev%></td>
            </tr>
            <tr>
                <td valign="top" width="120" nowrap><b>Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                <td colspan="3">
                    <select style="width: 160" id="cboStatus" name="cboStatus" language="javascript" onchange="cboStatus_onchange();">
                        <%if not blnservicepm then %>
                        <option value="0">Not Used Now</option>
                        <%end if %>
                        <%=strStatusList%>
                    </select>
                    <%
				if trim(strStatus)="0" or trim(strStatus)="" then
					strSupportDisplay = ""
				else
					strSupportDisplay = "none"
				end if
                    %>
                    <span id="SupportRow" style="display: <%=strSupportDisplay%>">
                        <input type="checkbox" id="chkDelete" disabled name="chkDelete">Completely remove support</span>


                    <%
				if trim(strStatus)="3" or strStatusSelected = "OOC" or strStatusSelected = "FCS" then
					strConfidenceDisplay = ""
				else
					strConfidenceDisplay = "none"
				end if

				strRiskReleaseDisplay = "none"

                    %>
                    <span id="RiskReleaseRow" style="display: <%=strRiskReleaseDisplay%>">
                        <input type="checkbox" <%=strRiskRelease%> id="chkRiskRelease" name="chkRiskRelease">Risk Release</span>
                    <span id="ConfidenceRow" style="display: <%=strConfidenceDisplay%>">&nbsp;<font size="2" face="verdana"><b>Confidence:</b>&nbsp;
			<SELECT id=cboConfidence name=cboConfidence>
				<%if strConfidence = "1" then%>
					<OPTION selected value=1>High</OPTION>
				<%else%>
					<OPTION value=1>High</OPTION>
				<%end if%>
				<%if strConfidence = "2" then%>
				<OPTION selected value=2>Med</OPTION>
				<%else%>
				<OPTION value=2>Med</OPTION>
				<%end if%>
				<%if strConfidence = "3" then%>
				<OPTION selected value=3>Low</OPTION>
				<%else%>
				<OPTION value=3>Low</OPTION>
				<%end if%>
			</SELECT>			
                    </span>
                </td>
            </tr>

            <%if trim(strStatus)="3" then %>
            <tr id="DateRow">
                <%else%>
            <tr id="DateRow" style="display: none">
                <%end if%>
                <td valign="top" width="120" nowrap><b>Date:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                <td colspan="3">
                    <input type="text" id="txtTestDate" name="txtTestDate" value="<%=strDate%>">&nbsp;<a href="javascript: cmdDate_onclick()"><img id="picTarget" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21"></a>
                </td>
            </tr>
            <%if blnRequiresWWANSignoff or blnRequiresODMSignoff or blnRequiresMITSignoff or blnRequiresDeveloperSignoff then%>
            <tr>
                <%else%>
            <tr style="display: none">
                <%end if%>
                <td nowrap valign="top" width="120" nowrap><b>Testing&nbsp;Summary:</b>&nbsp;</td>
                <td colspan="3" style="display: ">
                    <%if blnRequiresDeveloperSignoff then%>
                    <b>Dev</b>: <%=strDeveloperTestStatusName%>&nbsp;
		<%end if%>
                    <%if blnRequiresMITSignoff then%>
                    <b>MIT</b>: <%=strIntegrationTestStatus%>&nbsp;
		<%end if%>
                    <%if blnRequiresODMSignoff then%>
                    <b>ODM</b>: <%=strODMTestStatus%>&nbsp;
		<%end if%>
                    <%if blnRequiresWWANSignoff then%>
                    <b>COMM</b>: <%=strWWANtestStatus%>&nbsp;
		<%end if%>
                    <%if blnRequiresWWANSignoff then%>
                    <b>TTS</b>: <%=strTTS%>
                    <%end if%>
                </td>
                <td style="display: none">
                    <input <%=strIntegrationTestStatusCheck%> type="checkbox" id="chkIntegrationTesting" name="chkIntegrationTesting" value="1">HP&nbsp;&nbsp;
		<input <%=strODMTestStatusCheck%> type="checkbox" id="chkODMTesting" name="chkODMTesting" value="1">ODM&nbsp;&nbsp;
		<input <%=strDeveloperTestStatusCheck%> type="checkbox" id="chkDeveloperTesting" name="chkDeveloperTesting" value="1">Developer
                </td>
            </tr>
            <tr>
                <td valign="top" width="120" nowrap><b>HFCN/DCR:</b>&nbsp;</td>
                <td colspan="3">
                    <select style="width: 160" id="cboWhy" name="cboWhy" language="javascript" onchange="cboWhy_onchange();">
                        <%if strDCR = "0" then%>
                        <option selected value="0"></option>
                        <%else%>
                        <option value="0"></option>
                        <%end if%>
                        <%if strDCR = "1" then%>
                        <option selected value="1">POR</option>
                        <%else%>
                        <option value="1">POR</option>
                        <%end if%>
                        <%if strDCR = "2" then%>
                        <option selected value="2">HFCN</option>
                        <%else%>
                        <option value="2">HFCN</option>
                        <%end if%>
                        <%if clng(strDCR) > 2 then%>
                        <option selected value="3">DCR</option>
                        <%else%>
                        <option value="3">DCR</option>
                        <%end if%>
                    </select>
                </td>
            </tr>
            <%if clng(strDCR) < 3 then%>
            <tr id="DCRRow" style="display: none">
                <%else%>
            <tr id="DCRRow">
                <%end if%>
                <td valign="top" width="120" nowrap><b>DCR&nbsp;Number:</b>&nbsp;<font id="DCRReq" color="red" size="1">*</font>&nbsp;</td>
                <td colspan="3">
                    <select style="width: 100%" id="cboDCR" name="cboDCR" language="javascript" onkeydown="return combo_onkeydown()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeypress="return combo_onkeypress()">
                        <option selected value="0"></option>
                        <%
					rs.Open "spListApprovedDCRs " & clng(request("ProdID")) & ",1" ,cn,adOpenForwardOnly					
					do while not rs.EOF
						if strDCR = trim(rs("ID")) then
							Response.Write "<option selected value=""" & rs("ID") & """>" & rs("ID") & " [" & rs("Status") & "] - " & server.HTMLEncode(rs("Summary")) & "</option>"					
						elseif lcase(trim(rs("Status"))) <> "disapproved" then
							Response.Write "<option value=""" & rs("ID") & """>" & rs("ID") & " [" & rs("Status") & "] - " & server.HTMLEncode(rs("Summary")) & "</option>"					
						end if
						rs.MoveNext
					loop
					rs.Close
                        %>
                    </select>
                </td>
            </tr>
            <tr>
                <td valign="top" width="120" nowrap><b>Restrictions:</b></td>
                <td colspan="3">
                    <input <%=strSupplyRestricted%> type="checkbox" id="chkSupplyRestricted" name="chkSupplyRestricted" value="1">
                    Supply
			<input <%=strConfigurationRestricted%> type="checkbox" id="chkConfigurationRestricted" name="chkConfigurationRestricted" value="1">
                    Configuration
			
                </td>
            </tr>

            <tr>
                <td valign="top" width="120" nowrap><b>Comments:</b>
                    <%if instr(strCommentsRequired,"," & trim(strStatus) & ",") > 0 then %>
                    <font color="red" size="1" id="CommentStar">*</font>&nbsp;
		<%else%>
                    <font style="display: none" color="red" size="1" id="CommentStar">*</font>&nbsp;
		<%end if%>
                </td>
                <td colspan="7">
                    <textarea rows="3" style="width: 100%" id="txtComments" name="txtComments"><%=strComments%></textarea>
                </td>
            </tr>
            <tr>
                <td valign="top" width="120" nowrap><b>Subassemblies:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                <td colspan="3">
                    <%
			rs.Open "spCountBridgesOnOtherVersionsRelease " & clng(request("ProdID")) & "," & clng(request("VersionID")),cn,adOpenForwardOnly
			if not(rs.EOF and rs.BOF) then
				if trim(rs("bridgeCount") & "") <> "" then
					if clng(rs("BridgeCount")) =1 then
						Response.Write "<DIV style=""MARGIN-BOTTOM:5px""><font size=2 face=verdana><b>Note: <a target=""_blank"" href=""../HardwareMatrix.asp?ReportFormat=2&lstProducts=" & clng(request("ProdID")) & "&lstRoot=" & rs("BridgedIDs") & "&lstVendor=" & clng(strVendorID) & """>One other version</a> of this deliverable are bridged on this product.</b></font></DIV>"
					elseif clng(rs("BridgeCount")) > 1 then
						Response.Write "<DIV style=""MARGIN-BOTTOM:5px""><font size=2 face=verdana><b>Note: <a target=""_blank"" href=""../HardwareMatrix.asp?ReportFormat=2&lstProducts=" & clng(request("ProdID")) & "&lstRoot=" & rs("BridgedIDs") & "&lstVendor=" & clng(strVendorID) & """>" & rs("BridgeCount") & " other versions</a> of this deliverable are bridged on this product.</b></font></div>"
					end if
				end if
			end if
			rs.Close
                    %>
                    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll; border-left: steelblue 1px solid; width: 100%; border-bottom: steelblue 1px solid; height: 143px; background-color: white" id="DIV1">
                        <table id="TableSub" width="100%">
                            <thead bgcolor="LightSteelBlue">
                                <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset">&nbsp</td>
                                <td onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableSub', 1 ,0,1);" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset">&nbsp;Number&nbsp;</td>
                                <td onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableSub', 2 ,0,1);" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset; width=100%">&nbsp;Name&nbsp;</td>
                            </thead>
                            <% 
		
			dim SubAssemblyCount
			dim MissingSubAssemblyList
		
			rs.open "spGetSubassemblyCountRelease "& clng(request("ProdID")) & "," & clng(request("VersionID")) & "," & ProductDeliverableReleaseID,cn,adOpenForwardOnly
			if not (rs.EOF and rs.BOF)then
				SubAssemblyCount = rs("Number")
			else
				SubAssemblyCount = 0
			end if
			rs.Close
			
			rs.Open "spListSubassembliesAllAppropriateRelease " & clng(request("ProdID")) & "," & clng(request("VersionID")) & "," & ProductDeliverableReleaseID,cn,adOpenForwardOnly
			strSubsLoaded = ""
			strMissingSubAssembly=""
			do while not rs.eof
				if rs("isRealRoot") then
					strSubBGColor = "Lavender"
				elseif rs("Active") <> 1 then
					strSubBGColor = "LavenderBlush"
				else
					strSubBGColor = "white"
				end if
                            %>
                            <tr bgcolor="<%=strSubBGColor%>">
                                <td nowrap>
                                    <% if rs("Selected") or (SubAssemblyCount=0 and rs("isRealRoot")) then%>
                                    <input id="lstSub" checked type="checkbox" name="lstSub" value="<%=rs("ID")%>">
                                    <%
						                    if rs("Selected") then 'Don't count the auto-selected real root as ones loaded
							                    strSubsLoaded = strSubsLoaded & ", " & rs("ID")
						                    end if
                                                        %>
                                                        <%else%>
                                                        <input id="lstSub" type="checkbox" name="lstSub" value="<%=rs("ID")%>">
                                                        <%end if%>
                                                    </td>
                                                    <td>&nbsp;<%
					                    if trim(rs("Subassembly") & "") = "" then
						                    Response.Write "TBD"
						                    strMissingSubAssembly = "<font color=red face=verdana size=1>Note: ""TBD"" subassemblies will not appear on the matrix until the number is entered.</font>"
					                    else						
						                    Response.Write rs("Subassembly") 
					                    end if
                                                    %></td>
                                                    <td nowrap>&nbsp;<%=rs("Name")%></td>
                                                </tr>
                                                <%
				                    rs.MoveNext
			                    loop
			                    rs.Close
			
			                    if trim(strSubsLoaded) <> "" then
				                    strSubsLoaded = mid(strSubsLoaded,3)
			                    end if
                            %>
                        </table>

                    </div>
                    <%=strMissingSubAssembly%>			
		
                </td>
            </tr>

            <%if blnRequiresDeveloperSignoff then%>
            <tr>
                <%else%>
            <tr style="display: none">
                <%end if%>
                <td valign="top" width="120" nowrap><b>Developer&nbsp;Status:</b>&nbsp;</td>
                <td colspan="3"><%=strDevStatus & strDeveloperTestNotes%>&nbsp;&nbsp;</td>
            </tr>
            <%if blnRequiresMITSignoff then%>
            <tr>
                <%else%>
            <tr style="display: none">
                <%end if%>
                <td valign="top" width="120" nowrap><b>MIT&nbsp;Status:</b>&nbsp;</td>
                <td colspan="3"><%=strIntegrationTestStatus & strIntegrationTestNotes%>&nbsp;&nbsp;</td>
            </tr>
            <%if blnRequiresODMSignoff then%>
            <tr>
                <%else%>
            <tr style="display: none">
                <%end if%>
                <td valign="top" width="120" nowrap><b>ODM&nbsp;Status:</b>&nbsp;</td>
                <td colspan="3"><%=strODMTestStatus & strODMTestNotes%>&nbsp;&nbsp;</td>
            </tr>
            <%if blnRequiresWWANSignoff then%>
            <tr>
                <%else%>
            <tr style="display: none">
                <%end if%>
                <td valign="top" width="120" nowrap><b>COMM Status:</b>&nbsp;</td>
                <td colspan="3"><%=strWWANTestStatus & strWWANTestNotes%>&nbsp;&nbsp;
                </td>
            </tr>

        </table>


        <input style="display: none" type="text" id="txtID" name="txtID" value="<%=strID%>">
        <input type="hidden" id="txtProdDelRelID" name="txtProdDelRelID" value="<%=ProductDeliverableReleaseID%>" />
        <input type="hidden" id="txtReleaseName" name="txtReleaseName" value="<%=ProductDeliverableReleaseName%>" />
        <input type="hidden" id="txtProdID" name="txtProdID" value="<%=request("ProdID")%>">
        <input type="hidden" id="txtProduct" name="txtProduct" value="<%=strProduct%>">
        <input type="hidden" id="txtVendor" name="txtVendor" value="<%=strVendor%>">
        <input type="hidden" id="txtCategory" name="txtCategory" value="<%=strCategory%>">
        <input type="hidden" id="txtRestrictBody" name="txtRestrictBody" value="<%=server.htmlencode(strRestrictBody)%>">
        <input type="hidden" id="txtDevCenter" name="txtDevCenter" value="<%=strDevCenter%>">
        <input type="hidden" id="txtVersionID" name="txtVersionID" value="<%=request("VersionID")%>">
        <input type="hidden" id="txtCurrentUserID" name="txtCurrentUserID" value="<%=CurrentUserID%>">
        <input type="hidden" id="txtStatusText" name="txtStatusText" value="">
        <input type="hidden" id="txtStatusLoaded" name="txtStatusLoaded" value="<%=trim(strStatus)%>">
        <input type="hidden" id="txtSubsLoaded" name="txtSubsLoaded" value="<%=strSubsLoaded%>">
        <input type="hidden" id="txtUserName" name="txtUserName" value="<%=CurrentDomain & "_" & CurrentUser%>">
        <input type="hidden" id="txtUserEmail" name="txtUserEmail" value="<%=CurrentUserEmail%>">
        <input type="hidden" id="txtUserID" name="txtUserID" value="<%=CurrentUserID%>">
        <input type="hidden" id="txtFailSubject" name="txtFailSubject" value="<%=strFailSubject%>">
        <input type="hidden" id="txtFailBody" name="txtFailBody" value="<%=server.HTMLEncode(strFailBody)%>">
        <input type="hidden" id="txtPMsEmail" name="txtPMsEmail" value="<%=strPMs%>">
        <input type="hidden" id="txtDevEmail" name="txtDevEmail" value="<%=strDevEmail%>">
        <input type="hidden" id="txtQCompleteSubject" name="txtQCompleteSubject" value="<%=strQCompleteSubject%>">
        <input type="hidden" id="txtQCompleteBody" name="txtQCompleteBody" value="<%=server.HTMLEncode(strQCompleteBody)%>">
        <input type="hidden" id="tagSupplyRestriction" name="tagSupplyRestriction" value="<%=strSupplyRestrictionID%>">
        <input type="hidden" id="tagConfigurationRestriction" name="tagConfigurationRestriction" value="<%=strConfigurationRestrictionID%>">
        <input type="hidden" id="txtPartnerID" name="txtPartnerID" value="<%=strPartnerID%>">
        <input type="hidden" id="txtMissingSubAssemblyList" name="txtMissingSubAssemblyList" value="<%=MissingSubAssemblyList%>">
        <input type="hidden" id="txtTestingComplete" name="txtTestingComplete" value="<%=blnTestingComplete%>">
        <input type="hidden" id="txtTestLeadsTo" name="txtTestLeadsTo" value="<%=strTestLeadsTo%>">
        <input type="hidden" id="txtPartNo" name="txtPartNo" value="<%=strPartNumber%>">
        <input type="hidden" id="txtInitialStatusID" name="txtInitialStatusID" value="<%=strStatusID%>">
        <input type="hidden" id="txtModelNo" name="txtModelNo" value="<%=strModel%>">
        <input type="hidden" id="txtDeliverableName" name="txtDeliverableName" value="<%=strDeliverable%>">
        <input type="hidden" id="txtRedirect" name="txtRedirect" value="" />
        <input type="hidden" id="txtKeepItOpen" name="txtKeepItOpen" value="false" />
        <input type="hidden" id="txtProductDeliverableID" name="txtProductDeliverableID" value="<%=ProductDeliverableID%>" />
        <input type="hidden" id="txtTodayPageSection" name="txtTodayPageSection" value="<%=Request("TodayPageSection")%>" />
        <input type="hidden" id="txtShowOnlyTargetedRelease" name="txtShowOnlyTargetedRelease" value="<%=Request("ShowOnlyTargetedRelease")%>" />
        <% if strRiskRelease = "checked" then%>
        <input type="hidden" id="tagRiskRelease" name="tagRiskRelease" value="on">
        <%end if%>
    </form>
    <%end if

	cn.Close
	set cn = nothing
	set rs = nothing


    %>
    <input type="hidden" id="txtCommentsRequired" name="txtCommentsRequired" value="<%=strCommentsRequired%>">
</body>
</HTML>
