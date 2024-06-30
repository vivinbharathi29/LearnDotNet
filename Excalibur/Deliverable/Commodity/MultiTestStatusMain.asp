<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

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


    function cmdDate_onclick(FieldID) {
        var strID;

        if (FieldID == 1) {
            strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", $("#txtTestDate").val(), "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined")
                $("#txtTestDate").val(strID);
        }
        else if (FieldID == 2) {
            strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", $("#txtPilotDate").val(), "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined")
                $("#txtPilotDate").val(strID);
        }
        else {
            strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", $("#txtAccessoryDate").val(), "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined")
                $("#txtAccessoryDate").val(strID);
        }
    }

    function cboStatus_onchange() {

        var QualStatusVal = $("#cboStatus").val();
        var strStatus = $("#cboStatus option:selected").text();
        $("#NewValue").val(QualStatusVal);


        if (QualStatusVal == "3")
            $("#DateRow").show();
        else
            $("#DateRow").hide();

        if (strStatus == "FCS" || strStatus == "OOC" || QualStatusVal == "3")
            $("#ConfidenceRow").show();
        else
            $("#ConfidenceRow").hide();

        if (QualStatusVal == "0") {
            //	SupportRow.style.display="";
        }
        else {
            $("#chkDelete").prop('checked', false);
            //	SupportRow.style.display="none";
        }
        if (strStatus == "QComplete")
            $("#RiskReleaseRow").show();
        else
            $("#RiskReleaseRow").hide();

    }

    function cboPilotStatus_onchange() {

        var PilotStatusVal = $("#cboPilotStatus").val();
        var strRequired = $("#txtPilotDateRequired").val().indexOf(PilotStatusVal);
        var strShow = $("#txtPilotDateShow").val().indexOf(PilotStatusVal);


        $("#NewPilotValue").val(PilotStatusVal);


        if (strShow != -1 && $("#NewPilotValue").val() != "") {
            $("#PilotDateRow").show();
            if (strRequired != -1)
                $("#PilotDateStar").show();
            else
                $("#PilotDateStar").hide();
        }
        else {
            $("#PilotDateRow").hide();
            $("#txtPilotDate").val("");
            $("#PilotDateStar").hide();
        }


    }

    function cboAccessoryStatus_onchange() {

        var AccessoryStatusVal = $("#cboAccessoryStatus").val();
        var strRequired = $("#txtAccessoryDateRequired").val().indexOf(AccessoryStatusVal);
        var strShow = $("#txtAccessoryDateShow").val().indexOf(AccessoryStatusVal);

        $("#NewAccessoryValue").val(AccessoryStatusVal);


        if (strShow != -1 && $("#NewAccessoryValue").val() != "") {
            $("#AccessoryDateRow").show();
            if (strRequired != -1)
                $("#AccessoryDateStar").show();
            else
                $("#AccessoryDateStar").hide();
        }
        else {
            $("#AccessoryDateRow").hide();
            $("#txtAccessoryDate").val("");
            $("#AccessoryDateStar").hide();
        }
    }


    //-->
</SCRIPT>
</HEAD>
<STYLE>
	TD
	{
	VERTICAL-ALIGN: top
	}
	
</STYLE>
<BODY bgcolor="ivory"  LANGUAGE=javascript>
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<%

	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
	
	dim cn
	dim rs
	dim p
	dim cm
	dim CurrentUser
	dim CurrentUserPartnerID
	dim strConfidenceDisplay
	dim strDeliverableList
	dim strSQl
	dim strTestStatus
	dim strPilotStatus
	dim strAccessoryStatus
	dim blnAllSameStatus
	dim blnAllSamePilotStatus
	dim blnAllSameAccessoryStatus
	dim strRootName
	dim strVersion
	dim strVendor
	dim strPart
	dim strModel
	dim	blnCommodityPM
    dim blnOdmHwPM
	dim blnCommPM 
	dim blnPlatformDevelopmentPM
	dim blnProcessorPM
	dim blnVideoMemoryPM  
	dim blnGraphicsControllerPM
	dim blnHardwarePM
	dim blnServicePM
	dim blnSuperUser
	dim blnPilotEngineer
	dim blnAccessoryPM
	dim strNewTestStatus
	dim strNewPilotStatus
	dim strNewAccessoryStatus
    dim blnEngineeringCoordinator
	
	
	blnCommodityPM = 0
    blnOdmHwPM = 0
	blnCommPM  = 0
	blnPlatformDevelopmentPM = 0
	blnProcessorPM = 0
	blnGraphicsControllerPM = 0
	blnVideoMemoryPM   = 0
	blnHardwarePM = 0
	blnServicePM = 0
	blnSuperUser = 0
	blnPilotEngineer = 0
	blnAccessoryPM = 0
    blnEngineeringCoordinator = 0
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
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
		CurrentUserPartnerID = rs("PartnerID")
        
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		'if Currentuserid = 31 or CurrentUserID = 8 or CurrentUserID = 1396 then
			'blnCommodityPM = True
			'blnPilotEngineer = True
			'blnAccessoryPM = true
		'else
			blnCommodityPM = rs("CommodityPM")
			blnPilotEngineer = rs("SCFactoryEngineer") 
			blnAccessoryPM = rs("AccessoryPM")
            blnEngineeringCoordinator = rs("EngCoordinator")
		'end if
		blnServicePM = rs("ServicePM")
	else
		CurrentUserID = 0
		blnCommodityPM = 0
		blnPilotEngineer = 0
		blnAccessoryPM = 0
        blnEngineeringCoordinator = 0
		blnServicePM = false
	end if
	rs.Close

	rs.open "spGetHardwareTeamAccessList " & CurrentUserID & "," & clng(strID),cn,adOpenStatic
	do while not rs.EOF
        if rs("HWTeam") = "ProgramCoordinator" or blnEngineeringCoordinator > 0 then
				blnPlatformDevelopmentPM = true
		elseif rs("HWTeam") = "PlatformDevelopment" and rs("Products") > 0 then
			blnPlatformDevelopmentPM = true
		elseif rs("HWTeam") = "Processor" and rs("Products") > 0 then
			blnProcessorPM = true
		elseif rs("HWTeam") = "Comm" and rs("Products") > 0 then
			blnCommPM = true
		elseif rs("HWTeam") = "VideoMemory" and rs("Products") > 0 then
			blnVideoMemoryPM = true
		elseif rs("HWTeam") = "GraphicsController" and rs("Products") > 0 then
			blnGraphicsControllerPM = true
		elseif rs("HWTeam") = "SuperUser" and rs("Products") > 0 then
			blnSuperUser = true
		elseif rs("HWTeam") = "Commodity" and rs("Products") > 0 then
			blnCommodityPM = true
		end if
		rs.MoveNext			
	loop
	rs.Close
	if blnservicepm and (blnPlatformDevelopmentPM or blnProcessorPM or blnCommPM or blnVideoMemoryPM or blnGraphicsControllerPM or blnSuperUser) then
	    blnServicepm = false
	end if

	strProduct = ""
	if trim(request("ProdID")) <> ""  and trim(request("ProdID")) <> "0" then

        ''' By Design: ODM cannot use Batched Update from TodayPage (Section: Components Awaiting Final Approval)
	    rs.Open "select ODMHWPMID from ProductVersion where id=" & clng(request("ProdID")),cn,adOpenForwardOnly
	    if not (rs.EOF and rs.BOF) then
		    if trim(rs("ODMHWPMID") & "") = trim(CurrentUserID) then
                blnOdmHwPM = true
            end if
	    end if
	    rs.Close

		rs.Open "spGetProductVersionName " & clng(request("ProdID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strProduct = rs("Name") & ""
		end if
		rs.Close

	end if
	
	blnCommodityPM= blnCommodityPM or blnPlatformDevelopmentPM or blnServicePM or blnProcessorPM or blnCommPM or blnVideoMemoryPM or blnGraphicsControllerPM or blnSuperUser or blnOdmHwPM

	dim blnEnoughInfo
	dim blnAddLinkStatus
	dim blnShowProducts
	dim strEOL
	dim strEOLBGColor
	dim strPilotDateShow
	dim strAccessoryDateShow
	dim strPilotDateRequired
	dim strAccessoryDateRequired
	
	blnEnoughInfo = false
	blnShowProducts = false
	blnAddLinkStatus = true
	
	if trim(Request("RootID")) <> "" and trim(Request("RootID")) <> "0" then
		blnEnoughInfo = true
        strSQL = "Select v.tts,pd.targetnotes, pd.riskrelease, pd.accessorynotes,pd.pilotnotes, v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, pd.accessorystatusid,pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate, t.id as StatusID, t.status, pd.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                 "from dbo.TestStatus AS t WITH (NOLOCK) RIGHT OUTER JOIN " & _
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID ON v.ID = pd.DeliverableVersionID ON t.ID = pd.TestStatusID " & _
                "WHERE        (r.ID = " & clng(Request("RootID")) & ") AND (pd.ProductVersionID = " & clng(request("ProdID")) & ") " & _
                " order by r.name, vd.name, v.id desc;"


	elseif trim(Request("ProdID")) <> "" and trim(Request("ProdID")) <> "0" then
		blnEnoughInfo = true
		strSQl = "Select v.tts, pd.targetnotes, pd.riskrelease, pd.accessorynotes,pd.pilotnotes,  v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, pt.id as PilotStatusID,pd.accessorystatusid,pd.accessorydate, pt.Name as PilotStatus, pd.PilotDate, t.id as StatusID, t.status, pd.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                "FROM dbo.TestStatus AS t WITH (NOLOCK) RIGHT OUTER JOIN " & _
                "dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                "dbo.Product_Deliverable AS pd WITH (NOLOCK) INNER JOIN " & _
                "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID ON v.ID = pd.DeliverableVersionID ON t.ID = pd.TestStatusID " & _
                "WHERE v.ID IN (" & scrubsql(request("VersionList")) & ") " & _
                "AND (pd.ProductVersionID = " &  clng(request("ProdID")) & ") " & _
                "ORDER BY r.Name, vendor, VersionID DESC;"

	elseif trim(Request("VersionList")) <> "" and trim(Request("VersionList")) <> "0" and instr(Request("VersionList"),",")=0 then
		blnEnoughInfo = true
		blnShowProducts = true
		strSQl = "Select v.tts, pd.targetnotes, pd.riskrelease, pd.accessorynotes, pd.pilotnotes,  v.active as EOL, at.name as AccessoryStatus,  v.EndOfLifeDate as EOLDate,  c.commodity, pv.DotsName as Product,pd.accessorystatusid,pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate,  t.id as StatusID, t.status, pd.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                 "FROM dbo.DeliverableRoot AS r WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID LEFT OUTER JOIN " & _
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _
                  "WHERE v.ID = " & clng(request("VersionList")) & " " 
				if CurrentUserPartnerID <> 1 then
					strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID & " " 
				end if
		strSQl = strSQL &  " order by  pv.DotsName;"
	elseif trim(Request("VersionList")) <> "" and trim(Request("VersionList")) <> "0" then
		blnEnoughInfo = true
		blnShowProducts = true
		strSQl = "Select v.tts,pd.targetnotes, pd.riskrelease, pd.accessorynotes,pd.pilotnotes, v.active as EOL, at.name as AccessoryStatus,  v.EndOfLifeDate as EOLDate, c.commodity,  pv.DotsName as Product,pd.accessorystatusid,pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate,  t.id as StatusID, t.status, pd.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                 "FROM dbo.AccessoryStatus AS at WITH (NOLOCK) RIGHT OUTER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID ON pd.DeliverableVersionID = v.ID ON at.ID = pd.AccessoryStatusID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID  " & _
                "WHERE pd.ID IN (" & scrubsql(request("VersionList")) & ") "

				if CurrentUserPartnerID <> 1 then
					strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID & " " 
				end if

		strSQl = strSQL &  " order by r.name, vd.name, v.id desc;"

	end if
	
'	Response.Write strSQL
    'response.flush
	if blnEnoughInfo then
		strDeliverableList = ""
		rs.open strSQl, cn,adOpenForwardOnly
		blnAllSameStatus = true
		blnAllSamePilotStatus = true
		blnAllSameAccessoryStatus = true
		strTestStatus = ""
		TestStatusID = -1
		PilotStatusID = -1
		AccessoryStatusID = -1
		
		dim UpdatableVersionCount
		UpdatableVersionCount=0
		do while not rs.EOF
			strRootName = rs("Name")
			strVersion = rs("Version") & ""
			if trim(rs("Revision") & "") <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if	
			if trim(rs("Pass") & "") <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			
			if isnull(rs("Commodity")) then
					blnAddLinkStatus = false
			elseif not rs("Commodity") then
					blnAddLinkStatus = false
			end if
				
			strVendor = rs("Vendor") & ""
			strModel = rs("ModelNumber") & ""
			strPart = rs("PartNumber") & ""
			
			if TestStatusID <> -1 and TestStatusID <> rs("StatusID") & "" then
				blnAllSameStatus = false
			end if	
	
			if PilotStatusID <> -1 and PilotStatusID <> rs("PilotStatusID") & "" then
				blnAllSamePilotStatus = false
			end if	

			if AccessoryStatusID <> -1 and AccessoryStatusID <> rs("AccessoryStatusID") & "" then
				blnAllSameAccessoryStatus = false
			end if	

			TestStatusID = rs("StatusID") & ""
			PilotStatusID = rs("PilotStatusID") & ""
			AccessoryStatusID = rs("AccessoryStatusID")			
			if isnull(rs("Status")) then
				strTestStatus = "Not Used"
			elseif rs("Status") & "" = "Date" then
				strTestStatus = rs("TestDate")
			elseif rs("Status") & "" = "QComplete" and rs("RiskRelease") then
				strTestStatus = "Risk Release"
			else
				strTestStatus = rs("Status")
			end if
			if rs("PilotStatus") = "P_Scheduled" then
				strPilotStatus = rs("PilotDate") & ""
			else
				strPilotStatus = rs("PilotStatus")
			end if
			if rs("AccessoryStatus") = "Scheduled" then
				strAccessoryStatus = rs("AccessoryDate") & ""
			else
				strAccessoryStatus = rs("AccessoryStatus")
			end if
			
			dim blnQCompleteOK
			blnQCompleteOK = true
			
			if rs("WWANProduct") and rs("requiresWWANtestfinalapproval") and trim(rs("WWANTestStatus")) <> "1"  and trim(rs("WWANTestStatus")) <> "5" then
				blnQCompleteOK = false
			end if
			if rs("WWANProduct") and rs("requiresWWANtestfinalapproval") and lcase(trim(rs("tts"))) = "pending" then
				blnQCompleteOK = false
			end if
			if rs("requiresodmtestfinalapproval") and trim(rs("ODMTestStatus")) <> "1"  and trim(rs("ODMTestStatus")) <> "5" then
				blnQCompleteOK = false
			end if
			if rs("requiresMITtestfinalapproval")  and trim(rs("IntegrationTestStatus")) <> "1"  and trim(rs("IntegrationTestStatus")) <> "5" then
				blnQCompleteOK = false
			end if
			if rs("requiresdeveloperfinalapproval") and trim(rs("DeveloperTestStatus")) <> "1"  and trim(rs("DeveloperTestStatus")) <> "5" then
				blnQCompleteOK = false
			end if
			
			
			if instr(rs("Location") & "","Core Team") > 0 then
				strDeliverableList = strDeliverableList & "<TR><TD>&nbsp;</TD>"
				strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("VersionID") & "&nbsp;</td>"
				strDeliverableList = strDeliverableList & "<TD>Core Team</td>"
				strDeliverableList = strDeliverableList & "<TD>Core Team</td>"
			elseif instr(rs("Location") & "","Development") > 0 then
				strDeliverableList = strDeliverableList & "<TR><TD>&nbsp;</TD>"
				strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("VersionID") & "&nbsp;</td>"
				strDeliverableList = strDeliverableList & "<TD>Development</td>"
				strDeliverableList = strDeliverableList & "<TD>Development</td>"
				strDeliverableList = strDeliverableList & "<TD>Development</td>"
			else
				UpdatableVersionCount = UpdatableVersionCount + 1
				if Request("RootID") <> "" then
					strDeliverableList = strDeliverableList & "<TR><TD><INPUT type=""checkbox"" QCompleteOK=""" & blnQCompleteOK & """ style=""Width:16;Height:16"" TestStatus=" & TestStatusID & " id=txtMultiID name=txtMultiID value=""" & rs("ID") & """></TD>" '" & replace(replace(blnQCompleteOK,"False",""),"True",".") & "
				else
				    strDeliverableList = strDeliverableList & "<TR><TD><INPUT type=""checkbox"" QCompleteOK=""" & blnQCompleteOK & """ style=""Width:16;Height:16"" TestStatus=" & TestStatusID & " checked id=txtMultiID name=txtMultiID value=""" & rs("ID") & """></TD>"
				end if
				if blnShowProducts then
					strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Product") & "&nbsp;</td>"
				else
					strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("VersionID") & "&nbsp;</td>"
				end if
				strDeliverableList = strDeliverableList & "<TD nowrap>" & strTestStatus & "&nbsp;</td>"
				strDeliverableList = strDeliverableList & "<TD nowrap>" & strPilotStatus & "&nbsp;</td>"
				strDeliverableList = strDeliverableList & "<TD nowrap>" & strAccessoryStatus & "&nbsp;</td>"
			end if
				strTestComments=""
				if trim(rs("TargetNotes") & "") <> "" then
					strTestComments = "<b>QUAL:</b> " & rs("TargetNotes") & ""
				end if
				if trim(rs("pilotnotes") & "") <> "" then
					if strTestComments <> "" then
						strTestComments = strTestComments & "<BR>"
					end if
					strTestComments = strTestComments & "<b>PILOT:</b> " & rs("PilotNotes")
				end if
				if trim(rs("accessorynotes") & "") <> "" then
					if strTestComments <> "" then
						strTestComments = strTestComments & "<BR>"
					end if
					strTestComments = strTestComments & "<b>ACC:</b> " & rs("AccessoryNotes")
				end if
			if blnShowProducts then
				strDeliverableList = strDeliverableList & "<td nowrap>" & strTestComments & "&nbsp;</td></tr>"
			else

				strEOl = "&nbsp;"
				strEOLBGColor = ""
				if not rs("EOL") then
					strEOL = "Unavailable"
					strEOLBGColor = "red"
				elseif isdate(rs("EOLDate")) then
					strEOL = rs("EOLDATE") & "&nbsp;"
					if datediff("d",rs("EOLDate"),now) > 0 then
						strEOLBGColor = "yellow"
					end if
				end if
				strDeliverableList = strDeliverableList & "<td bgcolor=" & strEOLBGColor & ">" & strEOL & "</td>"


				if Request("RootID") = "" then
					strDeliverableList = strDeliverableList & "<TD>" & rs("Name") & "</td>"
				end if
				strDeliverableList = strDeliverableList & "<td>" & rs("Vendor") & "</td><td nowrap>" & strVersion & "</td><td nowrap>" & rs("ModelNumber") & "&nbsp;</td><td nowrap>" & rs("PartNumber") & "&nbsp;</td>"
				strDeliverableList = strDeliverableList & "<td nowrap>" & strTestComments & "&nbsp;</td></tr>"
			end if
			rs.MoveNext
		loop
		rs.Close
	
	end if

	if strDeliverableList="" then
		if Request("RootID") <> "" and Request("RootID") <> "0" then
			Response.Write "There are no versions currently available supporting this product."
		elseif  Request("RootID") = "0" then
			Response.Write "There are no active products supporting this version."
		else
			Response.Write "Not enough information supplied to process your request."
		end if
	else
		
		strStatusText = ""
		strStatusList = ""
		if blnCommodityPM then
			rs.Open "spListTestStatus",cn,adOpenForwardOnly
			do while not rs.EOF
				strStatusText = rs("Status") & ""
	
				if trim(rs("ID")) = trim(TestStatusID) and blnAllSameStatus then
					strStatusList = strStatusList & "<option selected value=""" & rs("ID") & """>" & strStatusText & "</option>"
				elseif (not blnservicepm) or  trim(rs("ID")) = "18" then
					strStatusList = strStatusList & "<option value=""" & rs("ID") & """>" & strStatusText & "</option>"
				end if
				rs.movenext
			loop
			rs.Close
		end if
		
		strPilotStatusText = ""
		if blnAllSamePilotStatus then
			strPilotStatusList = ""
		else
			strPilotStatusList = "<Option value=""""></OPTION>"
		end if
		if blnPilotEngineer then
		
			strPilotDateShow = ""
			strPilotDateRequired = ""

		
			rs.Open "spListPilotStatus",cn,adOpenForwardOnly
			do while not rs.EOF
				strPilotStatusText = rs("name") & ""
	
				if trim(rs("ID")) = trim(PilotStatusID) and blnAllSamePilotStatus then
					strPilotStatusList = strPilotStatusList & "<option selected value=""" & rs("ID") & """>" & strPilotStatusText & "</option>"
				else		
					strPilotStatusList = strPilotStatusList & "<option value=""" & rs("ID") & """>" & strPilotStatusText & "</option>"
				end if
				
				if trim(rs("DateField")) & "" = "2" then
					strPilotDateShow = strPilotDateShow & "," & trim(rs("ID"))
				elseif	trim(rs("DateField")) & "" = "1" then
					strPilotDateShow = strPilotDateShow & "," & trim(rs("ID"))
					strPilotDateRequired = strPilotDateRequired & "," & trim(rs("ID"))
				end if
				
				rs.movenext
			loop
			rs.Close
		
		end if



		strAccessoryStatusText = ""
		if blnAllSameAccessoryStatus then
			strAccessoryStatusList = ""
		else
			strAccessoryStatusList = "<Option value=""""></OPTION>"
		end if
		if blnAccessoryPM then
		
			strAccessoryDateShow = ""
			strAccessoryDateRequired = ""

		
			rs.Open "spListAccessoryStatus",cn,adOpenForwardOnly
			do while not rs.EOF
				strAccessoryStatusText = rs("name") & ""
	
				if trim(rs("ID")) = trim(AccessoryStatusID) and blnAllSameAccessoryStatus then
					strAccessoryStatusList = strAccessoryStatusList & "<option selected value=""" & rs("ID") & """>" & strAccessoryStatusText & "</option>"
				else		
					strAccessoryStatusList = strAccessoryStatusList & "<option value=""" & rs("ID") & """>" & strAccessoryStatusText & "</option>"
				end if
				
				if trim(rs("DateField")) & "" = "2" then
					strAccessoryDateShow = strAccessoryDateShow & "," & trim(rs("ID"))
				elseif	trim(rs("DateField")) & "" = "1" then
					strAccessoryDateShow = strAccessoryDateShow & "," & trim(rs("ID"))
					strAccessoryDateRequired = strAccessoryDateRequired & "," & trim(rs("ID"))
				end if
				
				rs.movenext
			loop
			rs.Close
			if blnAddLinkStatus then
				strAccessoryStatusList = strAccessoryStatusList &  "<option value=""-1"">Link to Commodity Status</option>"
			end if		
		end if

%>

<font face=verdana size=2><b>
<label ID="lblTitle">
<%if blnShowProducts then%>
	<%=strRootName%>
	<%if blnCommodityPM and blnPilotEngineer and blnAccessoryPM then%>
		 Product Status<BR></b>
	<%elseif blnPilotEngineer then%>
		 Product Pilot Status<BR></b>
	<%elseif blnCommodityPM then%>
		 Product Qualification Status<BR></b>
	<%elseif blnAccessoryPM then%>
		 Product Accessory Status<BR></b>
	<%end if%>
	<Table border=1 cellspacing=0 cellpadding=2 bgcolor=cornsilk bordercolor=tan><TR><TD>
	<b>Vendor:</b></td><td><%=strVendor%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD>
	<b>HW,FW,Rev:</b></td><td><%=strVersion%></TD></tr><TR><TD>
	<b>Part:</b></td><td><%=strPart%>&nbsp;</TD><TD>
	<b>Model:</b></td><td><%=strModel%>&nbsp;</td></tr>
	</table>
<%else%>
	<%=strProduct%> - 
	<% if strRootName <> "" and Request("RootID") <> "" then%>
		<%=strRootName%> 
	<%else%>
		Edit Deliverable
	<% end if %>
	<%if blnCommodityPM and blnPilotEngineer and blnAccessoryPM then%>
		Status
	<%elseif blnPilotEngineer then%>
		Pilot Status
	<%elseif blnAccessoryPM then%>
		Accessory Status
	<%else %>
		Qualification Status
	<%end if%>

<%end if%>
</label></b></font>
    <% if Request("hdnApp") = "PulsarPlus" then %>
	    <form id="frmStatus" method="post" action="MultiTestStatusSave.asp?app=PulsarPlus">
    <% else %>
        <form id="frmStatus" method="post" action="MultiTestStatusSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
    <%end if%>
<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<%if blnCommodityPM then%>
	<tr>
		<td valign=top width=10 nowrap><b>Qualification&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td width=100%>
			<SELECT style="width:160" id=cboStatus name=cboStatus language=javascript onchange="cboStatus_onchange();">
			<%if blnServicePM then%>
				<option value="" selected></option>
			<%elseif blnAllSameStatus then%>
				<option value=0>Not Used Now</option>
			<%else%>
				<option value="" selected></option>
				<option value=0>Not Used Now</option>
			<%end if%>

			<%=strStatusList%>
			</SELECT>
			<%
				'if trim(TestStatusID)="0" or trim(TestStatusID)="" then
				'	strSupportDisplay = ""
				'else
					strSupportDisplay = "none"
				'end if
			%>
			<%if trim(TestStatusID)="5" and blnAllSameStatus then %>
				<Span ID=RiskReleaseRow><INPUT type="checkbox" id=chkRiskRelease name=chkRiskRelease>Risk Release</Span>
			<%else%>
				<Span ID=RiskReleaseRow style="display:none"><INPUT type="checkbox" id=chkRiskRelease name=chkRiskRelease>Risk Release</Span>
			<%end if%>
			<span id=SupportRow style=display:<%=strSupportDisplay%>><INPUT type="checkbox" id=chkDelete name=chkDelete disabled>Completely remove support</span>
			
			
			<%
				if trim(TestStatusID)="3" or strStatusSelected = "OOC" or strStatusSelected = "FCS" then
					strConfidenceDisplay = ""
				else
					strConfidenceDisplay = "none"
				end if
			%>
			
			<Span ID=ConfidenceRow style="display:<%=strConfidenceDisplay%>">&nbsp;<font size=2 face=verdana ><b>Confidence:</b>&nbsp;
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
	
    	<%if trim(TestStatusID)="3" and blnAllSameStatus then %>
		    <tr ID=DateRow>
	    <%else%>
		    <tr ID=DateRow style="display:none">
	    <%end if%>
		<td valign=top width=120 nowrap><b>Qualification&nbsp;Date:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtTestDate name=txtTestDate value="<%=strDate%>">&nbsp;<a href="javascript: cmdDate_onclick(1)"><img ID="picTarget" SRC="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Qualification&nbsp;Comments:</b>&nbsp;</td>
		<td>
			<INPUT style="width:100%" type="text" id=txtTestComments name=txtTestComments value="" maxlength=255>
		</td>
	</tr>
	
	<%end if%>
	
	<%if blnPilotEngineer then%>
	<tr>
		<td valign=top width=10 nowrap><b>Pilot&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td width=100%>
			<SELECT style="width:160" id=cboPilotStatus name=cboPilotStatus language=javascript onchange="cboPilotStatus_onchange();">
			<%=strPilotStatusList%>
			</SELECT>
			
		</td>
	</tr>
	
	<%if instr("," & strPilotDateShow & ",","," & trim(PilotStatusID) & ",") > 0 and blnAllSamePilotStatus then %>
		<tr ID=PilotDateRow>
	<%else%>
		<tr ID=PilotDateRow style="display:none">
	<%end if%>
		<td valign=top width=120 nowrap><b>Pilot&nbsp;Date:</b>&nbsp;<span ID=PilotDateStar><font color="red" size="1">*</font></Span>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtPilotDate name=txtPilotDate value="<%=strPilotDate%>">&nbsp;<a href="javascript: cmdDate_onclick(2)"><img ID="picTarget" SRC="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</td>
	</tr>
	

	<tr>
		<td valign=top width=120 nowrap><b>Pilot&nbsp;Comments:</b>&nbsp;</td>
		<td>
			<INPUT style="width:100%" type="text" id=txtPilotComments name=txtPilotComments value="" maxlength=255>
		</td>
	</tr>
	

	
	<%end if%>

	<%if blnAccessoryPM then%>
	<tr>
		<td valign=top width=10 nowrap><b>Accessory&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td width=100%>
			<SELECT style="width:160" id=cboAccessoryStatus name=cboAccessoryStatus language=javascript onchange="cboAccessoryStatus_onchange();">
			<%=strAccessoryStatusList%>
			</SELECT>
			
		</td>
	</tr>
	
	<%if instr("," & strAccessoryDateShow & ",","," & trim(AccessoryStatusID) & ",") > 0 and blnAllSameAccessoryStatus then %>
		<tr ID=AccessoryDateRow>
	<%else%>
		<tr ID=AccessoryDateRow style="display:none">
	<%end if%>
		<td valign=top width=120 nowrap><b>Accessory&nbsp;Date:</b>&nbsp;<span ID=AccessoryDateStar><font color="red" size="1">*</font></Span>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtAccessoryDate name=txtAccessoryDate value="<%=strAccessoryDate%>">&nbsp;<a href="javascript: cmdDate_onclick(3)"><img ID="picTarget" SRC="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</td>
	</tr>
	
	
	<%end if%>


	<TR>
		<TD colspan=2 valign=top><b>Deliverables Selected:</b><BR>
			<TABLE bgcolor=white width=100% border=1>
			<THEAD>
			<TR bgcolor=gainsboro>
			<TD width=10>&nbsp;</td>
			
			<%if blnShowProducts then%>
				<TD><b>Product</b></TD>
				<TD><b>Qualification</b></TD>
				<TD><b>Pilot</b></TD>
				<TD><b>Accessory</b></TD>
				<TD width=100%><b>Comments</b></TD></TR>
			<%else%>
				<TD><b>ID</b></TD>
				<TD><b>Qualification</b></TD><TD><b>Pilot</b></TD><TD><b>Accessory</b></TD>
				<TD><b>Available&nbsp;Until</b></TD>
				<%if Request("RootID") = "" then%>
					<TD><b>Deliverable</b></TD>
				<%end if%>
				<TD><b>Vendor</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD width=100%><b>Comments</b></TD></TR>
			<%end if%>
			</THEAD>
				<%=strDeliverableList%>
			</TABLE>
		</TD>
	</TR>
</table>
<%
	if blnAllSameStatus then
		strNewTestStatus = trim(TestStatusID)
	else
		strNewTestStatus = ""
	end if	
	if blnAllSamePilotStatus then
		strNewPilotStatus = trim(PilotStatusID)
	else
		strNewPilotStatus = ""
	end if
	if blnAllSameAccessoryStatus then
		strNewAccessoryStatus = trim(AccessoryStatusID)
	else
		strNewAccessoryStatus = ""
	end if

%>
<INPUT type="hidden" id=NewValue name=NewValue value="<%=strNewTestStatus%>"><!--<%=trim(TestStatusID)%>-->
<INPUT type="hidden" id=NewPilotValue name=NewPilotValue value="<%=strNewPilotStatus%>"><!--<%=trim(PilotStatusID)%>-->
<INPUT type="hidden" id=NewAccessoryValue name=NewAccessoryValue value="<%=strNewAccessoryStatus%>"><!--<%=trim(PilotStatusID)%>-->
<INPUT type="hidden" id=txtStatusText name=txtStatusText value="">
<INPUT type="hidden" id=txtPilotStatusText name=txtPilotStatusText value="">
<INPUT type="hidden" id=txtAcessoryStatusText name=txtAccessoryStatusText value="">
<INPUT type="Hidden" id=Remaining name=Remaining value="2"><!--Tell the save reoutine not to refresh-->
<INPUT type="Hidden" id=txtCommodityPM name=txtCommodityPM value="<%=blnCommodityPM%>">
<INPUT type="Hidden" id=txtPilotEngineer name=txtPilotEngineer value="<%=blnPilotEngineer%>">
<INPUT type="Hidden" id=txtAccessoryPM name=txtAccessoryPM value="<%=blnAccessoryPM%>">
</form>

<%
	end if

	set rs = nothing
	cn.Close
	set cn = nothing
%>
<INPUT type="hidden" id=txtPilotDateShow name=txtPilotDateShow value="<%=strPilotDateShow%>">
<INPUT type="hidden" id=txtAccessoryDateShow name=txtAccessoryDateShow value="<%=strAccessoryDateShow%>">
<INPUT type="hidden" id=txtPilotDateRequired name=txtPilotDateRequired value="<%=strPilotDateRequired%>">
<INPUT type="hidden" id=txtAccessoryDateRequired name=txtAccessoryDateRequired value="<%=strAccessoryDateRequired%>">

<INPUT type="hidden" id=txtUpdatableVersionCount name=txtUpdatableVersionCount value="<%=UpdatableVersionCount%>">
 <INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=Request.form("app")%>">
</BODY>
</HTML>
