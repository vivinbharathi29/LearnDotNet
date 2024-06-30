<%@ Language=VBScript %>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <title></title>
    <link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
    <script id="clientEventHandlersJS" type="text/javascript">
    <!--

    var KeyString = "";
    
    $(document).ready(function () {
        if ($('#txtMultiID').not(':checked').length == 0) {
            $('#chkAll').prop('checked', true);
        }

        $('#chkAll').click(function () {            
            if ($(this).is(':checked')) {
                $(".chkProductDeliverableReleaseID").each(function () {
                    $(this).prop('checked', true);
                });

            }
            else {
                $(".chkProductDeliverableReleaseID").each(function () {
                    $(this).prop('checked', false);
                });
            }
        });

        $(".chkProductDeliverableReleaseID").click(function () {
            if ($('.chkProductDeliverableReleaseID:checked').length == $('.chkProductDeliverableReleaseID').length) {
                $("#chkAll").prop('checked', true);
            }
            else {
                $("#chkAll").prop('checked', false);
            }
        });
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

    function SwitchRelease(ProdID, VersionList, RootID, ProductVersionReleaseID, BSID) {
        document.location = "MultiTestStatusMainPulsar.asp?ProdID=" + ProdID + "&VersionList=" + VersionList + "&RootID=" + RootID + "&Type=&ProductVersionReleaseID=" + ProductVersionReleaseID + "&BSID=" + BSID + "&ShowOnlyTargetedRelease=" + $("#txtShowOnlyTargetedRelease").val() + "&TodayPageSection=" + $("#txtTodayPageSection").val();
        window.parent.RepositionPopup();
    }

    //-->
    </script>
    <style>
        TD {
            VERTICAL-ALIGN: top;
        }
    </style>
    <link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">
</head>

<body style="background-color:ivory">
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
    dim strReleaseLink 

    strReleaseLink  = ""
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

    Dim ShowOnlyTargetedRelease
    ShowOnlyTargetedRelease = Request.QueryString("ShowOnlyTargetedRelease")
    If (ShowOnlyTargetedRelease = "" Or Not IsNumeric(ShowOnlyTargetedRelease)) Then
        ShowOnlyTargetedRelease = 0
    end if

    Dim ProductVersionReleaseID
    ProductVersionReleaseID = Request.QueryString("ProductVersionReleaseID")
    If (ProductVersionReleaseID = "" Or Not IsNumeric(ProductVersionReleaseID)) Then
        ProductVersionReleaseID = 0
    end if

    Dim BSID
    BSID = Request.QueryString("BSID")
    If (BSID = "" Or Not IsNumeric(BSID)) Then
        BSID = 0
    end if

    Dim RootID
    RootID = Request.QueryString("RootID")
    If (RootID = "" Or Not IsNumeric(RootID)) Then
        RootID = 0
    end if 

    Dim ProductID
    ProductID = Request.QueryString("ProdID")
    If (ProductID = "" Or Not IsNumeric(ProductID)) Then
        ProductID = 0
    end if
     
    dim pdids, pdrids, NotSpecificProduct
    pdids = ""
    pdrids = ""
    NotSpecificProduct = false

    if InStr(Request("VersionList"),"_") > 0 then
        NotSpecificProduct = true
        dim arr
        arr = Split(Request("VersionList"),",")
        dim arrID
        if UBound(arr) > 0 then 
            For i = 0 to uBound(arr)                        
                arrID = Split(arr(i),"_")
                if arrID(1) > 0 then
                    if pdrids <> "" then
                        pdrids = pdrids & ","
                    end if
                    pdrids = pdrids & arrID(1) 
                else 
                    if pdids <> "" then
                        pdids = pdids & ","
                    end if
                    pdids = pdids & arrID(0)                   
                end if                
            Next
       else 
            arrID = Split(arr(0),"_")
            if arrID(1) > 0 then
                pdrids = arrID(1) 
            else 
                pdids = arrID(0)               
            end if        
       end if       
    end if

    if pdids = "" then
        pdids = "0"
    end if

     if pdrids = "" then
        pdrids = "0"
    end if

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
        
		blnCommodityPM = rs("CommodityPM")
		blnPilotEngineer = rs("SCFactoryEngineer") 
		blnAccessoryPM = rs("AccessoryPM")
        blnEngineeringCoordinator = rs("EngCoordinator")

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

    'end get user


    'Get hardware team access list
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

    'end get hardware team access list

    'Get Product and releases
    strProduct = ""
	if ProductID <> ""  and ProductID <> "0" then
        dim blnfusionrequirements
        blnfusionrequirements = false
        rs.Open "select ODMHWPMID, fusionrequirements = isnull(fusionrequirements,0) from ProductVersion where id=" & clng(ProductID),cn,adOpenForwardOnly
	    if not (rs.EOF and rs.BOF) then
            if rs("fusionrequirements") then
                blnfusionrequirements = true
            end if
		    if trim(rs("ODMHWPMID") & "") = trim(CurrentUserID) then
                blnOdmHwPM = true
            end if
	    end if
	    rs.Close

		rs.Open "spGetProductVersionName " & clng(ProductID),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strProduct = rs("Name") & ""
		end if
		rs.Close         
        
        if Request("TodayPageSection") = "" and not NotSpecificProduct then
            if RootID <> "" and RootID <> "0" then
                strSql = "select Distinct pvr.Name, ReleaseID = pvr.ID, pvr.ReleaseYear, pvr.ReleaseMonth " &_
                         "from Product_Deliverable pd " &_
                         "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                         "inner join DeliverableVersion dv on dv.ID = pd.DeliverableVersionID " &_
                         "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                         "where pd.ProductVersionID = " & clng(ProductID) & " and dv.DeliverableRootID = " & RootID
            else 
                strSql = "select Distinct pvr.Name, ReleaseID = pvr.ID, pvr.ReleaseYear, pvr.ReleaseMonth " &_
                         "from Product_Deliverable pd " &_
                         "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                         "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                         "where pd.ProductVersionID = " & clng(ProductID)
            end if

            if ShowOnlyTargetedRelease = 1 then
                strSql = strSql + " and pdr.targeted = 1"
            end if
            
            strSql = strSql + " order by pvr.Name, pvr.ID, pvr.ReleaseYear desc, pvr.ReleaseMonth desc"

            rs.open strSql, cn
        elseif Not NotSpecificProduct then
            strSql = "select Distinct pvr.Name, ReleaseID = pvr.ID, pvr.ReleaseYear, pvr.ReleaseMonth " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "inner join ProductVersion p on p.ID = pd.ProductVersionID where p.fusionrequirements = 1 and pvr.ID=" + ProductVersionReleaseID
     
            if BSID > 0 then
                strSql = strSql + " and p.BusinessSegmentID = " & BSID
            end if

            if ShowOnlyTargetedRelease = 1 then
                strSql = strSql + " and pdr.targeted = 1"
            end if
            
            strSql = strSql + " order by pvr.Name, pvr.ID, pvr.ReleaseYear desc, pvr.ReleaseMonth desc"

            rs.open strSql, cn
	    end if       

        strReleaseLink = ""
        If rs.State = adStateOpen then
            Do until rs.EOF            
                if strReleaseLink <> "" then
                    strReleaseLink = strReleaseLink & " | " 
                end if
        
                if clng(ProductVersionReleaseID) = clng(rs("ReleaseID")) then
                    strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & ProductID & ",'" & trim(Request("VersionList")) & "'," & RootID & "," & rs("ReleaseID") & "," & BSID & ");"">" & rs("Name") & "</a>"
                end if
            
                rs.MoveNext
            Loop

            if ProductVersionReleaseID > 0 then
               strReleaseLink = "Switch Releases:&nbsp;<a href=""#"" onclick=""SwitchRelease(" & ProductID & ",'" & trim(Request("VersionList")) & "'," & RootID & "," & clng(ProductVersionReleaseID) & "," & BSID & ");"">All</a> | " & strReleaseLink
            elseif strReleaseLink <> "" then
               strReleaseLink = "Switch Releases:&nbsp;All | " &  strReleaseLink
            end if
            rs.Close
        end if
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
	
	if RootID <> "" and RootID <> "0" then
		blnEnoughInfo = true

        strSQL = "Select ProductDeliverableReleaseID = 0, '' as Release, v.tts, pd.targetnotes, pd.riskrelease, pd.accessorynotes, pd.pilotnotes, v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, " &_
                 "pd.accessorystatusid, pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate, t.id as StatusID, t.status, pd.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval,pd.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                 "from   dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _                         
                         "dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pd.TestStatusID inner join " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID " & _
                   "WHERE   (r.ID = " & clng(RootID) & ") AND (pd.ProductVersionID = " & clng(ProductID) & ") and isnull(pv.fusionrequirements,0) = 0 "

                
          strSQL = strSQL & "Union Select ProductDeliverableReleaseID = pdr.ID, pvr.Name as Release, v.tts,pdr.targetnotes, pdr.riskrelease, pdr.accessorynotes, pdr.pilotnotes, v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, " &_
                 "pdr.accessorystatusid, pdr.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pdr.PilotDate, t.id as StatusID, t.status, pdr.testdate, pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pdr.WWANTestStatus, c.requiresWWANtestfinalapproval,c.requiresdeveloperfinalapproval, pdr.IntegrationTestStatus,c.requiresMITtestfinalapproval, pdr.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct, pdr.WWANTestStatus,c.requiresWWANtestfinalapproval,pdr.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pdr.IntegrationTestStatus,c.requiresMITtestfinalapproval,pdr.ODMTestStatus,c.requiresodmtestfinalapproval  " & _
                 "from   dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _
                         "dbo.Product_Deliverable_Release AS pdr WITH (NOLOCK) ON pdr.ProductDeliverableID = pd.ID LEFT OUTER JOIN " &_
                         "dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pdr.TestStatusID inner join " & _
                         "dbo.ProductVersionRelease AS pvr WITH (NOLOCK) ON pvr.ID = pdr.ReleaseID INNER JOIN " &_
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pdr.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pdr.AccessoryStatusID = at.ID " & _
                   "WHERE   (r.ID = " & clng(RootID) & ") AND (pd.ProductVersionID = " & clng(ProductID) & ") and pv.fusionrequirements = 1 "   

          if ShowOnlyTargetedRelease = 1 then
            strSql = strSql + " and pdr.targeted = 1"
          end if

          if clng(ProductVersionReleaseID) > 0 then
            strSql = strSql + " and pvr.id = " &  ProductVersionReleaseID      
          end if

          strSQL = strSQL & " ORDER BY r.Name, v.id, pd.ID, ProductDeliverableReleaseID, vd.Name DESC;"      

	elseif trim(ProductID) <> "" and trim(ProductID) <> "0" then
		blnEnoughInfo = true

        strSQl = "Select ProductDeliverableReleaseID = 0, '' as Release, v.tts, targetnotes = isnull(pd.targetnotes,''), riskrelease = isnull(pd.riskrelease,0), accessorynotes = isnull(pd.accessorynotes,''), pd.pilotnotes, " &_
                    "v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, pt.id as PilotStatusID, accessorystatusid = isnull(pd.accessorystatusid,0), " &_
                    "accessorydate = isnull(pd.accessorydate,''), pt.Name as PilotStatus, pd.PilotDate, t.id as StatusID, t.status, testdate = isnull(pd.testdate,''), pd.id, r.name, v.id as VersionID, " &_
                    "v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pd.WWANTestStatus,c.requiresWWANtestfinalapproval, " &_
                    "c.requiresdeveloperfinalapproval, pd.IntegrationTestStatus, c.requiresMITtestfinalapproval, pd.ODMTestStatus, " &_
                    "c.requiresodmtestfinalapproval, pv.wwanproduct, pd.WWANTestStatus, c.requiresWWANtestfinalapproval, pd.DeveloperTestStatus, c.requiresdeveloperfinalapproval, " &_
                    "pd.IntegrationTestStatus, c.requiresMITtestfinalapproval, pd.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                    "FROM dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                    "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                    "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                    "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                    "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _                   
                    "dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pd.TestStatusID INNER JOIN " & _                    
                    "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                    "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                    "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID " & _
                    "WHERE v.ID IN (" & scrubsql(request("VersionList")) & ") " & _
                    "AND (pd.ProductVersionID = " &  clng(ProductID) & ") and isnull(pv.fusionrequirements,0) = 0 "                    

        
        strSQl = strSQl & "Union Select ProductDeliverableReleaseID = pdr.ID, pvr.Name as Release, v.tts, targetnotes = isnull(pdr.targetnotes,''), riskrelease = isnull(pdr.riskrelease,0), accessorynotes = isnull(pdr.accessorynotes,''), pdr.pilotnotes, " &_
                    "v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity, pt.id as PilotStatusID, accessorystatusid = isnull(pdr.accessorystatusid,0), " &_
                    "accessorydate = isnull(pdr.accessorydate,''), pt.Name as PilotStatus, pd.PilotDate, t.id as StatusID, t.status, testdate = isnull(pdr.testdate,''), pd.id, r.name, v.id as VersionID, " &_
                    "v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pdr.WWANTestStatus,c.requiresWWANtestfinalapproval, " &_
                    "c.requiresdeveloperfinalapproval, pdr.IntegrationTestStatus, c.requiresMITtestfinalapproval, pdr.ODMTestStatus, " &_
                    "c.requiresodmtestfinalapproval, pv.wwanproduct, pdr.WWANTestStatus, c.requiresWWANtestfinalapproval, pdr.DeveloperTestStatus, c.requiresdeveloperfinalapproval, " &_
                    "pdr.IntegrationTestStatus, c.requiresMITtestfinalapproval, pdr.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                    "FROM dbo.DeliverableVersion AS v WITH (NOLOCK) INNER JOIN " & _
                    "dbo.DeliverableRoot AS r WITH (NOLOCK) ON v.DeliverableRootID = r.ID INNER JOIN " & _
                    "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                    "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                    "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _
                    "dbo.Product_Deliverable_Release AS pdr WITH (NOLOCK) ON pd.ID = pdr.ProductDeliverableID LEFT OUTER JOIN " &_
                    "dbo.TestStatus AS t WITH (NOLOCK) ON t.ID = pdr.TestStatusID INNER JOIN " & _
                    "dbo.ProductVersionRelease AS pvr WITH (NOLOCK) ON pvr.ID = pdr.ReleaseID INNER JOIN " &_
                    "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID INNER JOIN " & _
                    "dbo.PilotStatus AS pt WITH (NOLOCK) ON pdr.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                    "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pdr.AccessoryStatusID = at.ID " & _
                    "WHERE v.ID IN (" & scrubsql(request("VersionList")) & ") " & _
                    "AND (pd.ProductVersionID = " &  clng(ProductID) & ") and pv.fusionrequirements = 1 "                    

        
        if ShowOnlyTargetedRelease = 1 then
            strSql = strSql + " and pdr.targeted = 1"
        end if

        if clng(ProductVersionReleaseID) > 0 then
            strSql = strSql + " and pvr.id = " &  ProductVersionReleaseID     
        end if

        strSQL = strSQL & " ORDER BY r.Name, v.id, pd.id, ProductDeliverableReleaseID, vd.Name DESC;"
	elseif trim(Request("VersionList")) <> "" and trim(Request("VersionList")) <> "0" and instr(Request("VersionList"),",")=0 and instr(Request("VersionList"),"_")=0 then
		blnEnoughInfo = true
		blnShowProducts = true

		strSQl = "Select  ProductDeliverableReleaseID = 0, '' as Release, v.tts, pd.targetnotes, pd.riskrelease, pd.accessorynotes, pd.pilotnotes, v.active as EOL, at.name as AccessoryStatus, " &_
                 "v.EndOfLifeDate as EOLDate, c.commodity, pv.DotsName as Product, pd.accessorystatusid, pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate,  t.id as StatusID, t.status, pd.testdate, " &_
                 "pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pd.WWANTestStatus, c.requiresWWANtestfinalapproval, " &_
                 "c.requiresdeveloperfinalapproval,pd.IntegrationTestStatus,c.requiresMITtestfinalapproval,pd.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pd.WWANTestStatus,c.requiresWWANtestfinalapproval, " &_
                 "pd.DeveloperTestStatus, c.requiresdeveloperfinalapproval, pd.IntegrationTestStatus, c.requiresMITtestfinalapproval, pd.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                 "FROM    dbo.DeliverableRoot AS r WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID INNER JOIN " & _                                             
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pd.AccessoryStatusID = at.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _
                  "WHERE v.ID = " & clng(request("VersionList")) & " and isnull(pv.fusionrequirements,0) = 0 "

		if CurrentUserPartnerID <> 1 then
		    strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID & " " 
		end if

        strSQl = strSQl & "Union Select  ProductDeliverableReleaseID = pdr.ID, pvr.Name as Release, v.tts, pdr.targetnotes, pdr.riskrelease, pdr.accessorynotes, pdr.pilotnotes, v.active as EOL, at.name as AccessoryStatus, " &_
                 "v.EndOfLifeDate as EOLDate, c.commodity, pv.DotsName as Product, pdr.accessorystatusid, pdr.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pdr.PilotDate,  t.id as StatusID, t.status, pdr.testdate, " &_
                 "pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pdr.WWANTestStatus, c.requiresWWANtestfinalapproval, " &_
                 "c.requiresdeveloperfinalapproval,pdr.IntegrationTestStatus,c.requiresMITtestfinalapproval,pdr.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pdr.WWANTestStatus,c.requiresWWANtestfinalapproval, " &_
                 "pdr.DeveloperTestStatus, c.requiresdeveloperfinalapproval, pdr.IntegrationTestStatus, c.requiresMITtestfinalapproval, pdr.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                 "FROM    dbo.DeliverableRoot AS r WITH (NOLOCK) INNER JOIN " & _
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID INNER JOIN " & _
                         "dbo.Product_Deliverable_Release AS pdr WITH (NOLOCK) ON pdr.ProductDeliverableID = pd.ID LEFT OUTER JOIN " &_
                         "dbo.ProductVersionRelease AS pvr WITH (NOLOCK) ON pvr.ID = pdr.ReleaseID INNER JOIN " &_
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON pdr.AccessoryStatusID = at.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pdr.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pdr.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID " & _
                  "WHERE v.ID = " & clng(request("VersionList")) & " and pv.fusionrequirements = 1 "

		if CurrentUserPartnerID <> 1 then
		    strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID & " " 
		end if

        if clng(ProductVersionReleaseID) > 0 then
            strSQL = strSQL & " and pvr.ID=" & clng(ProductVersionReleaseID)        
        end if

        if ShowOnlyTargetedRelease = 1 then
            strSql = strSql + " and pdr.targeted = 1"
        end if

		strSQl = strSQL &  " order by  pv.DotsName;"
	elseif trim(Request("VersionList")) <> "" and trim(Request("VersionList")) <> "0" then
		
        blnEnoughInfo = true
        if not NotSpecificProduct then
		    blnShowProducts = true
        end if
	            
		strSQl = "Select ProductDeliverableReleaseID = 0, Release = '', v.tts, pd.targetnotes, pd.riskrelease, pd.accessorynotes, pd.pilotnotes, v.active as EOL, at.name as AccessoryStatus, v.EndOfLifeDate as EOLDate, c.commodity,  " &_
                 "pv.DotsName as Product, pd.accessorystatusid, pd.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pd.PilotDate,  t.id as StatusID, t.status, pd.testdate, " &_
                 "pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pd.WWANTestStatus, c.requiresWWANtestfinalapproval, " &_
                 "c.requiresdeveloperfinalapproval, pd.IntegrationTestStatus, c.requiresMITtestfinalapproval, pd.ODMTestStatus,c.requiresodmtestfinalapproval, pv.wwanproduct, pd.WWANTestStatus, c.requiresWWANtestfinalapproval, " &_
                 "pd.DeveloperTestStatus, c.requiresdeveloperfinalapproval, pd.IntegrationTestStatus, c.requiresMITtestfinalapproval, pd.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                 "FROM    dbo.Product_Deliverable AS pd WITH (NOLOCK) LEFT OUTER JOIN " & _                         
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON at.ID = pd.AccessoryStatusID INNER JOIN " & _                         
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pd.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pd.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID  " & _
                 "WHERE pd.ID IN (" & scrubsql(pdids) & ") and isnull(pv.fusionrequirements,0) = 0 "

		if CurrentUserPartnerID <> 1 then
		    strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID 
		end if
        
        strSQL = strSQL & " Union "
     
        strSQl = strSQL & "Select ProductDeliverableReleaseID = pdr.ID, pvr.Name as Release, v.tts, pdr.targetnotes, pdr.riskrelease, pdr.accessorynotes, pdr.pilotnotes, v.active as EOL, at.name as AccessoryStatus,  v.EndOfLifeDate as EOLDate, c.commodity,  " &_
                 "pv.DotsName as Product, pdr.accessorystatusid, pdr.accessorydate, pt.id as PilotStatusID, pt.Name as PilotStatus, pdr.PilotDate,  t.id as StatusID, t.status, pdr.testdate, " &_
                 "pd.id, r.name, v.id as VersionID, v.version, v.revision, v.pass, v.modelnumber, v.partnumber, vd.name as vendor, v.location, pdr.WWANTestStatus, c.requiresWWANtestfinalapproval, " &_
                 "c.requiresdeveloperfinalapproval, pdr.IntegrationTestStatus, c.requiresMITtestfinalapproval, pdr.ODMTestStatus,c.requiresodmtestfinalapproval,pv.wwanproduct,pdr.WWANTestStatus,c.requiresWWANtestfinalapproval, " &_
                 "pdr.DeveloperTestStatus,c.requiresdeveloperfinalapproval,pdr.IntegrationTestStatus, c.requiresMITtestfinalapproval, pdr.ODMTestStatus, c.requiresodmtestfinalapproval  " & _
                 "FROM    dbo.Product_Deliverable AS pd WITH (NOLOCK) INNER JOIN " & _
                         "dbo.Product_Deliverable_Release AS pdr WITH (NOLOCK) ON pdr.ProductDeliverableID = pd.ID LEFT OUTER JOIN " &_
                         "dbo.AccessoryStatus AS at WITH (NOLOCK) ON at.ID = pd.AccessoryStatusID INNER JOIN " & _
                         "dbo.ProductVersionRelease AS pvr WITH (NOLOCK) ON pvr.ID = pdr.ReleaseID INNER JOIN " &_
                         "dbo.DeliverableVersion AS v WITH (NOLOCK) ON pd.DeliverableVersionID = v.ID INNER JOIN " & _
                         "dbo.DeliverableRoot AS r WITH (NOLOCK) ON r.ID = v.DeliverableRootID INNER JOIN " & _
                         "dbo.Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID INNER JOIN " & _
                         "dbo.PilotStatus AS pt WITH (NOLOCK) ON pdr.PilotStatusID = pt.ID LEFT OUTER JOIN " & _
                         "dbo.TestStatus AS t WITH (NOLOCK) ON pdr.TestStatusID = t.ID INNER JOIN " & _
                         "dbo.DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID INNER JOIN " & _
                         "dbo.ProductVersion AS pv WITH (NOLOCK) ON pd.ProductVersionID = pv.ID  " & _
                 "WHERE pv.fusionrequirements = 1 "   

        if pdrids <> "" then 
            strSQL = strSQL & " and pdr.ID IN (" & scrubsql(pdrids) & ") "
        else 
            strSQL = strSQL & " and pd.ID IN (" & scrubsql(pdids) & ") "
        end if

        if CurrentUserPartnerID <> 1 then
		    strSQL = strSQL & " and pv.PartnerID = " & CurrentUserPartnerID 
		end if        
        
        if ShowOnlyTargetedRelease = 1 then
            strSql = strSql + " and pdr.targeted = 1"
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
				if RootID <> "" then
					strDeliverableList = strDeliverableList & "<TR><TD><INPUT type=""checkbox"" QCompleteOK=""" & blnQCompleteOK & """ style=""Width:16;Height:16"" TestStatus=" & TestStatusID & " id=txtMultiID name=txtMultiID class=chkProductDeliverableReleaseID value=""" & rs("ID") & "_" & rs("ProductDeliverableReleaseID") & """></TD>" '" & replace(replace(blnQCompleteOK,"False",""),"True",".") & "
				else
				    strDeliverableList = strDeliverableList & "<TR><TD><INPUT type=""checkbox"" QCompleteOK=""" & blnQCompleteOK & """ style=""Width:16;Height:16"" TestStatus=" & TestStatusID & " checked id=txtMultiID name=txtMultiID class=chkProductDeliverableReleaseID value=""" & rs("ID") & "_" & rs("ProductDeliverableReleaseID") & """></TD>"
				end if

				if blnShowProducts then
					strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Product") & "&nbsp;</td>"
                    if clng(ProductVersionReleaseID) = 0 then
                        strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Release") & "&nbsp;</td>"
                    end if
				else
					strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("VersionID") & "&nbsp;</td>"
                    if NotSpecificProduct then
                        strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Product") & "&nbsp;</td>"
                    end if
                    strDeliverableList = strDeliverableList & "<TD nowrap>" & rs("Release") & "&nbsp;</td>"
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


				if RootID = "" then
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
		if RootID <> "" and RootID <> "0" then
			Response.Write "There are no versions currently available supporting this product."
		elseif RootID = "0" then
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

    <font face="verdana" size="2"><b>
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
        <br /><%=strReleaseLink %><br />
    <%else%>
        <%if not NotSpecificProduct then%>
	    <%=strProduct%> - 
	    <% if strRootName <> "" and RootID <> "" then%>
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

        <br /><%=strReleaseLink %>
        <%end if%>
    <%end if%>
    </label></b></font>

        <%  %>
        <form id="frmStatus" method="post" action="MultiTestStatusSavePulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">

            <table id="tabGeneral" width="100%" bgcolor="cornsilk" border="1" cellspacing="0" cellpadding="2" bordercolor="tan">
                <%if blnCommodityPM then%>
                <tr>
                    <td valign="top" width="10" nowrap><b>Qualification&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td width="100%">
                        <select style="width: 160" id="cboStatus" name="cboStatus" language="javascript" onchange="cboStatus_onchange();">
                            <%if blnServicePM then%>
                            <option value="" selected></option>
                            <%elseif blnAllSameStatus then%>
                            <option value="0">Not Used Now</option>
                            <%else%>
                            <option value="" selected></option>
                            <option value="0">Not Used Now</option>
                            <%end if%>

                            <%=strStatusList%>
                        </select>
                        <%

					    strSupportDisplay = "none"

                        %>
                        <%if trim(TestStatusID)="5" and blnAllSameStatus then %>
                        <span id="RiskReleaseRow">
                            <input type="checkbox" id="chkRiskRelease" name="chkRiskRelease">Risk Release</span>
                        <%else%>
                        <span id="RiskReleaseRow" style="display: none">
                            <input type="checkbox" id="chkRiskRelease" name="chkRiskRelease">Risk Release</span>
                        <%end if%>
                        <span id="SupportRow" style="display: <%=strSupportDisplay%>">
                            <input type="checkbox" id="chkDelete" name="chkDelete" disabled>Completely remove support</span>


                        <%
				    if trim(TestStatusID)="3" or strStatusSelected = "OOC" or strStatusSelected = "FCS" then
					    strConfidenceDisplay = ""
				    else
					    strConfidenceDisplay = "none"
				    end if
                        %>

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

                <%if trim(TestStatusID)="3" and blnAllSameStatus then %>
                <tr id="DateRow">
                    <%else%>
                <tr id="DateRow" style="display: none">
                    <%end if%>
                    <td valign="top" width="120" nowrap><b>Qualification&nbsp;Date:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                    <td>
                        <input type="text" id="txtTestDate" name="txtTestDate" value="<%=strDate%>">&nbsp;<a href="javascript: cmdDate_onclick(1)"><img id="picTarget" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21"></a>
                    </td>
                </tr>
                <tr>
                    <td valign="top" width="120" nowrap><b>Qualification&nbsp;Comments:</b>&nbsp;</td>
                    <td>
                        <input style="width: 100%" type="text" id="txtTestComments" name="txtTestComments" value="" maxlength="255">
                    </td>
                </tr>

                <%end if%>

                <%if blnPilotEngineer then%>
                <tr>
                    <td valign="top" width="10" nowrap><b>Pilot&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td width="100%">
                        <select style="width: 160" id="cboPilotStatus" name="cboPilotStatus" language="javascript" onchange="cboPilotStatus_onchange();">
                            <%=strPilotStatusList%>
                        </select>

                    </td>
                </tr>

                <%if instr("," & strPilotDateShow & ",","," & trim(PilotStatusID) & ",") > 0 and blnAllSamePilotStatus then %>
                <tr id="PilotDateRow">
                    <%else%>
                <tr id="PilotDateRow" style="display: none">
                    <%end if%>
                    <td valign="top" width="120" nowrap><b>Pilot&nbsp;Date:</b>&nbsp;<span id="PilotDateStar"><font color="red" size="1">*</font></span>&nbsp;</td>
                    <td>
                        <input type="text" id="txtPilotDate" name="txtPilotDate" value="<%=strPilotDate%>">&nbsp;<a href="javascript: cmdDate_onclick(2)"><img id="picTarget" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21"></a>
                    </td>
                </tr>


                <tr>
                    <td valign="top" width="120" nowrap><b>Pilot&nbsp;Comments:</b>&nbsp;</td>
                    <td>
                        <input style="width: 100%" type="text" id="txtPilotComments" name="txtPilotComments" value="" maxlength="255">
                    </td>
                </tr>



                <%end if%>

                <%if blnAccessoryPM then%>
                <tr>
                    <td valign="top" width="10" nowrap><b>Accessory&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td width="100%">
                        <select style="width: 160" id="cboAccessoryStatus" name="cboAccessoryStatus" language="javascript" onchange="cboAccessoryStatus_onchange();">
                            <%=strAccessoryStatusList%>
                        </select>

                    </td>
                </tr>

                <%if instr("," & strAccessoryDateShow & ",","," & trim(AccessoryStatusID) & ",") > 0 and blnAllSameAccessoryStatus then %>
                <tr id="AccessoryDateRow">
                    <%else%>
                <tr id="AccessoryDateRow" style="display: none">
                    <%end if%>
                    <td valign="top" width="120" nowrap><b>Accessory&nbsp;Date:</b>&nbsp;<span id="AccessoryDateStar"><font color="red" size="1">*</font></span>&nbsp;</td>
                    <td>
                        <input type="text" id="txtAccessoryDate" name="txtAccessoryDate" value="<%=strAccessoryDate%>">&nbsp;<a href="javascript: cmdDate_onclick(3)"><img id="picTarget" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="21"></a>
                    </td>
                </tr>


                <%end if%>


                <tr>
                    <td colspan="2" valign="top"><b>Deliverables Selected:</b><br>
                        <table bgcolor="white" width="100%" border="1">
                            <thead>
                                <tr bgcolor="gainsboro">
                                    <td width="10"><input type="checkbox" id="chkAll" style="Width:16; Height:16" /></td>

                                    <%if blnShowProducts then%>
                                    <td><b>Product</b></td>
                                        <%if clng(ProductVersionReleaseID) = 0 then %>
                                            <td><b>Release</b></td>
                                        <%end if %>
                                    <td><b>Qualification</b></td>
                                    <td><b>Pilot</b></td>
                                    <td><b>Accessory</b></td>
                                    <td width="100%"><b>Comments</b></td>
                                    </tr>
                                    <%else%>                                        
                                        <td><b>Version ID</b></td>
                                        <%if NotSpecificProduct then%>
                                            <td><b>Product</b></td>
                                        <%end if %>
                                        <td><b>Release</b></td>
                                        <td><b>Qualification</b></td>
                                        <td><b>Pilot</b></td>
                                        <td><b>Accessory</b></td>
                                        <td><b>Available&nbsp;Until</b></td>
                                        <%if RootID = "" then%>
                                            <td><b>Deliverable</b></td>
                                        <%end if%>
                                        <td><b>Vendor</b></td>
                                        <td><b>HW,FW,Rev</b></td>
                                        <td><b>Model</b></td>
                                        <td><b>Part</b></td>
                                        <td width="100%"><b>Comments</b></td>
                                    </tr>
                                    <%end if%>
			            </THEAD>
				    <%=strDeliverableList%>
            </table>
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
        <input type="hidden" id="NewValue" name="NewValue" value="<%=strNewTestStatus%>"><!--<%=trim(TestStatusID)%>-->
        <input type="hidden" id="NewPilotValue" name="NewPilotValue" value="<%=strNewPilotStatus%>"><!--<%=trim(PilotStatusID)%>-->
        <input type="hidden" id="NewAccessoryValue" name="NewAccessoryValue" value="<%=strNewAccessoryStatus%>"><!--<%=trim(PilotStatusID)%>-->
        <input type="hidden" id="txtStatusText" name="txtStatusText" value="">
        <input type="hidden" id="txtPilotStatusText" name="txtPilotStatusText" value="">
        <input type="hidden" id="txtAcessoryStatusText" name="txtAccessoryStatusText" value="">
        <input type="Hidden" id="Remaining" name="Remaining" value="2"><!--Tell the save reoutine not to refresh-->
        <input type="Hidden" id="txtCommodityPM" name="txtCommodityPM" value="<%=blnCommodityPM%>">
        <input type="Hidden" id="txtPilotEngineer" name="txtPilotEngineer" value="<%=blnPilotEngineer%>">
        <input type="Hidden" id="txtAccessoryPM" name="txtAccessoryPM" value="<%=blnAccessoryPM%>">
        <input type="hidden" id=txtTodayPageSection name=txtTodayPageSection value="<%=Request("TodayPageSection")%>">
    </form>

    <%
	end if

	set rs = nothing
	cn.Close
	set cn = nothing
    %>
    <input type="hidden" id="txtPilotDateShow" name="txtPilotDateShow" value="<%=strPilotDateShow%>">
    <input type="hidden" id="txtAccessoryDateShow" name="txtAccessoryDateShow" value="<%=strAccessoryDateShow%>">
    <input type="hidden" id="txtPilotDateRequired" name="txtPilotDateRequired" value="<%=strPilotDateRequired%>">
    <input type="hidden" id="txtAccessoryDateRequired" name="txtAccessoryDateRequired" value="<%=strAccessoryDateRequired%>">
    <input type="hidden" id="txtBSID" name="txtBSID" value="<%=request("BSID")%>" />
    <input type="hidden" id="txtUpdatableVersionCount" name="txtUpdatableVersionCount" value="<%=UpdatableVersionCount%>"> 
    <input type="hidden" id="txtShowOnlyTargetedRelease" name="txtShowOnlyTargetedRelease" value="<%=ShowOnlyTargetedRelease%>" /> 
</body>
</html>