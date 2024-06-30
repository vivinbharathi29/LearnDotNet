<%@  language="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="../includes/no-cache.asp" -->
<%
Response.Clear
Response.Buffer = False

    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim DRID : DRID = regEx.Replace(Request.QueryString("ID"), "")
    
    regEx.Pattern = "[^0-9a-zA-Z_ ]"
    Dim sList : sList = regEx.Replace(Request("List"), "")
    If sList = "" Then
		on error resume next
        sList = regEx.Replace(Request.Cookies("PMTab"), "")
        on error goto 0
    End If

    regEx.Pattern = "[^0-9a-zA-Z]"
    Dim hideForPulsar : hideForPulsar = regEx.replace(Request.QueryString("HideHeader"), "")
    
    Dim sTab : sTab = regEx.Replace(Request.QueryString("Tab"), "")
    Dim sClass : sClass = regEx.Replace(Request.QueryString("Class"), "")
    Dim sStatus : sStatus = regEx.Replace(Request.QueryString("Status"), "")
    Dim sView : sView = regEx.Replace(Request.QueryString("View"), "")
    on error resume next
    Dim sDmStatus : sDmStatus = regEx.Replace(Request.Cookies("DMStatus"), "")

    if LCASE(hideForPulsar) = "true" Then hideForPulsar = "display:none;"

	on error goto 0

Dim rs, dw, cn, cmd, strSql
on error resume next
Dim strTitleColor : strTitleColor = regex.Replace(Request.Cookies("TitleColor"),"")
if strTitleColor = "" then 
	strTitleColor = "#0000cd"
end if
on error goto 0
Dim isSysAdmin : isSysAdmin = false
Dim isCertAdmin : isCertAdmin = false
Dim isEditModeOn : isEditModeOn = false

Dim currentUser : currentUser = lcase(Session("LoggedInUser"))
Dim currentDomain
Dim currentUserPartner
Dim currentUserName
Dim currentUserID
Dim currentUserSysAdmin
Dim currentWorkgroupId
Dim faves
Dim faveCount
Dim productName
Dim displayedProductName
Dim devCenter
Dim SEPMID
Dim PMID
Dim productVersion
Dim displayedList
Dim currentUserDefaultTab
Dim productType
Dim deliverableCategoryId

Dim AppRoot
AppRoot = Session("ApplicationRoot")

If instr(currentUser,"\") > 0 Then
	currentDomain = left(currentUser, instr(currentUser,"\") - 1)
	currentUser = mid(currentUser,instr(currentUser,"\") + 1)
End If

'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim securityObj
	Set securityObj = New ExcaliburSecurity
	
	isSysAdmin = securityObj.IsSysAdmin()
    isCertAdmin = securityObj.UserInRole("", "CERTADMIN")
    
    isEditModeOn = (isSysAdmin Or isCertAdmin)
    
	Set securityObj = Nothing
'##############################################################################	


'
' Setup the data connections
'
Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "spGetUserInfo")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 80, currentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, currentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then
	currentUserName = rs("Name") & ""
	currentUserID = rs("ID") & ""
	currentUserSysAdmin = rs("SystemAdmin")
	currentWorkgroupId = rs("WorkgroupID") & ""
	currentUserPartner = trim(rs("PartnerID") & "")
	currentUserDefaultTab = rs("DefaultProductTab") & ""

	faves = trim(rs("Favorites") & "")
	faveCount = trim(rs("FavCount") & "")
End If
rs.Close

Function PrepForWeb( value )
	Dim myString
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
	    myString = Server.HTMLEncode( value )
	    myString = Replace(myString, vbCrLf, "<br />")
	    myString = Replace(myString, vbCr , "<br />")
	    myString = Replace(myString, vbLf, "<br />")
	    myString = Replace(myString, Chr(10), "<br />")
		PrepForWeb = myString
	End If

End Function

'***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
'or currentUserID = 391 or currentUserID = 397
If currentUserSysAdmin or SEPMID = currentUserID or instr(trim(PMID),"_" & trim(currentUserID) & "_") > 0  Then
	isSysAdmin = true
End If

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=9" />
    <title></title>
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <link href="<%= AppRoot %>/uploadify/uploadify.css" type="text/css" rel="stylesheet" />

    <style type="text/css">
        body
        {
        }
        #loading
        {
            text-align: center;
            font: bold medium verdana;
        }
        #body, #ToolMenu
        {
            display: none;
        }
        #status
        {
            font-size: small;
            font-weight: bold;
        }
        #status dl dd
        {
            margin-left: 10px;
        }
        #status dl dt
        {
            padding-left: 3px;
        }
        #status dt
        {
            border: 1px solid grey;
        }
        #status dl
        {
            border: 1px solid black;
        }
        .StatusTable
        {
            width: 100%;
            border-collapse: collapse;
        }
        .StatusTable tbody td, .StatusTable tbody th, .ui-state-default, .ui-widget-content
        {
            font: x-small verdana;
            padding: 2px;
        }
        .red
        {
            background-color: #de2e43;
        }
        .yellow
        {
            background-color: #fdc643;
        }
        .green
        {
            background-color: #9cdb2c;
        }
        .highlight
        {
            background-color: lightsteelblue;
            background-image: none;
        }
        #status div
        {
            width: 99.7%;
            border: solid 1px black;
            padding-left: 3px;
        }
        .StatusTable thead tr
        {
            height: 25px;
        }
        
        .Link
        {
            font-size: xx-small;
            font-weight: normal;
        }
    </style>

    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/uploadify/swfobject.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/uploadify/jquery.uploadify.v2.1.4.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/Agency/AgencyDetailsDialog.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/Agency/AgencyInitWorkflowDialog.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/Agency/AgencyUploadDialog.js" type="text/javascript"></script>

    <script type="text/javascript">
        (function ($) {
            $.fn.isNullOrEmpty = function () {
                return (this === null || this === undefined || this == '');
            };
        })(jQuery);

        $(function () {

            if ($("#isAgencyEditModeOn").val() == "True") {
                //$("#ToolMenu").show();
                $(".deleverage").show();
            } else {
                $("#ToolMenu").hide();
                $(".deleverage").hide();
            }

            //Setup Table Format
            $(".StatusTable th").addClass("ui-state-default");
            $(".StatusTable td").addClass("ui-widget-content");

            $(".StatusTable thead th").hover(function () {
                $(this).addClass("hover");
            }, function () {
                $(this).removeClass("hover");
            });

            $(".StatusTable tr").hover(function () {
                $(this).children("td").addClass("ui-state-hover2");
                $(this).addClass("hover");
            }, function () {
                $(this).children("td").removeClass("ui-state-hover2");
                $(this).removeClass("hover");
            });

            $(".StatusTable tbody tr").click(function () {
                $('#txtProjectedDate').attr("disabled", true);
                var row = $(this);
                var statusId = $(".statusId", row).val();
                if (!$(statusId).isNullOrEmpty())
                    DetailsDialog.Show(statusId);
            });

            $(".initWorkflow").click(function () {
                var certGroup = $(this).parents(".CertificationGroup");
                var deliverableVersionId = $(".deliverableVersionId", certGroup).val();
                var agencyTypeId = $(".agencyTypeId", certGroup).val();
                InitDialog.Show(deliverableVersionId, agencyTypeId);

            });

            $(".selectFollowers").click(function () {
                var certGroup = $(this).parents(".CertificationGroup");
                var deliverableVersionId = $(".deliverableVersionId", certGroup).val();
                var agencyTypeId = $(".agencyTypeId", certGroup).val();
                Leverage.Show(deliverableVersionId, agencyTypeId);
            });

            var showPastDueWarning = false;
            var showInitializationWarning = false;

            if ($(".StatusTable tbody.Uninitialized").length > 0) {
                showInitializationWarning = true;

                $(".StatusTable tbody.Uninitialized tr td").addClass("ui-state-highlight");
                $(".StatusTable tbody.Uninitialized tr td:nth-child(4)").text("");

                var certGroup = $(".StatusTable tbody.Uninitialized").parents(".CertificationGroup");
                $(".selectFollowers", certGroup).hide();

            }

            $(".CertificationGroup").each(function () {
                if ($(".Uninitialized", this).length == 0) {
                    $(".initWorkflow", this).hide();
                }
            });

            $(".StatusTable tbody.In_Progress tr").each(function () {
                var daysToTargetCell = $("td:nth-child(3)", $(this));
                var days = parseInt($(daysToTargetCell).text());
                if (days < 7) {
                    $(daysToTargetCell).parent().children().addClass("ui-state-highlight");
                }
                if (days < 0) {
                    showPastDueWarning = true;
                    $(daysToTargetCell).parent().children().addClass("ui-state-error");
                    $(daysToTargetCell).text("Past Due");
                }
            });

            $(".StatusTable tbody.Complete tr").each(function () {
                if (parseInt($(".nextStepId", $(this)).val()) == 0) {
                    daysToTargetCell = $("td:nth-child(3)", $(this)).text("Complete");
                }
            });

            $(".StatusTable thead").click(function () {
                $(this).next().toggle();
            });

            if ($(".StatusTable").length > 1) {
                $(".StatusTable tbody").hide();
            }

            if (showPastDueWarning) {
                $("#errorText").append("You have workflow steps that are past due!");
            }

            if (showInitializationWarning) {
                $("#warningText").append("You have certifications that need to be initialized.");
            }

            if ($("#errorText").text() == '') {
                $("#error").hide();
            }

            if ($("#warningText").text() == '') {
                $("#warning").hide();
            }

            $(".Link").hover(function () {
                $(this).addClass("hover");
            }, function () {
                $(this).removeClass("hover");
            });

            $("#initWorkflow").click(function () {
                InitDialog.Show();
            });

            $("#selectFollowers").click(function () {
                Leverage.Show();
            });

            $("#showFrame").click(function () {
                ShowIframeDialog();
            });

            $("#iframeDialog").dialog({
                modal: true,
                autoOpen: false,
                width: 800,
                height: 800
            });

            if ($("#isExcaliburAdmin").val() == "True") {
                $("#testToolLink").show();
                $("#testHarness").click(function () {
                    UploadDialog.Show();
                });
            }
            //Hide Loading panel and show the rest of the body.
            $("#loading").hide();
            $("#body").fadeIn('slow');
        });

        function ShowIframeDialog() {
            $("#iframeDialog iframe").attr("width", "95%");
            $("#iframeDialog iframe").attr("height", "95%");
            $("#iframeDialog iframe").attr("src", "Agency.asp?ID=5000");
            $("#iframeDialog").dialog("open");
        }

       
    </script>

    <script type="text/javascript">
        // DMView Header
        function ShowProperties(DisplayedID, strTab) {
            var strID;
            if (txtFilename.value == "HFCN")
                strID = window.showModalDialog("<%= AppRoot %>/HFCN/HFCNAdd.asp?ID=" + DisplayedID, "", "dialogWidth:600px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
            else
                strID = window.showModalDialog("<%= AppRoot %>/root.asp?ID=" + DisplayedID, "", "dialogWidth:800px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
            if (txtView.value != "1") {
                if (typeof (strID) != "undefined") {
                    window.parent.frames["RightWindow"].navigate("dmview.asp?Prog=1&ID=" + DisplayedID);
                    window.parent.frames["LeftWindow"].navigate("tree.asp?Prog=1&ID=" + DisplayedID);
                }
            }
            else {
                window.location.reload();
            }

        }

        function AddFavorites(strID) {
            var strFavorites;
            var FoundAt;
            var FavCount;

            AddingID = strID;

            strFavorites = txtFavs.value;
            FavCount = txtFavCount.value;
            if (FavCount == "" || FavCount == "NaN")
                FavCount = 0;
            FavCount = Number(FavCount) + 1;
            FoundAt = strFavorites.indexOf(strID + ",");
            if (!FoundAt > -1) {
                strFavorites = strFavorites + strID + ","
                txtFavCount.value = String(FavCount);
                txtFavs.value = strFavorites;
                //jsrsExecute("FavoritesRSupdate.asp", myCallback, "UpdateFavs", Array(strFavorites, String(FavCount), txtUser.value));
                ajaxurl = "FavoritesRSupdate.asp?CurrentUserID=" + txtUser.value + "&FavCount=" + String(FavCount) + "&Favorites=" + strFavorites;
                $.ajax({
                    url: ajaxurl,
                    type: "POST",
                    success: function (data) {
                        if (data == "1") {
                            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1&ID=" + AddingID);
                            RFLink.style.display = "";
                            AFLink.style.display = "none";
                        }
                    },
                    error: function (xhr, status, error) {
                        alert(error);
                    }

                });
            }
        }


        function myCallback(returnstring) {
            if (returnstring == "1") {
                window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1&ID=" + AddingID);
                RFLink.style.display = "";
                AFLink.style.display = "none";
            }
        }

        function ShowScorecard(DisplayedID) {
            var strID;

            strID = window.showModalDialog("<%= AppRoot %>/Deliverable/Scorecard/RootScorecard.asp?ID=" + DisplayedID, "", "dialogWidth:700px;dialogHeight:600px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
            //window.location.reload();
        }

        function ShowScorecardReport(DisplayedID) {
            var strID;

            window.open("<%= AppRoot %>/Deliverable/OTSCoreTeamDashboard.asp?RootID=" + DisplayedID, "_blank")
        }

        function ShowScorecardCoreTeam(CoreTeamID, ReportID) {
            var strID;

            window.open("<%= AppRoot %>/Deliverable/OTSCoreTeamDashboard.asp?CoreTeamID=" + CoreTeamID + "&Report=" + ReportID, "_blank")
        }


        function RemoveFavorites(strID) {
            var strFavorites;
            var FoundAt;
            var FavCount;

            AddingID = strID;

            strFavorites = txtFavs.value;
            FavCount = txtFavCount.value;
            if (FavCount == "" || FavCount == "NaN")
                FavCount = 0;
            FavCount = Number(FavCount) - 1;
            if (FavCount < 0)
                FavCount = 0;

            FoundAt = strFavorites.indexOf(strID + ",");
            if (!FoundAt > -1) {
                strFavorites = strFavorites.replace(strID + ",", "")
                txtFavCount.value = String(FavCount);
                txtFavs.value = strFavorites;
                //jsrsExecute("FavoritesRSupdate.asp", myCallback2, "UpdateFavs", Array(strFavorites, String(FavCount), txtUser.value));
                ajaxurl = "FavoritesRSupdate.asp?CurrentUserID=" + txtUser.value + "&FavCount=" + String(FavCount) + "&Favorites=" + strFavorites;
                $.ajax({
                    url: ajaxurl,
                    type: "POST",
                    success: function (data) {
                        if (data == "1") {
                            window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1");
                            RFLink.style.display = "none";
                            AFLink.style.display = "";
                        }
                    },
                    error: function (xhr, status, error) {
                        alert(error);
                    }

                });
            }
        }


        function myCallback2(returnstring) {
            if (returnstring == "1") {
                window.parent.frames("LeftWindow").location.replace("tree.asp?Prog=1");
                RFLink.style.display = "none";
                AFLink.style.display = "";
            }
        }

        function SetDMView(PageList, DelRootID, strClass) {
            window.location.href = "dmview.asp?Tab=" + PageList + "&ID=" + DelRootID + "&Class=" + strClass;
        }

    </script>

</head>
<body>
<div style="display:none"><a href="UpdateUserAccess.asp"></a></div>
<div id="HideForPulsar" style="<%= hideForPulsar%>">
    <div id="DMViewHeader">
        <h2>
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
	
	
	dim strID
	dim strDescription
	dim strNotes
	dim strManager
	dim strCategory
	dim strVendor
	dim strPart
	dim strType
	dim strGreenSpec
	dim strFilename
	dim strPathCell
	dim strLeadfree
	dim strDisplayedList
	dim InactiveCount
	dim blnPM
	dim blnAccessoryPM
	dim blnSysAdmin
	dim strTester
	dim strSpec
	dim strManagerName
	dim strBuildLevel
	dim strDeveloper
	dim strDevManagerId
	dim strDeliverableName
	dim strCoreTeamId
	
	blnSysAdmin = 0
	blnPM = 0
	blnAccessoryPM=0

	strDisplayedList = "Certification"
	
	InactiveCount = "0"
	
	strVendor = ""
	strFilename = ""
	strBuildLevel = ""
	strID = clng(DRID)
	
	if strID <> "" and isnumeric(strID) then
        
		dim CurrentUserSite
		dim strFavs
		dim strFavCount
		dim strDomainSite
		dim blnPreinstallGroup
		dim blnProcurementEngineer
        dim blnServiceCommodityManager
		dim blnLockReleases
		dim DevPMManagerID
		dim DevPMManagerName
		dim strTeamID
        dim blnShowOnStatus
        dim strProcurementGroup

        strProcurementGroup = "0"

        DevPMManagerID = 0
        DevPMManagerName = ""

		blnProcurementEngineer=false
        blnServiceCommodityManager = false
		blnPreinstallGroup=false
		
		regEx.Pattern = "[^0-9a-fA-F#]"
		on error resume next
		strTitleColor = regEx.Replace(Request.Cookies("TitleColor"), "")
		if strTitleColor = "" then
			strTitleColor = "#0000cd"
		end if
		on error goto 0

       if blnServiceCommodityManager then
            strProcurementGroup = "2"
       elseif blnProcurementEngineer then
            strProcurementGroup = "1"
       else
            strProcurementGroup = "0"
       end if

		if currentdomain = "americas" then
			CurrentUserSite = "1"
		else
			CurrentUserSite = "2"
		end if

        Dim strError, strWarning
        
		strSQL = "spGetDelPropSummary " & clng(strID)
		rs.Open strSQL,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strError = strError & "Unable to find the selected deliverable."
			strError = strError & "<BR><a id=RFLink style=""Display:none"" href=""javascript:RemoveFavorites(" & Server.HTMLEncode(DRID) & ");""><font face=verdana size=1>Remove From Favorites</font></a>"
			strError = strError & "<a id=AFLink style=""Display:none""><font face=verdana size=1></font></a>"
			strError = strError & "<font face=verdana size=2 id=LoadingMessage>Loading.  Please wait...</font>"
		else
			strDescription = server.HTMLEncode(rs("Description") & "") & "&nbsp;"
			strNotes = rs("Notes") & "&nbsp;"
			strManager = rs("Manager") & "&nbsp;"
			strDeveloper = rs("Developer") & "&nbsp;"
			strCategory = rs("Category") & "&nbsp;"
			blnLockReleases = rs("LockReleases")
			strVendor = rs("Vendor") & "&nbsp;"
			strPart = rs("BasePartNumber") & "&nbsp;"
			strTeamID = rs("TeamID") & ""
			strDevManagerID = rs("DevManagerID") 
			strTester = rs("Tester") & ""
			strType = rs("TypeID") & ""
			strSpec = trim(rs("DeliverableSpec") & "")
			strFilename = rs("Filename") & ""
			Response.Write rs("Name") & " Information"
			strDeliverableName = rs("Name") & ""
            strCoreTeamID = trim(rs("CoreTeamID") & "")
            blnShowOnStatus = rs("ShowOnStatus")
            DeliverableCategoryId = rs("CategoryId") & ""
			if rs("Active")=0 then
				Response.write "<BR><BR><font size=1 face=verdana color=red>This deliverable is inactive.</font><BR>"
			end if
		rs.Close
		if left(strSpec,2) = "\\" then
			strSpec = "<a href=""file://" & strSpec & """>" & strSpec & "</a>"
		end if


        'Lookup Manager ID and name
        strManagerName = ""
        if isnumeric(strDevManagerID) then
            rs.open "spGetManagerInfo " & clng(strDevManagerID),cn,adOpenForwardOnly
            if not (rs.eof and rs.bof) then
                DevPMManagerID  = rs("ID")
                DevPMManagerName = rs("Name")
                strManagerName = " or " & longname(DevPMManagerName)
            end if
            rs.close
        end if

            %>
        </h2>
        <div id="FavoritesLinks" style="margin-bottom:40px;">
        <span id="EditRootDeliverableLink" class="Link" onclick="ShowProperties(<%=DRID%>,'<%=Server.HTMLEncode(sTab)%>')">Edit Root Deliverable</span> 
        <span id="RemoveFavorites">| <span id="RemoveFavoritesLink" class="Link" onclick="RemoveFavorites(<%=DRID%>)">Remove From Favorites</span></span>
        <span id="AddFavorites">| <span id="AddFavoritesLink" class="Link" onclick="AddFavorites(<%=DRID%>)"><font face="verdana" size="1">Add To Favorites</span></span> 
        <span id="CloneRoot">| <span id="CloneRootLink" class="Link" onclick="CloneRoot(<%=DRID%>)">Clone This Root</span></span>
        </div>
        <input type="hidden" id="txtTypeID" name="txtTypeID" value="<%=strType%>" />
        <input type="hidden" id="txtFilename" name="txtFilename" value="<%=strFilename%>" />
        <table cellspacing="1" cellpadding="1" width="100%" border="1" bordercolor="tan"
            bgcolor="ivory">
            <tr>
                <td nowrap width="100" bgcolor="cornsilk">
                    <strong><font size="1">Manager/OTS&nbsp;PM:</font></strong>
                </td>
                <td>
                    <font size="1">
                        <%=strManager%></font>
                </td>
                <td bgcolor="cornsilk">
                    <strong><font size="1">Vendor:</font></strong>
                </td>
                <td>
                    <font size="1">
                        <%=strVendor%></font>
                </td>
            </tr>
            <%if strFilename = "HFCN" then%>
            <tr style="display: none">
                <%else%>
                <tr>
                    <%end if%>
                    <%if trim(strType) = "1" then%>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Execution&nbsp;Engineer:</font></strong>
                    </td>
                    <%else%>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Tester:</font></strong>
                    </td>
                    <%end if%>
                    <td>
                        <font size="1">
                            <%=strTester%>&nbsp;</font>
                    </td>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Category:</font></strong>
                    </td>
                    <td>
                        <font size="1">
                            <%=strcategory%></font>
                    </td>
                </tr>
                <tr>
                    <%if trim(strType) = "1" then%>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Development&nbsp;Engineer:</font></strong>
                    </td>
                    <%else%>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Developer:</font></strong>
                    </td>
                    <%end if%>
                    <%if trim(strType) = "1" then%>
                    <td colspan="3">
                        <font size="1">
                            <%=strDeveloper%>&nbsp;</font>
                    </td>
                    <%else%>
                    <td>
                        <font size="1">
                            <%=strDeveloper%>&nbsp;</font>
                    </td>
                    <td nowrap width="100" bgcolor="cornsilk">
                        <strong><font size="1">Root&nbsp;Filename:</font></strong>
                    </td>
                    <td>
                        <font size="1">
                            <%=strFilename%></font>
                    </td>
                    <%end if%>
                </tr>
                <%if trim(strSpec) <> "" then%>
                <tr>
                    <td bgcolor="cornsilk">
                        <strong><font size="1">Functional Spec:</font></strong>
                    </td>
                    <td colspan="4">
                        <font size="1">
                            <%=strSpec%></font>
                    </td>
                </tr>
                <%end if%>
                <tr>
                    <td bgcolor="cornsilk" valign="top">
                        <strong><font size="1">Description:</font></strong>
                    </td>
                    <td colspan="4">
                        <font size="1">
                            <%=replace(strDescription,vbcrlf,"<BR>")%></font>
                    </td>
                </tr>
                <%if strFilename = "HFCN" then%>
                <tr style="display: none">
                    <%else%>
                    <tr>
                        <%end if%>
                        <td bgcolor="cornsilk">
                            <strong><font size="1">Notes:</font></strong>
                        </td>
                        <td colspan="4">
                            <font size="1">
                                <%=strNotes%></font>
                        </td>
                    </tr>
                    <tr>
                        <td bgcolor="cornsilk" valign="top">
                            <strong><font size="1">Scorecard:</font></strong>
                        </td>
                        <td colspan="4">
                            <font face="verdana" size="1" color="black"></font><a href="javascript:ShowScorecard(<%=DRID%>)">
                                <font face="verdana" size="1">Edit&nbsp;Scorecard</font></a> | <a href="javascript:ShowScorecardReport(<%=DRID%>)">
                                    <font face="verdana" size="1">Deliverable&nbsp;Report</font></a>
                            <%if trim(strCoreTeamID) <> "" and trim(strCoreTeamID) <> "0" and blnShowOnStatus then%>
                            | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,0)"><font face="verdana"
                                size="1">Core&nbsp;Team&nbsp;Report</font></a> | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,1)">
                                    <font face="verdana" size="1">Executive&nbsp;Summary</font></a> | <a href="javascript:ShowScorecardCoreTeam(<%=clng(strCoreTeamID)%>,2)">
                                        <font face="verdana" size="1">Action&nbsp;Items</font></a>
                            <%end if%>
                        </td>
                    </tr>
        </table>
        <br>
        <table border="1" bordercolor="Ivory" cellspacing="0" cellpadding="2" id="menubar"
            class="MenuBar">
            <tr bgcolor="<%=strTitleColor%>">
                <%if strDisplayedList = "Versions" or strDisplayedList = ""  then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Version List&nbsp;&nbsp;</font>
                </td>
                <%else%>
                <td>
                    <font size="1" face="verdana"><a href="javascript:SetDMView('Versions', '<%= strID%>', '<%= sClass%>');">
                        &nbsp;&nbsp;Version List</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "OTS" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Observations&nbsp;&nbsp;</font>
                </td>
                <%else%>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('OTS', '<%= strID%>', '<%= sClass%>');">Observations</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "Agency" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Agency&nbsp;&nbsp;</font>
                </td>
                <%elseif trim(strType) = "1" then %>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Agency', '<%= strID%>', '<%= sClass%>');">Agency</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "Certification" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Certification&nbsp;<span style="color:Red; font-weight:bold;">(Beta)</span>&nbsp;&nbsp;</font>
                </td>
                <%elseif trim(strType) = "1" then %>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Certification', '<%= strID%>', '<%= sClass%>');">Certification&nbsp;<span style="color:Red; font-weight:bold;">(Beta)</span></a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "Restriction" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Restrictions&nbsp;&nbsp;</font>
                </td>
                <%elseif trim(strType) = "1" then %>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Restriction', '<%= strID%>', '<%= sClass%>');">Restrictions</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "SMR" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;SMR&nbsp;&nbsp;</font>
                </td>
                <%else%>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('SMR', '<%= strID%>', '<%= sClass%>');">SMR</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "Documents" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Documents&nbsp;&nbsp;</font>
                </td>
                <%elseif trim(strTeamID) = "3" then%>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Documents', '<%= strID%>', '<%= sClass%>');">Documents</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
                <%if strDisplayedList = "Products" then%>
                <td class="ButtonSelected">
                    <font size="1" face="verdana">&nbsp;&nbsp;Products&nbsp;&nbsp;</font>
                </td>
                <%else%>
                <td>
                    <font size="1" face="verdana">&nbsp;&nbsp;<a href="javascript:SetDMView('Products', '<%= strID%>', '<%= sClass%>');">Products</a>&nbsp;&nbsp;</font>
                </td>
                <%end if%>
            </tr>
        </table>
        <%End If %>
    </div>
  </div>
    <br />
    <div id="ToolMenu" style="display:none;">
        <span class="Title">Tools:&nbsp;</span> <span id="initWorkflow" class="Link">Init Workflow</span>
        <span id="showFrame" class="Link Hidden">Show Frame</span>
        <span class="seperator">|</span><span id="selectFollowers" class="Link">Choose Followers</span>
        <span id="testToolLink" class="ui-helper-hidden"><span class="seperator">|</span><span id="testHarness" class="Link">Test</span></span>
    </div>
    <div id="loading">
        <img src="../images/loading24.gif" alt="Loading" />
        Loading...</div>
    <div id="body">
        <div class="ui-widget ui-helper-clearfix">
            <div id="error" class="ui-state-error ui-corner-all" style="padding: 0 .7em;">
                <p>
                    <span class="ui-icon ui-icon-alert" style="float: left; margin-right: .3em;"></span>
                    <strong>Alert:&nbsp;</strong><span id="errorText"><%= strError & ""%></span></p>
            </div>
            <div id="warning" class="ui-state-highlight ui-corner-all" style="padding: 0 .7em;">
                <p>
                    <span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
                    <strong>Notice:&nbsp;</strong><span id="warningText"><%= strWarning & ""%></span></p>
            </div>
        </div>
        <%
    Dim lastWorkflow, lastWorkflowStep, isFirstPass, isFirstStep, AgencyTypeId, lastVersion
    isFirstPass = True
    Set cmd = dw.CreateCommandSP(cn, "usp_AgencySelectStatus")
    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 0, ""
    dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 0, DRID
    dw.CreateParameter cmd, "@p_DeliverableVersionID", adInteger, adParamInput, 8, ""
    Set rs = dw.ExecuteCommandReturnRS(cmd)
    rs.Sort="AgencyType, DeliverableRootId, DeliverableVersionId, Status"


    Do Until rs.EOF
        If lastVersion <> rs("DeliverableVersionId")&"" Or lastWorkflow <> rs("AgencyType")&"" Then
            lastVersion = rs("DeliverableVersionId") & ""
            lastWorkflow = rs("AgencyType")&""
            lastWorkflowStep = ""

            If NOT isFirstPass Then
                isFirstPass = false
                Response.Write "</tbody></table><br /><br /></div>"
            End If
            Response.Write "<div class=""CertificationGroup""><h3>" & rs("DeliverableRootName") & "&nbsp;-&nbsp;" & rs("Version") & "&nbsp;-&nbsp;" & rs("AgencyType") 
            If CLng(rs("LeveragedStatusId")) = 0 Then Response.Write " - Lead"
            Response.Write "&nbsp;&nbsp;"
            Response.Write "<span class=""Link initWorkflow"">Init Workflow</span>"
            If CLng(rs("LeveragedStatusId")) = 0 Then Response.Write "<span class=""Link selectFollowers"">Choose Followers</span>"
            Response.Write "</h3>"
            
            Response.Write "<input type='hidden' class='deliverableVersionId' value='" & rs("DeliverableVersionId") & "' />"
            Response.Write "<input type='hidden' class='agencyTypeId' value='" & rs("AgencyTypeId") & "' />"
            
            Response.Write "<table class=""StatusTable"
            If  CLng(rs("LeveragedStatusId")) = 0 Then Response.Write " LeadDeliverable"
            Response.Write """>"
            isFirstStep = True
            isFirstPass = false            
        End If

        If lastWorkflowStep <> rs("Status")&"" Then
            lastWorkflowStep = rs("Status") & ""
            If NOT isFirstStep Then
                Response.Write "</tbody>"
            End If
            Response.Write "<thead><tr><th colspan=""4"">" & rs("Status") & "</th></tr></thead><tbody class=""" & Replace(rs("Status"), " ", "_") & """><tr><th>Name</th><th>Current Step</th><th>Days till Due</th><th>Certification Target Dt.</th></tr>"
            isFirstStep = false
        End If

        If rs("AgencyStatusId") & "" <> "" Then
            Response.Write "<tr><td><input type=""hidden"" class=""statusId"" value=""" & rs("AgencyStatusId") & """><input type=""hidden"" class=""nextStepId"" value=""" & rs("NextStepId") & """>" & rs("AgencyName") & "</td><td>" & rs("StepName") & "</td><td>" & rs("DaysToTarget") & "</td><td>" & rs("TargetDate") & "</td></tr>"
        End If

        
        rs.MoveNext
    Loop
    
    If Not isFirstPass Then
        Response.Write "</tbody></table>"
    End If
    
    rs.close
End If    
        %>
        <div id="iframeDialog" title="Coolbeans">
            <iframe frameborder="0"></iframe>
        </div>
        <!-- #include file = "AgencyDetailsDialog.htm" -->
        <!-- #include file = "AgencyInitWorflowDialog.htm" -->
        <input type="hidden" id="currentUser" value="<%=currentUserName %>" />
        <input type="hidden" id="deliverableRootId" value="<%=DRID %>" />
        <input type="hidden" id="deliverableCategoryId" value="<%=DeliverableCategoryId %>" />
        <input type="hidden" id="agencyTypeId" value="<%= AgencyTypeId %>" />
        <input type="hidden" id="isExcaliburAdmin" value="<%= isSysAdmin %>" />
        <input type="hidden" id="isAgencyEditModeOn" value="<%= isEditModeOn %>" />
    </div>
</body>
</html>
