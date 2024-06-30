<%@  language="VBScript" CODEPAGE=65001 %>
<%
	Response.Charset="UTF-8"
	if request("cboFormat")= "Excel" and request("txtReportSections") <> "3" then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition","attachment; filename=observation.xls"
	elseif request("cboFormat")= "Word" and request("txtReportSections") <> "3" then
		Response.ContentType = "application/msword"
		Response.AddHeader "content-disposition","attachment; filename=observation.doc"
	elseif ucase(request("cboFormat"))="XML" and request("txtReportSections") = "0" then
		response.ContentType="text/xml"
	else
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
	end if

	'SECTIONS
	'-2 - History Report
	'-1 - SQL Syntax checker
	' 0 - Summary Report
	' 1 - Detailed Report
	' 2 - Working Notes Reports
	' 3 - Email Report
	' 4 - Standard Backlong Chart
	' 5 - Table of Counts by Priority
	' 6 - Table of Counts by Developer
	' 7 - Affected Product List
	' 8 - Table of Counts by Deviverable
	' 9 - Table of Counts by SubSystem
	'10 - Table of Counts by State
	'11 - Table of Counts by Core Team
	'12 - Table of Counts by Component PM
	'13 - Table of Counts by Status
	'28 - Table of Counts by Owner

'	if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'		response.Write "<pre>"
'		response.Write request.Form
'		response.Write "</pre>"
'		response.End
'	end if
	if not( ucase(request("cboFormat"))="XML" and request("txtReportSections") = "0") then
%>
<html>
<head>
	<title>Observation Reports</title>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
		//<!-- #include file = "../../_ScriptLibrary/sort.js" -->
		var MouseX, MouseY;
		function SaveMouseCoordinates() {
			MouseX = event.clientX;
			MouseY = event.clientY;
		}

		function ShowIDMenu(strObservation, strPartner) {
			var popupBody;

			popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

			popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
			popupBody = popupBody + "<font face=Arial size=2>";
			popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayDetails(" + strObservation + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Details</SPAN></font></DIV>";

			popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
			popupBody = popupBody + "<font face=Arial size=2>";
			popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayUpdates(" + strObservation + ")'\" >&nbsp;&nbsp;&nbsp;View&nbsp;Updates</SPAN></font></DIV>";

			popupBody = popupBody + "<DIV>";
			popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

			popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
			popupBody = popupBody + "<font face=Arial size=2>";
			popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:OpenSI(" + strObservation + "," + strPartner + ")'\" >&nbsp;&nbsp;&nbsp;Open&nbsp;Sudden&nbsp;Impact&nbsp;</SPAN></font></DIV>";

			popupBody = popupBody + "</DIV>";

			var NewHeight;
			var NewWidth;

			mnuPopup.style.display = "";
			mnuPopup.innerHTML = popupBody;

			mnuPopup.style.width = mnuPopup.scrollWidth + 10;
			mnuPopup.style.height = mnuPopup.scrollHeight;
			mnuPopup.style.left = document.body.scrollLeft + MouseX;
			mnuPopup.style.top = document.body.scrollTop + MouseY;
		}

		function OpenSI(strID, PartnerID) {
			if (PartnerID == "1")
				window.open("https://si.austin.hp.com/si/?ObjectType=6&Object=" + strID);
			else
				window.open("https://prp-si.corp.hp.com/si/?ObjectType=6&Object=" + strID);

		}

		function DisplayDetails(strID) {
//			window.open("report_mattH.asp?txtReportSections=1&txtObservationID=" + strID);
			window.open("report.asp?txtReportSections=1&txtObservationID=" + strID);
		}

		function DisplayUpdates(strID) {
//			window.open("report_mattH.asp?txtReportSections=2&txtObservationID=" + strID);
			window.open("report.asp?txtReportSections=2&txtObservationID=" + strID);
		}

		function window_onmouseup() {
			if (typeof (mnuPopup) != "undefined")
				mnuPopup.style.display = "none";
		}

		function SendEmail() {
			var blnChecked = false;

			if (typeof (frmEmail.chkTo.length) == "undefined")
				blnChecked = frmEmail.chkTo.checked;
			else {
				for (i = 0; i < frmEmail.chkTo.length; i++)
					if (frmEmail.chkTo[i].checked)
						blnChecked = true;
			}
			if (!blnChecked) {
				alert("You must select at least one person to send an email.");
			}
			else if (frmEmail.txtSubject.value == "") {
				alert("Subject is required.");
				frmEmail.txtSubject.focus();
			}
			else if (frmEmail.txtNotes.value == "") {
				alert("Notes are required.");
				frmEmail.txtNotes.focus();
			}
			else {
				var i = 1;

				while (document.all("txtEmailTable" + i) != null) {
					document.all("txtEmailTable" + i).value = document.all("tblOTS" + i).outerHTML;
					i = i + 1;
				}
				frmEmail.submit();
			}
		}

		function window_onload() {
			if (typeof (txtClosedCount) != "undefined")
				if (txtClosedCount.value != "0" && txtClosedCount.value != "")
					ClosedWarningRow.style.display = "";
		}

		function ShowAffectedWindow(strID) {
			var MyLeft = window.screenLeft + (document.body.offsetWidth / 2) - 350;
			var MyTop = window.screenTop + (document.body.offsetHeight / 2) - 150;
//			window.open("Report_mattH.asp?txtReportSections=7&txtObservationID=" + strID, "_blank", "width=700, height=200, resizable=yes, scrollbars=yes,top=" + MyTop + ",left=" + MyLeft);
			window.open("Report.asp?txtReportSections=7&txtObservationID=" + strID, "_blank", "width=700, height=200, resizable=yes, scrollbars=yes,top=" + MyTop + ",left=" + MyLeft);
		}
//-->
	</script>
	<script runat="server" type="text/javascript" language="javascript">
		function getTimezoneOffset() {
			var d = new Date();
			return d.getTimezoneOffset();
		}
	</script>
	<style type="text/css">
		td
		{
			font-family: Verdana;
			font-size: xx-small;
			vertical-align: top;
		}
		thead
		{
			background-color: #f5f5dc; /*beige*/
			font-family: Verdana;
			font-size: xx-small;
		}
		thead.statusHeader
		{
			background-color: white;
			font-family: Verdana;
			font-size: xx-small;
		}
		h1
		{
			font-family: Verdana;
			font-size: x-small;
		}
		body
		{
			font-family: Verdana;
			font-size: xx-small;
		}
		A:link
		{
			color: Blue;
		}
		A:visited
		{
			color: Blue;
		}
		A:hover
		{
			color: red;
		}
		h1
		{
			font-family: Verdana;
			font-size: small;
		}
	</style>
</head>
<body onmouseup="window_onmouseup()" onload="window_onload();">
	<%
	end if
	dim ProfileData
	dim ProfileReportID
	dim PageSections
	dim XMLFormat

	if ucase(request("cboFormat"))="XML" and request("txtReportSections") = "0" then
		if request("XMLFormat") = "2" then
			XMLFormat = 2
		else
			XMLFormat = 1
		end if
	else
		XMLFormat=0
	end if

	dim strNoticeTable,BulletinHeaderColors,BulletinBodyColors
	Dim WeeksParam, EndDateParam, LegendParam, blnActivityGraphParam, blnTotalBacklogParam, strTitleParam, Widthparam, HeightParam, ItemsPerGridLineParam, strChartType

	dim cnExcalibur, cnSIO, rs, cm
	set cnExcalibur = server.CreateObject("ADODB.Connection")
	set cnSIO = server.CreateObject("ADODB.Connection")
	cnExcalibur.ConnectionString = Session("PDPIMS_ConnectionString")
	cnSIO.ConnectionString = "Provider=SQLOLEDB.1;Data Source=housireport01.auth.hpicorp.net;Initial Catalog=sio;User ID=Excalibur_RO;Password=sQ8be9AyqPQKEcqsa3mE;"
	on error resume next
	cnSIO.Open
	on error goto 0
    if cnSIO.errors.count > 0 then
        Response.Redirect "offline.asp"
    end if
	cnExcalibur.Open
	set rs = server.CreateObject("ADODB.recordset")
	cnExcalibur.CommandTimeout = 120 '90 '50 '180
	cnSIO.CommandTimeout = 120 '90 '50 '180

	strNoticeTable = ""
	BulletinHeaderColors = split("SeaGreen,Firebrick,Gold,SeaGreen",",")
	BulletinBodyColors = split("Honeydew,MistyRose,LightYellow,Honeydew",",")
	rs.Open "Select * FROM Bulletins with (NOLOCK) where active=1 and OTS=1 Order By id;",cnExcalibur,adOpenForwardOnly
	do while not rs.eof
		strNoticeTable = strNoticeTable & "<table cellSpacing=0 cellPadding=2 width=""100%"" border=0 bordercolor=black>"
			if rs("Severity") > -1 and rs("Severity") < 4 then
				strNoticeTable = strNoticeTable & "<TR><TD bgcolor=" & BulletinHeaderColors(rs("Severity")) & "><strong><font size=2 face=Verdana color=white>" & rs("Subject") & "</font></strong></TD></TR>"
				strNoticeTable = strNoticeTable & "<TR><TD bgcolor=" & BulletinBodyColors(rs("Severity"))& "><font size=1 color=black face=verdana>" & rs("Body") & "<BR><BR></b></font></TD></TR>"
			else
				strNoticeTable = strNoticeTable & "<TR><TD bgcolor=" & BulletinHeaderColors(0) & "><strong><font size=2 face=Verdana color=white>" & rs("Subject") & "</font></strong></TD></TR>"
				strNoticeTable = strNoticeTable & "<TR><TD bgcolor=" & BulletinBodyColors(0)& "><font size=1 color=black face=verdana>" & rs("Body") & "<BR><BR></b></font></TD></TR>"
			end if
		strNoticeTable = strNoticeTable & "</TABLE><BR>"
		rs.movenext
	loop
	rs.close

	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserName
	dim CurrentUserEmail
	dim CurrentUserGroup
	dim CurrentUserDivision
	dim CurrentUserPartner
	dim strColumnHeaders
	dim Section5Title
	dim blnMobileODMUser
	dim CurrentUserPartnerName
	dim CurrentUserOtherPartnerNames
	dim OtherPartnerNameArray

	blnMobileODMUser = false

	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cnExcalibur
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
		CurrentUserID = rs("ID")
		CurrentUserName = rs("Name") & ""
		CurrentUserEmail = rs("Email") & ""
		CurrentUserGroup = rs("WorkgroupID") & ""
		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = rs("PartnerID") & ""
		CurrentUserOtherPartnerNames = rs("OtherPartnerNames") & ""
	else
		Response.Redirect "../../Excalibur.asp"
	end if
	rs.Close

	if trim(Currentuserpartner) <> "1" and trim(CurrentUserDivision) = "1" then
		rs.open "spGetPartnerType " & trim(Currentuserpartner),cnExcalibur
		if not (rs.eof and rs.bof) then
			if trim(rs("partnerTypeID") & "") = "1" or CurrentUserPartner = 80 then
				blnMobileODMUser = true
			end if
		end if
		rs.Close
	end if

	dim strLimitPartner
	strLimitPartner = ""

	if CurrentUserPartner = 1 then
		CurrentUserPartnerName = "HP"
	elseif trim(CurrentUserPartner) = "9" then
		Response.Redirect "../../mobilese/modusmain.asp"
	else
		rs.Open "spGetPartnerName " & CurrentUserPartner,cnExcalibur,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			CurrentUserPartnerName = ""
		else
			CurrentUserPartnerName = rs("Name") & ""
		end if
		rs.Close

		strLimitPartner = " ("
'		strLimitPartner = strLimitPartner & " o.ownergroup like '%" & currentuserpartnername & "%'"
'		strLimitPartner = strLimitPartner & " or o.originatorgroup like '%" & currentuserpartnername & "%'"
'		strLimitPartner = strLimitPartner & " or o.developergroup like '%" & currentuserpartnername & "%'"
'		strLimitPartner = strLimitPartner & " or o.ComponentPMgroup like '%" & currentuserpartnername & "%' "
		strLimitPartner = strLimitPartner & " olv.odms like '%" & currentuserpartnername & "%'"
		strLimitPartner = strLimitPartner & " or olv.suppliers like '%" & currentuserpartnername & "%' "

		if trim(CurrentUserOtherPartnerNames) <> "" then
			dim strpartnername
			OtherPartnerNameArray = split(CurrentUserOtherPartnerNames,",")

			for each strpartnername in OtherPartnerNameArray
				strpartnername = trim(strpartnername)
				if strpartnername <> "" then
'					strLimitPartner = strLimitPartner & " or o.ownergroup like '%" & strpartnername & "%'"
'					strLimitPartner = strLimitPartner & " or o.originatorgroup like '%" & strpartnername & "%'"
'					strLimitPartner = strLimitPartner & " or o.developergroup like '%" & strpartnername & "%'"
'					strLimitPartner = strLimitPartner & " or o.ComponentPMgroup like '%" & strpartnername & "%' "
					strLimitPartner = strLimitPartner & " or olv.odms like '%" & strpartnername & "%'"
					strLimitPartner = strLimitPartner & " or olv.suppliers like '%" & strpartnername & "%' "
				end if
			next
		end if

		strLimitPartner = strLimitPartner & " ) "
	end if

	ProfileData = ""
	ProfileReportID = ""
	if request("ProfileID") <> "" then
		rs.open "spGetReportProfile " & clng(request("ProfileID")),cnExcalibur
		if not (rs.eof and rs.bof) then
			ProfileData = rs("SelectedFilters") & ""
			ProfileReportID = rs("TodayPageLink") & ""
		end if
		rs.Close
	end if

	dim strPageParams

	strPageParams = ""

	if trim(ProfileReportID) <> "" or left(trim(request("txtReportSections")),1) = "#" then
		PageSections = ""
		if left(trim(request("txtReportSections")),1) = "#" then
			ProfileReportID = mid(trim(request("txtReportSections")),2)
		end if
		if trim(ProfileReportID) = "5" then
			PageSections = "5,6,8,4"
		elseif trim(ProfileReportID) = "6" then
			PageSections = "5,9,11,12,10,13,4"
		else
			rs.open "spGetProfileReportDefinition " & clng(ProfileReportID),cnExcalibur
			if not (rs.eof and rs.bof) then
				PageSections=trim(rs("ReportSections") & "")
				strPageParams=trim(rs("SectionParameters") & "")
			end if
			rs.Close
		end if
		if trim(PageSections) = "" then
			PageSections = "0"
		end if
	elseif request("txtReportSections") <> "" then
		PageSections=request("txtReportSections")
		strPageParams = request("txtReportSectionParameters")
	else
		PageSections="0"
		strPageParams = ""
	end if

	if PageSections = "" or PageSections = "-1" or PageSections = "0" or PageSections = "1" or PageSections = "2" or PageSections = "3" or PageSections = "7" then
		blnStatusReport = false
	else
		blnStatusReport = true
	end if

	dim strSavedAffectedProduct,strSavedAffectedState,strSavedProduct,strSavedProductAndVersion,strSavedProductFamily,strSavedOwner,strSavedOwnerGroup,strSavedOriginatorGroup, strSavedDeveloperGroup
	dim strSavedProductPMGroup,strSavedTesterGroup,strSavedComponentTestLeadGroup,strSavedProductTestLeadGroup,strSavedApproverGroup,strSavedComponentPMGroup,strSavedCoreTeam
	dim strSavedTitle,strSavedComponent,strSavedState,strSavedSortColumn1,strSavedSortColumn2,strSavedSortColumn3,strSavedSubsystem,strSavedGatingMilestone,strSavedDeveloper
	dim strSavedComponentPM,strSavedProductPM,strSavedOriginator,strSavedTester,strSavedComponentTestLead,strSavedProductTestLead,strSavedApprover,strSavedObservationID
	dim strSavedTargetDateCompare,strSavedDaysOpenCompare,strSavedDaysStateCompare,strSavedDaysOwnerCompare,strSavedPriority,strSavedLargeFields,strSavedType,strSavedStatus
	dim strSavedSubject,strSavedNotes,strSavedAdvanced,strSaveColumns,strSavedDateOpenedCompare,strSavedDateClosedCompare,strSavedDateModifiedCompare, strSavedAssigned
	dim strSavedFormat,strSavedSeverity,strSavedEscape, strSavedImpact,strSavedTextSearch
	dim strSaveFrequency, strSavedFeature,strSavedProductGroup

	dim strSaveSearchDetails, strSaveSearchSummary,strSaveSearchImpact,strSaveSearchReproduce,strSaveSearchHistory,strSaveSearchType, strSavedDaysOpenDays, strSavedDaysOpenRange
	dim strSavedDaysStateDays, strSavedDaysStateRange, strSavedDaysOwnerDays, strSavedDaysOwnerRange, strSavedDateOpenedDays,strSavedDateDateOpenedRange1
	dim strSavedDateDateOpenedRange2,strSavedDateClosedDays,strSavedDateClosedRange1,strSavedDateClosedRange2,strSavedDateModifiedDays,strSavedDateModifiedRange1
	dim strSavedDateModifiedRange2,strSavedTargetDateDays,strSavedTargetDateRange1,strSavedTargetDateRange2, strSavedSort1Direction, strSavedSort2Direction, strSavedSort3Direction

	strSaveSearchHistory = getvalue("chkSearchHistory")
	strSaveSearchReproduce = getvalue("chkSearchReproduce")
	strSaveSearchImpact = getvalue("chkSearchImpact")
	strSaveSearchSummary = getvalue("chkSearchSummary")
	strSaveSearchDetails = getvalue("chkSearchDetails")
	strSavedAffectedProduct = getvalue("lstAffectedProduct")
	strSaveSearchType = getvalue("cboSearchType")
	strSavedAffectedState = getvalue("lstAffectedState")
	strSavedProduct = getvalue("lstProduct")
	strSavedProductAndVersion = getvalue("lstProductAndVersion")
	strSavedProductFamily = getvalue("lstProductFamily")
	strSavedOwner = getvalue("lstOwner")
	strSavedOwnerGroup = getvalue("lstOwnerGroup")
	strSavedOriginatorGroup = getvalue("lstOriginatorGroup")
	strSavedDeveloperGroup = getvalue("lstDeveloperGroup")
	strSavedProductPMGroup = getvalue("lstProductPMGroup")
	strSavedTesterGroup = getvalue("lstTesterGroup")
	strSavedComponentTestLeadGroup = getvalue("lstComponentTestLeadGroup")
	strSavedProductTestLeadGroup = getvalue("lstProductTestLeadGroup")
	strSavedApproverGroup= getvalue("lstApproverGroup")
	strSavedComponentPMGroup= getvalue("lstComponentPMGroup")
	strSavedCoreTeam = getvalue("lstCoreTeam")
	strSavedState = getvalue("lstState")
	strSavedSortColumn1 = getvalue("cboSortColumn1")
	strSavedSort1Direction = getvalue("cboSort1Direction")
	strSavedSortColumn2 = getvalue("cboSortColumn2")
	strSavedSort2Direction = getvalue("cboSort2Direction")
	strSavedSortColumn3 = getvalue("cboSortColumn3")
	strSavedSort3Direction = getvalue("cboSort3Direction")
	strSavedSubsystem = getvalue("lstSubsystem")
	strSavedGatingMilestone = getvalue("lstGatingMilestone")
	strSavedDeveloper = getvalue("lstDeveloper")
	strSavedAssigned = getvalue("lstAssigned")
	strSavedComponentPM = getvalue("lstComponentPM")
	strSavedProductPM = getvalue("lstProductPM")
	strSavedOriginator = getvalue("lstOriginator")
	strSavedTester = getvalue("lstTester")
	strSavedComponentTestLead = getvalue("lstComponentTestLead")
	strSavedProductTestLead = getvalue("lstProductTestLead")
	strSavedApprover = getvalue("lstApprover")
	strSavedObservationID = replace(lcase(getvalue("txtObservationID")), "sio", "")
	strSavedTitle = getvalue("txtTitle")
	strSavedSubject = getvalue("txtSubject")
	strSavedNotes = getvalue("txtNotes")
	strSavedAdvanced = getvalue("txtAdvanced")
	strSaveColumns = getvalue("lstColumns")
	strSavedDateOpenedCompare = getvalue("cboDateOpenedCompare")
	strSavedDateClosedCompare = getvalue("cboDateClosedCompare")
	strSavedDateModifiedCompare = getvalue("cboDateModifiedCompare")
	strSavedTargetDateCompare = getvalue("cboTargetDateCompare")
	strSavedDaysOpenCompare = getvalue("cboDaysOpenCompare")
	strSavedDaysStateCompare = getvalue("cboDaysStateCompare")
	strSavedDaysOwnerCompare = getvalue("cboDaysOwnerCompare")
	strSavedDaysOpenDays = getvalue("txtDaysOpenDays")
	strSavedDaysOpenRange = getvalue("txtDaysOpenRange")
	strSavedDaysStateDays = getvalue("txtDaysStateDays")
	strSavedDaysStateRange = getvalue("txtDaysStateRange")
	strSavedDaysOwnerDays = getvalue("txtDaysOwnerDays")
	strSavedDaysOwnerRange = getvalue("txtDaysOwnerRange")
	strSavedDateOpenedDays = getvalue("txtDateOpenedDays")
	strSavedDateDateOpenedRange1 = getvalue("txtDateOpenedRange1")
	strSavedDateDateOpenedRange2 = getvalue("txtDateOpenedRange2")
	strSavedDateClosedDays = getvalue("txtDateClosedDays")
	strSavedDateClosedRange1 = getvalue("txtDateClosedRange1")
	strSavedDateClosedRange2 = getvalue("txtDateClosedRange2")
	strSavedDateModifiedDays = getvalue("txtDateModifiedDays")
	strSavedDateModifiedRange1 = getvalue("txtDateModifiedRange1")
	strSavedDateModifiedRange2 = getvalue("txtDateModifiedRange2")
	strSavedTargetDateDays = getvalue("txtTargetDateDays")
	strSavedTargetDateRange1 = getvalue("txtTargetDateRange1")
	strSavedTargetDateRange2 = getvalue("txtTargetDateRange2")

	strSavedPriority = getvalue("chkPriority")
	strSavedLargeFields = getvalue("txtLargeFieldLimit")

	strSavedType = getvalue("lstType")
	strSavedStatus = getvalue("cboStatus")
	strSavedFormat = getvalue("cboFormat") & ""
	strSavedSeverity = getvalue("cboSeverity")
	strSavedEscape = getvalue("cboEscape")
	strSavedDivision = getvalue("cboDivision")
	strSavedImpact = getvalue("cboImpact")
	strSavedTextSearch = getvalue("txtSearch")
	strSaveFrequency = getvalue("lstFrequency")
	strSavedComponent = getvalue("lstComponent")
	strSavedFeature = getvalue("lstFeature")
	strSavedProductGroup = getvalue("lstProductGroup")

	if xmlformat = 0 then
		if trim(request("txtReportSections")) <> "-1" then
			if strNoticeTable <> "" and strSavedFormat <> "Excel" and strSavedFormat <> "Word" and strSavedFormat <> "XML" then
				response.write strNoticeTable
			end if
			if strSavedTitle = "" then
				if trim(request("txtReportSections")) = "3" then
					response.write "<h1>Email Observation Owners</h1>"
				elseif trim(request("txtReportSections")) <> "7" then
					if blnStatusReport then
						response.write "<h1 align=center>Observation Report</h1>"
					elseif not (strPageParams = "macro" and request("cboFormat") = "Excel") then
						response.write "<h1>Observation Report</h1>"
					end if
				end if
			else
				if not (strPageParams = "macro" and request("cboFormat") = "Excel") then
					response.write "<h1>" & server.HTMLEncode(strSavedTitle) & "</h1>"
				end if
			end if
		end if
	end if
	dim ColumnArray

	if (PageSections = "0" and strPageParams = "macro") then
'		strSaveColumns = "Observation ID,Status,Reviewed,Priority,Severity,Frequency,State,Short Description,Long Description,Steps to Reproduce,Customer Impact,Updates2,Component Type,Sub System,Component,Component Version,Localization,Component PartNo,Earliest Product Milestone,Gating Milestone,Owner,Owner Group,Owner Email,Owner Location,Owner Manager,Originator,Originator Group,Originator Email,Originator Location,Originator Manager,Product Segment,Product Family,Primary Product,Affected Product,Product PM,Product PM Group,Product PM Email,Product PM Location,Product PM Manager,Product Test Lead,Product Test Lead Group,Product Test Lead Email,Product Test Lead Location,Product Test Lead Manager,Component PM,Component PM Group,Component PM Email,Component PM Location,Component PM Manager,Component Test Lead,Component Test Lead Group,Component Test Lead Email,Component Test Lead Location,Component Test Lead Manager,Developer,Developer Group,Developer Email,Developer Location,Developer Manager,Tester,Tester Group,Tester Email,Tester Location,Tester Manager,Approver,Approver Group,Approver Email,Approver Location,Approver Manager,Date Opened,Days Open,Days In State,Target Date,Date Closed,Date Modified,Last Modified By,EA Status,Impacts,Test Escape,Test Procedure,Reference Number,ODMs,Suppliers"
		strSaveColumns = "Observation ID,Status,Reviewed,Priority,Severity,Frequency,State,Short Description,Long Description,Steps to Reproduce,Customer Impact,Updates,Component Type,Sub System,Component,Component Version,Localization,Component PartNo,Earliest Product Milestone,Gating Milestone,Owner,Owner Group,Owner Email,Owner Location,Owner Manager,Originator,Originator Group,Originator Email,Originator Location,Originator Manager,Product Segment,Product Family,Primary Product,Affected Product,Product PM,Product PM Group,Product PM Email,Product PM Location,Product PM Manager,Product Test Lead,Product Test Lead Group,Product Test Lead Email,Product Test Lead Location,Product Test Lead Manager,Component PM,Component PM Group,Component PM Email,Component PM Location,Component PM Manager,Component Test Lead,Component Test Lead Group,Component Test Lead Email,Component Test Lead Location,Component Test Lead Manager,Developer,Developer Group,Developer Email,Developer Location,Developer Manager,Tester,Tester Group,Tester Email,Tester Location,Tester Manager,Approver,Approver Group,Approver Email,Approver Location,Approver Manager,Date Opened,Days Open,Days In State,Target Date,Date Closed,Date Modified,Last Modified By,EA Status,Impacts,Test Escape,Test Procedure,Reference Number,ODMs,Suppliers"
	end if

	if trim(strSaveColumns) = "" then
		if trim(request("txtReportSections")) = "3" then
			ColumnArray = split("Observation ID, Primary Product, State, Days In State, Component, Priority, ShortDescription",",")
		else
			ColumnArray = split("Observation ID, Primary Product, State, Component, Priority, Owner, ShortDescription",",")
		end if
	else
		ColumnArray = split(trim(strSaveColumns),",")
	end if

	dim statesUI
	dim statesFIP
	statesUI = "'New*/Reopen', 'Understood/Problem Identified', 'Under Investigation', 'Fix Failed', 'Need Info', 'Transfer Requested', 'Cannot Duplicate - Disagree', 'Duplicate - Disagree', 'No Fix Needed - Disagree', 'Will Not Fix - Disagree'"
	statesFIP = "'Fix Implemented', 'Fix In Progress - Waiting on Vendor', 'Fix In Progress'"

	'Generate SQL Statement
	dim blnExcaliburDataRequired
	dim blnEmployeeDataRequired
	dim strSQLJoins
	dim strSQLTables
	dim strSQLSelect
	dim strSQLFilters
	dim strSQL
	dim RowsDisplayed
	dim groupName

	RowsDisplayed = 0

	if instr(lcase(strSaveColumns),"core team")>0 or instr(lcase(strSaveColumns),"coreteam")>0 then
		blnExcaliburDataRequired= true
	elseif instr(lcase(strSavedAdvanced),"coreteam")>0 then
		blnExcaliburDataRequired= true
	elseif trim(strSavedCoreTeam) <> "" then
		blnExcaliburDataRequired= true
	else
		blnExcaliburDataRequired= false
	end if

	if instr(lcase(strSaveColumns), "manager") > 0 or instr(lcase(strSaveColumns), "location") > 0 then
		blnEmployeeDataRequired = true
	end if

	strSQLSelect = "Select ObservationID, ProductFamily, PrimaryProduct, ComponentType, SubSystem, Component, ComponentVersion, ComponentPartNo, SAPartNumber, OriginatorGroup, Localization, ShortDescription, OriginatorID, OriginatorName, Originator, OriginatorGroupID, OwnerID, OwnerName, Owner, OwnerGroupID, OwnerGroup, FailedFixes, LastModifiedBy, Status, State, Priority, Severity, SeverityName, Frequency, GatingMilestone, Impacts, EAStatus, EANumber, DaysInState, DaysCurrentOwner, DaysOpen, ClosedInVersion, LastReleaseTested, ReleaseFixImplemented, Reproducible, ReferenceNumber, ActualFixTime, StepsToReproduce, LongDescription, EADate, DateClosed, TargetDate, DateOpened, DateModified, SupplierVersion ,CustomerImpact, ImplementationCheck, ApprovalCheck, TestEscape, OnBoard, DivisionID, Division, TestProcedure, Reviewed, EarliestProductMilestone, Feature,SourceSystem, DeveloperID, DeveloperName, Developer, DeveloperGroupID, DeveloperGroup, ComponentPMID, ComponentPMName, ComponentPM, ComponentPMGroupID, ComponentPMGroup, ProductPMID, ProductPMName, ProductPM, ProductPMGroupID, ProductPMGroup, ComponentTestLeadID, ComponentTestLeadName, ComponentTestLead, ComponentTestLeadGroupID, ComponentTestLeadGroup, ProductTestLeadID, ProductTestLeadName, ProductTestLead, ProductTestLeadGroupID, ProductTestLeadGroup, ApproverID, ApproverName, Approver, ApproverGroupID, ApproverGroup, TesterID, TesterName, Tester, TesterGroupID, TesterGroup "
	strSQLTables = " from dbo.SI_Observation_Report o with (NOLOCK) "
	strSQlJoins = ""
	if instr(lcase(strSavedAdvanced),"affectedproduct") > 0 then
		strSQLSelect = strSQLSelect & ", a.AffectedState as AffectedState, a.AffectedProduct as AffectedProduct, 1 as AffectedproductCount "
		strSQLTables = strSQLTables & " , dbo.vAffectedProducts a with (NOLOCK) "
		strSQLJoins = strSQLJoins & " and a.observation_id = o.observationid "
	elseif trim(strSavedAffectedProduct) <> "" then
		strDirectAffectedFilters = ListPrep(strSavedAffectedProduct ,"AffectedProduct",0)
		if strSavedAffectedState <> "" then
		strDirectAffectedFilters = strDirectAffectedFilters & ListPrep(strSavedAffectedState ,"AffectedState",0)
		end if
		if strDirectAffectedFilters <> "" then
			strDirectAffectedFilters = mid(strDirectAffectedFilters,6)
		end if
		strSQLSelect = strSQLSelect & ", a.AffectedState as AffectedState, a.AffectedProduct as AffectedProduct, a.AffectedProductCount "
		strSQLTables = strSQLTables & " , (SELECT observation_id, COUNT(1) as AffectedProductCount, MAX(AffectedProduct) as AffectedProduct, Max(AffectedState) as AffectedState FROM [dbo].[vAffectedProducts] with (NOLOCK) where " & strDirectAffectedFilters & " group by observation_id) a "
		strSQLJoins = strSQLJoins & " and a.observation_id = o.observationid "
	else
		strSQLSelect = strSQLSelect & ",'' as AffectedState, '' as AffectedProduct, 1 as AffectedProductCount "
	end if

	if instr("," & lcase(replace(strSaveColumns," ","")) & ",",",updates,") > 0 or trim(request("txtReportSections")) = "1" then
		if CurrentUserPartner = 1 then
			'HP gets the cached updates, which include secure updates
			strSQLSelect = strSQLSelect & ", o.Updates "
		else
			'non-HP must not get secure updates!
			strSQLSelect = strSQLSelect & ", dbo.udf_collateUpdatesFor(o.observationid, char(3), char(4), 0) as Updates "
		end if
	else
		strSQLSelect = strSQLSelect & ",'' as Updates "
	end if
'	if instr("," & lcase(replace(strSaveColumns," ","")) & ",",",history,") > 0 then
'			strSQLSelect = strSQLSelect & ", dbo.udf_collateHistoryFor(o.observationid, char(3), char(4)) as History "
'	else
'		strSQLSelect = strSQLSelect & ",'' as History "
'	end if
	if blnExcaliburDataRequired then
		strSQLSelect = strSQLSelect & " , [dbo].[ufn_GetCoreTeamNameFromComponentName](o.Component) as CoreTeam, [dbo].[ufn_GetCoreTeamIDFromComponentName](o.Component) as CoreteamID "
'		strSQLTables = strSQLTables & ", (Select ct.id as CoreTeamID, ct.name as coreteam, v.id as VersionID, v.otspartnumber from prs.dbo.deliverablecoreteam ct with (NOLOCK), prs.dbo.deliverableroot r with (NOLOCK), prs.dbo.deliverableversion v with (NOLOCK) where v.deliverablerootid = r.id and ct.id = r.coreteamid union Select 0 as CoreTeamID, 'None' as CoreTeam,ID as versionid, OTSPartNumber from prs.dbo.OTSComponent with (NOLOCK)) ex "
'		strSQLJoins = strSQLJoins & " and ex.versionid = o.ExcaliburNumber and o.DivisionID = 6 "
'		strSQLTables = strSQLTables & ", (Select ct.id as CoreTeamID, ct.name as coreteam, v.id as VersionID from prs.dbo.deliverablecoreteam ct with (NOLOCK), prs.dbo.deliverableroot r with (NOLOCK), prs.dbo.deliverableversion v with (NOLOCK) where v.deliverablerootid = r.id and ct.id = r.coreteamid union Select 0 as CoreTeamID, 'None' as CoreTeam, ID as versionid from prs.dbo.OTSComponent with (NOLOCK) union Select 0 as CoreTeamID, 'None' as CoreTeam,0 as versionid) ex "
'		strSQLJoins = strSQLJoins & " and ((o.DivisionID = 6 and ex.versionid = o.ExcaliburNumber) or (o.DivisionID <> 6 and ex.versionid = 0)) "
	else
		strSQLSelect = strSQLSelect & ", '' as CoreTeam , 0 as CoreteamID "
	end if

	if blnEmployeeDataRequired then
		strSQLSelect = strSQLSelect & _
			", aml.location as approverLocation " & _
			", aml.managername as approverManager " & _
			", cpmml.location as componentPmLocation " & _
			", cpmml.managername as componentPmManager " & _
			", ctlml.location as componentTestLeadLocation " & _
			", ctlml.managername as componentTestLeadManager " & _
			", dml.location as developerLocation " & _
			", dml.managername as developerManager " & _
			", oml.location as ownerLocation " & _
			", oml.managername as ownerManager " & _
			", orml.location as originatorLocation " & _
			", orml.managername as originatorManager " & _
			", ppmml.location as productPmLocation " & _
			", ppmml.managername as productPmManager " & _
			", ptlml.location as productTestLeadLocation " & _
			", ptlml.managername as productTestLeadManager " & _
			", tml.location as testerLocation " & _
			", tml.managername as testerManager "
		strSQLTables = strSQLTables & _
			", dbo.vUserManagerAndLocation aml with (nolock) " & _
			", dbo.vUserManagerAndLocation cpmml with (nolock) " & _
			", dbo.vUserManagerAndLocation ctlml with (nolock) " & _
			", dbo.vUserManagerAndLocation dml with (nolock) " & _
			", dbo.vUserManagerAndLocation oml with (nolock) " & _
			", dbo.vUserManagerAndLocation orml with (nolock) " & _
			", dbo.vUserManagerAndLocation ppmml with (nolock) " & _
			", dbo.vUserManagerAndLocation ptlml with (nolock) " & _
			", dbo.vUserManagerAndLocation tml with (nolock) "
		strSQLJoins = strSQLJoins & _
			" and aml.user_id = approverid " & _
			" and cpmml.user_id = componentpmid " & _
			" and ctlml.user_id = componenttestleadid " & _
			" and dml.user_id = developerid " & _
			" and oml.user_id = ownerid " & _
			" and orml.user_id = originatorid " & _
			" and ppmml.user_id = productpmid " & _
			" and ptlml.user_id = producttestleadid " & _
			" and tml.user_id = testerid "
	end if

	if trim(strSaveSearchDetails) = "1" or trim(strSaveSearchSummary) = "1" or trim(strSaveSearchImpact) = "1" or trim(strSaveSearchReproduce) = "1" then
		blnMainTextSearchSelected = true
	else
		blnMainTextSearchSelected = false
	end if
	if trim(strSavedTextSearch) <> "" and (blnMainTextSearchSelected or trim(strSaveSearchHistory) = "1") then
		dim strSearchFieldList

		strSearchFieldList = ""
		if trim(strSaveSearchSummary) = "1" then
			strSearchFieldList = strSearchFieldList & ",short_description"
		end if
		if trim(strSaveSearchDetails) = "1" then
			strSearchFieldList = strSearchFieldList & ",long_description"
		end if
		if trim(strSaveSearchImpact) = "1" then
			strSearchFieldList = strSearchFieldList & ",steps_to_reproduce"
		end if
		if trim(strSaveSearchReproduce) = "1" then
			strSearchFieldList = strSearchFieldList & ",Customer_impact"
		end if
		if trim(strSaveSearchType) <> "2" then
			strSQLSelect = strSQLSelect & " , ts.[RANK] as SearchRank "
		end if
		if trim(strSaveSearchType) = "1" or trim(strSaveSearchType) = "3" then
			strKeywordSearchList=""
			KeyWordSearchArray = split(strSavedTextSearch," ")
			for i = 0 to ubound(KeyWordSearchArray)
				strCleanedSearchText = replace(replace(KeyWordSearchArray(i),"'","''"),"""","""""")
				if trim(KeyWordSearchArray(i)) <> "" then
					if strKeywordSearchList = "" then
						strKeywordSearchList = """*" & strCleanedSearchText & "*"""
					elseif trim(strSaveSearchType) = "3" then
						strKeywordSearchList = strKeywordSearchList & " or ""*" & strCleanedSearchText & "*"""
					else
						strKeywordSearchList = strKeywordSearchList & " and ""*" & strCleanedSearchText & "*"""
					end if
				end if
			next
		end if

		if trim(strSaveSearchType) = "2" then
			strSQLSelect = strSQLSelect & " , 0 as SearchRank "
		else
			if blnMainTextSearchSelected and trim(strSaveSearchHistory) = "1" then
				'Main and History Fields
				if trim(strSaveSearchType) = "1" or trim(strSaveSearchType) = "3" then
					strSQLTables = strSQLTables & ", ( select distinct [KEY], max([RANK]) as [RANK] from ( Select [KEY],[RANK] from CONTAINSTABLE(dbo.Observation, (" & mid(strSearchFieldList,2) & "),'(" & strKeywordSearchList & ")') union Select h.Object_ID as [KEY], max([RANK]) as [RANK] from CONTAINSTABLE(dbo.history, (Log_Summary),'(" & strKeywordSearchList & ")') AS fh, dbo.History h with (NOLOCK) where fh.[KEY] = h.ID Group by h.Object_ID ) ts_temp group by [KEY] ) ts "
				else
					strSQLTables = strSQLTables & ", ( select distinct [KEY], max([RANK]) as [RANK] from ( Select [KEY],[RANK] from FREETEXTTABLE(dbo.Observation, (" & mid(strSearchFieldList,2) & "),'" & replace(strSavedTextSearch,"'","''") & "') union Select h.Object_ID as [KEY], max([RANK]) as [RANK] from FREETEXTTABLE(dbo.history, (Log_Summary),'" & replace(strSavedTextSearch,"'","''") & "') AS fh, dbo.History h with (NOLOCK) where fh.[KEY] = h.ID Group by h.Object_ID ) ts_temp group by [KEY] ) ts "
				end if
			elseif trim(strSaveSearchHistory) = "1" then
				'History Only
				if trim(strSaveSearchType) = "1" or trim(strSaveSearchType) = "3" then
					strSQLTables = strSQLTables & ", (Select h.Object_ID as [KEY], max([RANK]) as [RANK] from CONTAINSTABLE(dbo.history, (Log_Summary),'(" & strKeywordSearchList & ")') AS fh, dbo.History h with (NOLOCK) where fh.[KEY] = h.ID Group by h.Object_ID) ts "
				else
					strSQLTables = strSQLTables & ", (Select h.Object_ID as [KEY], max([RANK]) as [RANK] from FREETEXTTABLE(dbo.history, (Log_Summary),'" & replace(strSavedTextSearch,"'","''") & "') AS fh, dbo.History h with (NOLOCK) where fh.[KEY] = h.ID Group by h.Object_ID) ts "
				end if

			else
				'Main Fields Only
				if trim(strSaveSearchType) = "1" or trim(strSaveSearchType) = "3" then
					strSQLTables = strSQLTables & ", CONTAINSTABLE(dbo.Observation, (" & mid(strSearchFieldList,2) & "),'(" & strKeywordSearchList & ")') ts "
				elseif trim(strSaveSearchType) = "0" or trim(strSaveSearchType) = "" then
					strSQLTables = strSQLTables & ", FREETEXTTABLE(dbo.Observation, (" & mid(strSearchFieldList,2) & "),'" & replace(strSavedTextSearch,"'","''") & "') ts "
				end if
			end if
			strSQLJoins = strSQLJoins & " and ts.[key] = o.observationid "
		end if
	else
		strSQLSelect = strSQLSelect & " , 0 as SearchRank "
	end if

	if strSavedProductGroup <> "" then
		strSQLSelect = strSQLSelect & " , ep.ProductPartnerID, ep.ProductStatusID, ep.DevCenterID "
		strSQLTables = strSQLTables & ", dbo.vProductGroups ep with (NOLOCK) "
		strSQLJoins = strSQLJoins & " and ep.ProductName = o.PrimaryProduct "
	else
		strSQLSelect = strSQLSelect & ",0 as ProductPartnerID, 0 as ProductStatusID , 0 as DevCenterID "
	end if
	strSQLFilters = ""

	if trim(strSavedObservationID) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(replace(strSavedObservationID,"%A3%AC",","),"ObservationID",1)
	end if

	if trim(strSavedProduct) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProduct,"PrimaryProduct",0)
	end if

	if trim(strSavedSubsystem) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedSubsystem,"SubSystem",0)
	end if

	if trim(strSavedCoreTeam) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedCoreTeam,"CoreTeamID",1)
	end if

	if trim(strSaveFrequency) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSaveFrequency,"FrequencyID",1)
	end if

	if trim(strSavedOwner) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedOwner,"OwnerID",1)
	end if

	if trim(strSavedOwnerGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedOwnerGroup,"OwnerGroupID",1)
	end if

	if trim(strSavedType) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedType,"ComponentType",0)
	end if

	if trim(strSavedComponent) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedComponent,"Component",0)
	end if

	if trim(strSavedAffectedProduct) <> "" and instr(lcase(strSavedAdvanced),"affectedproduct") > 0 then
		strSQLFilters = strSQLFilters & ListPrep(strSavedAffectedProduct ,"AffectedProduct",0)
	end if

	if trim(strSavedAffectedState) <> "" and instr(lcase(strSavedAdvanced),"affectedproduct") > 0 then
		strSQLFilters = strSQLFilters & ListPrep(strSavedAffectedState,"AffectedState",0)
	end if

	if trim(strSavedApprover) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedApprover,"ApproverID",1)
	end if

	if trim(strSavedApproverGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedApproverGroup,"ApproverGroupID",1)
	end if

	if trim(strSavedProductPM) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductPM,"ProductPMID",1)
	end if

	if trim(strSavedProductPMGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductPMGroup,"ProductPMGroupID",1)
	end if

	if trim(strSavedOriginator) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedOriginator,"OriginatorID",1)
	end if

	if trim(strSavedOriginatorGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedOriginatorGroup,"OriginatorGroupID",1)
	end if

	if trim(strSavedState) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedState,"State",0)
	end if

	if trim(strSavedComponentPM) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedComponentPM,"ComponentPMID",1)
	end if

	if trim(strSavedComponentPMGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedComponentPMGroup,"ComponentPMGroupID",1)
	end if

	if trim(strSavedDeveloper) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedDeveloper,"DeveloperID",1)
	end if

	if trim(strSavedAssigned) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedAssigned,"Assigned",1)
	end if

	if trim(strSavedDeveloperGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedDeveloperGroup,"DeveloperGroupID",1)
	end if

	if trim(strSavedTester) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedTester,"TesterID",1)
	end if

	if trim(strSavedTesterGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedTesterGroup,"TesterGroupID",1)
	end if

	if trim(strSavedGatingMilestone) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedGatingMilestone,"GatingMilestone",0)
	end if

	if trim(strSavedFeature) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedFeature,"Feature",0)
	end if

	if trim(strSavedComponentTestLead) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedComponentTestLead,"ComponentTestLeadID",1)
	end if

	if trim(strSavedComponentTestLeadGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedComponentTestLeadGroup,"ComponentTestLeadGroupID",1)
	end if

	if trim(strSavedProductTestLead) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductTestLead,"ProductTestLeadID",1)
	end if

	if trim(strSavedProductTestLeadGroup) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductTestLeadGroup,"ProductTestLeadGroupID",1)
	end if

	if trim(strSavedProductAndVersion) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductAndVersion,"ProductAndVersion",0)
	end if

	if trim(strSavedProductFamily) <> "" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedProductFamily,"ProductFamily",0)
	end if

	if trim(strSavedPriority) <> "" and trim(strSavedPriority) <> "0, 1, 2, 3, 4" then
		strSQLFilters = strSQLFilters & ListPrep(strSavedPriority,"Priority",1)
	end if

	if trim(strSavedSeverity) <> "" then
		strSQLFilters = strSQLFilters & ItemPrep(strSavedSeverity ,"SeverityName",0)
	end if

	if trim(strSavedImpact) <> "" then
		strSQLFilters = strSQLFilters & ItemPrep(strSavedImpact ,"Impacts",2)
	end if

	if trim(strSavedStatus) <> "" then
		strSQLFilters = strSQLFilters & ItemPrep(strSavedStatus ,"Status",0)
	end if

	if trim(strSavedEscape) <> "" then
		strSQLFilters = strSQLFilters & ItemPrep(strSavedEscape ,"TestEscape",2)
	end if

	if lcase(trim(strSavedDivision)) = "mobile" then
		strSQLFilters = strSQLFilters & " and DivisionID = 6 "
	elseif trim(strSavedDivision) <> "" then
		strSQLFilters = strSQLFilters & " and DivisionID <> 6 "
	end if

	if trim(CurrentUserPartner) <> "1" and (not blnMobileODMUser) then
		strSQLTables = strSQLTables & ", dbo.vObservationLinkedVendors olv with (NOLOCK) "
		strSQLJoins = strSQLJoins & " and olv.Observation_ID = o.ObservationId "
		strSQLFilters = strSQLFilters & " and (" & strLimitPartner & ") "
	end if

	if trim(strSavedDaysOpenCompare) <> "" then
		strSQLFilters = strSQLFilters & DaysPrep(strSavedDaysOpenDays ,strSavedDaysOpenRange,"DaysOpen",strSavedDaysOpenCompare)
	end if

	if trim(strSavedDaysStateCompare) <> "" then
		strSQLFilters = strSQLFilters & DaysPrep(strSavedDaysStateDays, strSavedDaysStateRange,"DaysInState",strSavedDaysStateCompare)
	end if

	if trim(strSavedDaysOwnerCompare) <> "" then
		strSQLFilters = strSQLFilters & DaysPrep(strSavedDaysOwnerDays , strSavedDaysOwnerRange,"DaysCurrentOwner",strSavedDaysOwnerCompare)
	end if

	if trim(strSavedDateOpenedCompare) <> "" then
		strSQLFilters = strSQLFilters & DatePrep( strSavedDateOpenedDays ,strSavedDateDateOpenedRange1,strSavedDateDateOpenedRange2,"DateOpened",strSavedDateOpenedCompare)
	end if

	if trim(strSaveSearchType) = "2" then
		strTemp = ""
		if trim(strSaveSearchSummary) = "1" then
			strTemp = strTemp & " or shortDescription like '%" & replace(replace(replace(replace(replace(replace(strSavedTextSearch,"'","''"),"""",""""""),"[]"," "),"[","[[]"),"%","[%]"),"_","[_]") & "%' "
		end if
		if trim(strSaveSearchDetails) = "1" then
			strTemp = strTemp & " or longDescription like '%" & replace(replace(replace(replace(replace(replace(strSavedTextSearch,"'","''"),"""",""""""),"[]"," "),"[","[[]"),"%","[%]"),"_","[_]") & "%' "
		end if
		if trim(strSaveSearchImpact) = "1" then
			strTemp = strTemp & " or customerimpact like '%" & replace(replace(replace(replace(replace(replace(strSavedTextSearch,"'","''"),"""",""""""),"[]"," "),"[","[[]"),"%","[%]"),"_","[_]") & "%' "
		end if
		if trim(strSaveSearchReproduce) = "1" then
			strTemp = strTemp & " or stepstoreproduce like '%" & replace(replace(replace(replace(replace(replace(strSavedTextSearch,"'","''"),"""",""""""),"[]"," "),"[","[[]"),"%","[%]"),"_","[_]") & "%' "
		end if
		if strTemp <> "" then
			strSQLFilters = strSQLFilters & " and (" & scrubsql(mid(strTemp,4)) & ") "
		end if
	end if

	if trim(strSavedDateClosedCompare) <> "" then
		strSQLFilters = strSQLFilters & DatePrep(strSavedDateClosedDays ,strSavedDateClosedRange1,strSavedDateClosedRange2,"DateClosed",strSavedDateClosedCompare)
	end if

	if trim(strSavedDateModifiedCompare) <> "" then
		strSQLFilters = strSQLFilters & DatePrep(strSavedDateModifiedDays ,strSavedDateModifiedRange1,strSavedDateModifiedRange2,"DateModified",strSavedDateModifiedCompare)
	end if

	if trim(strSavedTargetDateCompare) <> "" then
		strSQLFilters = strSQLFilters & DatePrep(strSavedTargetDateDays , strSavedTargetDateRange1,strSavedTargetDateRange2,"TargetDate",strSavedTargetDateCompare)
	end if

	if trim(strSavedAdvanced) <> "" then
		if instr(lcase(strSavedAdvanced), "coreteam") > 0 then
			strSavedAdvanced = replace(lcase(strSavedAdvanced),"coreteam","[dbo].[ufn_GetCoreTeamNameFromComponentName](o.Component)")
		end if
		strSQLFilters = strSQLFilters & " and ( " & replace(ScrubSQL(strSavedAdvanced),"""","'") & " ) "
	end if

	if trim(strSavedProductGroup) <> "" then
		dim strCycle, strDevCenter, strODM, strProdPhase, strProductGroupSQL, strTemp

		ItemArray = split(strSavedProductGroup,",")
		strCycle = ""
		strDevCenter = ""
		strODM = ""
		strProdPhase = ""

		for each strItem in ItemArray
			if trim(strItem) <> "" then
				ValuePair = split(strItem,":")
				select case clng(ValuePair(0))
				case 1
					strODM = strODM & "," & clng(ValuePair(1))
				case 2
					strCycle = strCycle & "," & clng(ValuePair(1))
				case 3
					strDevCenter = strDevCenter & "," & clng(ValuePair(1))
				case 4
					strProdPhase = strProdPhase & "," & clng(ValuePair(1))
				end select
			end if
		next

		if trim(strODM) <> "" then
			strSQLFilters = strSQLFilters & " and ProductPartnerID in (" & scrubsql(mid(strODM,2)) & ") "
		end if
		if trim(strDevCenter) <> "" then
			strSQLFilters = strSQLFilters & " and DevCenterID in (" & scrubsql(mid(strDevCenter,2)) & ") "
		end if
		if trim(strProdPhase) <> "" then
			strSQLFilters = strSQLFilters & " and ProductStatusID in (" & scrubsql(mid(strProdPhase,2)) & ") "
		end if
		if trim(strCycle) <> "" then
			strProductGroupSQL = "Select distinct v.DotsName as Product " & _
								"from Product_Program p with (NOLOCK), ProductVersion v with (NOLOCK) " & _
								"where v.id = p.productversionid " & _
								"and p.programid in (" & scrubsql(mid(strCycle,2)) & ")"
			rs.open strProductGroupSQL,cnExcalibur
			strTemp = ""
			do while not rs.EOF
				strTemp = strTemp & ",'" & rs("Product") & "' "
				rs.MoveNext
			loop
			rs.Close

			if trim(strTemp) <> "" then
				strSQLFilters = strSQLFilters & " and PrimaryProduct in (" & scrubsql(mid(strTemp,2)) & ") "
			end if
		end if
	end if

'	if blnStatusReport and not FilterContains(strSQLFilters,"status","closed") then
'		strSQLFilters = strSQLFilters & " and Status <> 'Closed' "
'	end if

	if instr(lcase(strSaveColumns), "odms") or instr(lcase(strSaveColumns), "suppliers") then
		strSQLTables = strSQLTables & ", dbo.vObservationLinkedVendors lv with (NOLOCK) "
		strSQLJoins = strSQLJoins & " and lv.observation_id = o.observationid "
	end if
	if instr(lcase(strSaveColumns), "odms") then
		strSQLSelect = strSQLSelect & ", lv.[odms] "
	end if
	if instr(lcase(strSaveColumns), "suppliers") then
		strSQLSelect = strSQLSelect & ", lv.[suppliers] "
	end if
	if instr(lcase(strSaveColumns), "segment") then
		strSQLSelect = strSQLSelect & ", ops.[productsegment] "
		strSQLTables = strSQLTables & ", dbo.vObservationProductSegments ops with (NOLOCK) "
		strSQLJoins = strSQLJoins & " and ops.observation_id = o.observationid "
	end if
	if instr(lcase(strSaveColumns), "updates2") then
		strSQLSelect = strSQLSelect & ", dbo.udf_collateUpdates2For(o.observationid, char(3), char(4)) as updates2 "
	end if
	if instr(lcase(strSaveColumns), "count component assignments") > 0 then
		strSQLSelect = strSQLSelect & ", dbo.udf_countAssignmentsFor(o.observationid, 'Primary Component') as CountComponentAssignments "
	end if
	if instr(lcase(strSaveColumns), "count owner assignments") > 0 then
		strSQLSelect = strSQLSelect & ", dbo.udf_countAssignmentsFor(o.observationid, 'Owner') as CountOwnerAssignments "
	end if

	strSQl = strSQLSelect & " " & strSQlTables & " where "

	if trim(strSQLJoins) <> "" then
		strSQl = strSQl & " " & mid(strSQLJoins,5) & " "
	end if

	if trim(strSQLFilters) <> "" and trim(strSQLJoins) = "" then
		strSQl = strSql & mid(strSQLFilters,5)
	elseif trim(strSQLFilters) <> "" then
		strSQl = strSql & strSQLFilters
	end if

	dim strSQLOpen

	if blnStatusReport and not FilterContains(strSQLFilters,"status","closed") then
		strSQLOpen = strSQl & " and status <> 'Closed' "
	else
		strSQLOpen = strSQl
	end if

	dim strSortOrder
	dim strSortCount
	strSortOrder = ""
	strSortCount = 0
	if trim(request("txtReportSections")) = "3" then
		strSortOrder = strSortOrder & ",owner"
		strSortCount = strSortCount + 1
		if lcase(trim(strSavedSortColumn1)) = "owner" then
			strSavedSortColumn1 = ""
		end if
		if lcase(trim(strSavedSortColumn2)) = "owner" then
			strSavedSortColumn2 = ""
		end if
		if lcase(trim(strSavedSortColumn3)) = "owner" then
			strSavedSortColumn3 = ""
		end if
	end if
	if trim(strSavedSortColumn1) <> "" then
		strSortOrder = strSortOrder & "," & replace(strSavedSortColumn1," ","")
		if trim(lcase(strSavedSort1Direction)) = "desc" then
			strSortOrder = strSortOrder & " desc"
		end if
		strSortCount = strSortCount + 1
	end if
	if trim(strSavedSortColumn2) <> "" and lcase(trim(strSavedSortColumn1)) <> lcase(trim(strSavedSortColumn2)) then
		strSortOrder = strSortOrder & "," & replace(strSavedSortColumn2," ","")
		if trim(lcase(strSavedSort2Direction)) = "desc" then
			strSortOrder = strSortOrder & " desc"
		end if
		strSortCount = strSortCount + 1
	end if
	if strSortCount < 3 and trim(strSavedSortColumn3) <> "" and lcase(trim(strSavedSortColumn3)) <> lcase(trim(strSavedSortColumn2)) and lcase(trim(strSavedSortColumn3)) <> lcase(trim(strSavedSortColumn1)) then
		strSortOrder = strSortOrder & "," & replace(strSavedSortColumn3," ","")
		if trim(lcase(strSavedSort3Direction)) = "desc" then
			strSortOrder = strSortOrder & " desc"
		end if
		strSortCount = strSortCount + 1
	end if
	if trim(strSavedTextSearch) <> "" and strSortCount <= 3 and lcase(trim(strSavedSortColumn1)) <> "search rank" and lcase(trim(strSavedSortColumn2)) <> "search rank" and lcase(trim(strSavedSortColumn3)) <> "search rank" then
		strSortOrder = strSortOrder & ",searchrank desc"
	end if

	if strSortOrder = "" then
		strSortOrder = " order by observationid;"
		strSQl = strSQl & strSortOrder
		strSQLOpen = strSQLOpen & strSortOrder
	else
		strSortOrder = mid(strSortOrder,2) & ";"
		strSQl = strSQl & " order by " & strSortOrder
		strSQLOpen = strSQLOpen & " order by " & strSortOrder
	end if


	if strSQLFilters = "" then
		response.write "Not enough filters were selected for this report. Please select at least one primary filter and try again."
	else
'		if clng(currentuserid) = 31 and trim(request("txtReportSections")) <> "-1" then
'			response.write "<BR>" & strSQL & "<BR><BR>"
'			response.Flush
'		end if
		dim PageSectionArray
		dim blnAllowHeaderSorting
		dim ColumnCount
		dim PageParamsArray

		if instr(strSQl, "searchrank") > 0 then
			redim preserve ColumnArray(ubound(ColumnArray)+1)
			ColumnArray(ubound(ColumnArray)) = "searchrank"
		end if
		
		PageSectionArray = split(PageSections,",")
		if strPageParams <> "" then
			PageParamsArray = split(strPageParams,"|")
		end if
		strColumnHeaders=""
		SectionCount=0
		for each strSection in PageSectionArray
			blnAllowHeaderSorting = not (trim(strSavedFormat) = "Excel" or trim(strSavedFormat) = "Word" or trim(strSection) = "3")
			'Generate SQL Header string
			if (trim(strSection)="0" or trim(strSection)="" or trim(strSection) = "3") and trim(strColumnHeaders)="" then
				i=0
				for each strColumn in ColumnArray
					Select case lcase(replace(strColumn," ",""))
					case "searchrank"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Search&nbsp;Rank</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,2);"">Search&nbsp;Rank</a></td>"
						end if
					case "observationid"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>OTS&nbsp;ID</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,2);"">OTS&nbsp;ID</a></td>"
						end if
					case "primaryproduct"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Primary&nbsp;Product</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Primary&nbsp;Product</a></td>"
						end if
					case "state"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>State</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">State</a></td>"
						end if
					case "component"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component</a></td>"
						end if
					case "priority"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Pr</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Pr</a></td>"
						end if
					case "owner"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Owner</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Owner</a></td>"
						end if
					case "shortdescription"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Short&nbsp;Description</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Short&nbsp;Description</a></td>"
						end if
					case "failedfixes"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Failed&nbsp;Fixes</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Failed&nbsp;Fixes</a></td>"
						end if
					case "productfamily"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Family</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Family</a></td>"
						end if
					case "feature"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Feature</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Feature</a></td>"
						end if
					case "approver"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approver</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approver</a></td>"
						end if
					case "approveremail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approver&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approver&nbsp;Email</a></td>"
						end if
					case "approvergroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approver&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approver&nbsp;Group</a></td>"
						end if
					case "approverlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approver&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approver&nbsp;Location</a></td>"
						end if
					case "approvermanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approver&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approver&nbsp;Manager</a></td>"
						end if
					case "componentpm"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PM</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PM</a></td>"
						end if
					case "componentpmemail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PM&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PM&nbsp;Email</a></td>"
						end if
					case "componentpmgroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PM&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PM&nbsp;Group</a></td>"
						end if
					case "componentpmlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PM&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PM&nbsp;Location</a></td>"
						end if
					case "componentpmmanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PM&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PM&nbsp;Manager</a></td>"
						end if
					case "closedinversion"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Closed&nbsp;In&nbsp;Version</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Closed&nbsp;In&nbsp;Version</a></td>"
						end if
					case "affectedproduct"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Affected&nbsp;Product</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Affected&nbsp;Product</a></td>"
						end if
					case "affectedstate"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Affected&nbsp;State</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Affected&nbsp;State</a></td>"
						end if
					case "approvalcheck"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Approval&nbsp;Check</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Approval&nbsp;Check</a></td>"
						end if
					case "componenttestlead"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Test&nbsp;Lead</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Test&nbsp;Lead</a></td>"
						end if
					case "componenttestleademail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Test&nbsp;Lead&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Test&nbsp;Lead&nbsp;Email</a></td>"
						end if
					case "componenttestleadgroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Test&nbsp;Lead&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Test&nbsp;Lead&nbsp;Group</a></td>"
						end if
					case "componenttestleadlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Test&nbsp;Lead&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Test&nbsp;Lead&nbsp;Location</a></td>"
						end if
					case "componenttestleadmanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Test&nbsp;Lead&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Test&nbsp;Lead&nbsp;Manager</a></td>"
						end if
'					case "history"
'						strColumnHeaders = strColumnHeaders & "<td>History</td>"
					case "updates"
						strColumnHeaders = strColumnHeaders & "<td>Updates</td>"
					case "updates2"
						strColumnHeaders = strColumnHeaders & "<td>Updates</td>"
					case "componenttype"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Type</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Type</a></td>"
						end if
					case "componentversion"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;Version</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;Version</a></td>"
						end if
					case "componentpartno"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Component&nbsp;PartNo</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Component&nbsp;PartNo</a></td>"
						end if
					case "countcomponentassignments"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Count&nbsp;Component&nbsp;Assignments</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,1);"">Count&nbsp;Component&nbsp;Assignments</a></td>"
						end if
					case "countownerassignments"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Count&nbsp;Owner&nbsp;Assignments</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,1);"">Count&nbsp;Owner&nbsp;Assignments</a></td>"
						end if
					case "coreteam"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Core&nbsp;Team</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Core&nbsp;Team</a></td>"
						end if
					case "originator"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Originator</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Originator</a></td>"
						end if
					case "originatoremail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Originator&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Originator&nbsp;Email</a></td>"
						end if
					case "originatorgroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Originator&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Originator&nbsp;Group</a></td>"
						end if
					case "originatorlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Originator&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Originator&nbsp;Location</a></td>"
						end if
					case "originatormanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Originator&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Originator&nbsp;Manager</a></td>"
						end if
					case "customerimpact"
'						if trim(strSavedFormat) = "Excel" then
							strColumnHeaders = strColumnHeaders & "<td>Customer&nbsp;Impact</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Customer&nbsp;Impact</a></td>"
'						end if
					case "longdescription"
'						if trim(strSavedFormat) = "Excel" then
							strColumnHeaders = strColumnHeaders & "<td>Long&nbsp;Description</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Customer&nbsp;Impact</a></td>"
'						end if
					case "stepstoreproduce"
'						if trim(strSavedFormat) = "Excel" then
							strColumnHeaders = strColumnHeaders & "<td>Steps&nbsp;To&nbsp;Reproduce</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Customer&nbsp;Impact</a></td>"
'						end if
					case "dateopened"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Date&nbsp;Opened</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">Date&nbsp;Opened</a></td>"
						end if
					case "dateclosed"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Date&nbsp;Closed</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">Date&nbsp;Closed</a></td>"
						end if
					case "datemodified"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Date&nbsp;Modified</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">Date&nbsp;Modified</a></td>"
						end if
					case "daysopen"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Days&nbsp;Open</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,1);"">Days&nbsp;Open</a></td>"
						end if
					case "daysinstate"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Days&nbsp;State</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,1);"">Days&nbsp;State</a></td>"
						end if
					case "dayscurrentowner"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Days&nbsp;Owner</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",1,1);"">Days&nbsp;Owner</a></td>"
						end if
					case "developer"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Developer</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Developer</a></td>"
						end if
					case "developeremail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Developer&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Developer&nbsp;Email</a></td>"
						end if
					case "developergroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Developer&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Developer&nbsp;Group</a></td>"
						end if
					case "developerlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Developer&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Developer&nbsp;Location</a></td>"
						end if
					case "developermanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Developer&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Developer&nbsp;Manager</a></td>"
						end if
					case "division"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Division</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Division</a></td>"
						end if
					case "eadate"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>EA&nbsp;Date</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">EA&nbsp;Date</a></td>"
						end if
					case "eanumber"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>EA&nbsp;Number</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",4,1);"">EA&nbsp;Number</a></td>"
						end if
					case "eastatus"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>EA&nbsp;Status</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">EA&nbsp;Status</a></td>"
						end if
					case "earliestproductmilestone"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Earliest&nbsp;Milestone</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Earliest&nbsp;Milestone</a></td>"
						end if
					case "frequency"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Frequency</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Frequency</a></td>"
						end if
					case "gatingmilestone"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Gating&nbsp;Milestone</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Gating&nbsp;Milestone</a></td>"
						end if
					case "impacts"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Impacts</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Impacts</a></td>"
						end if
					case "implementationcheck"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Implementation&nbsp;Check</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Implementation&nbsp;Check</a></td>"
						end if
					case "lastreleasetested"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Last&nbsp;Release&nbsp;Tested</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">Last&nbsp;Release&nbsp;Tested</a></td>"
						end if
					case "lastmodifiedby"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Last&nbsp;Modified&nbsp;By</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Last&nbsp;Modified&nbsp;By</a></td>"
						end if
					case "localization"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Localization</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Localization</a></td>"
						end if
					case "odms"
'						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>ODMs</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">ODMs</a></td>"
'						end if
					case "onboard"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>On&nbsp;Board</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">On&nbsp;Board</a></td>"
						end if
					case "owneremail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Owner&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Owner&nbsp;Email</a></td>"
						end if
					case "ownergroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Owner&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Owner&nbsp;Group</a></td>"
						end if
					case "ownerlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Owner&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Owner&nbsp;Location</a></td>"
						end if
					case "ownermanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Owner&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Owner&nbsp;Manager</a></td>"
						end if
					case "productpm"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;PM</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;PM</a></td>"
						end if
					case "productpmemail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;PM&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;PM&nbsp;Email</a></td>"
						end if
					case "productpmgroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;PM&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;PM&nbsp;Group</a></td>"
						end if
					case "productpmlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;PM&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;PM&nbsp;Location</a></td>"
						end if
					case "productpmmanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;PM&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;PM&nbsp;Manager</a></td>"
						end if
					case "productsegment"
'						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Segment</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Segment</a></td>"
'						end if
					case "producttestlead"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Test&nbsp;Lead</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Test&nbsp;Lead</a></td>"
						end if
					case "producttestleademail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Test&nbsp;Lead&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Test&nbsp;Lead&nbsp;Email</a></td>"
						end if
					case "producttestleadgroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Test&nbsp;Lead&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Test&nbsp;Lead&nbsp;Group</a></td>"
						end if
					case "producttestleadlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Test&nbsp;Lead&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Test&nbsp;Lead&nbsp;Location</a></td>"
						end if
					case "producttestleadmanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Product&nbsp;Test&nbsp;Lead&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Product&nbsp;Test&nbsp;Lead&nbsp;Manager</a></td>"
						end if
					case "referencenumber"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Reference&nbsp;Number</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Reference&nbsp;Number</a></td>"
						end if
					case "releasefiximplemented"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Release&nbsp;Fix&nbsp;Implemented</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Release&nbsp;Fix&nbsp;Implemented</a></td>"
						end if
					case "reviewed"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Reviewed</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Reviewed</a></td>"
						end if
					case "sapartnumber"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>SA&nbsp;Part&nbsp;Number</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">SA&nbsp;Part&nbsp;Number</a></td>"
						end if
					case "severity"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Severity</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Severity</a></td>"
						end if
					case "status"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Status</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Status</a></td>"
						end if
					case "subsystem"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Sub&nbsp;System</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Sub&nbsp;System</a></td>"
						end if
					case "suppliers"
'						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Suppliers</td>"
'						else
'							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Suppliers</a></td>"
'						end if
					case "supplierversion"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Supplier&nbsp;Version</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Supplier&nbsp;Version</a></td>"
						end if
					case "targetdate"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Target&nbsp;Date</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",2,1);"">Target&nbsp;Date</a></td>"
						end if
					case "testescape"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Test&nbsp;Escape</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Test&nbsp;Escape</a></td>"
						end if
					case "testprocedure"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Test&nbsp;Procedure</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Test&nbsp;Procedure</a></td>"
						end if

					case "tester"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Tester</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Tester</a></td>"
						end if
					case "testeremail"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Tester&nbsp;Email</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Tester&nbsp;Email</a></td>"
						end if
					case "testergroup"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Tester&nbsp;Group</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Tester&nbsp;Group</a></td>"
						end if
					case "testerlocation"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Tester&nbsp;Location</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Tester&nbsp;Location</a></td>"
						end if
					case "testermanager"
						if not blnAllowHeaderSorting then
							strColumnHeaders = strColumnHeaders & "<td>Tester&nbsp;Manager</td>"
						else
							strColumnHeaders = strColumnHeaders & "<td><a href=""javascript: SortTable( 'tblOTS', " & i & ",0,1);"">Tester&nbsp;Manager</a></td>"
						end if
					end select
					i=i+1
				next
				ColumnCount = i
				if blnStatusReport then
					if FilterContains(strSQLFilters,"status","closed") then
						strColumnHeaders = "<thead class=statusHeader><tr bgcolor=gainsboro><td align=center colspan=" & ColumnCount & "><b>Observation List</b></td><tr>" & strColumnHeaders & "</tr></thead>"
					else
						strColumnHeaders = "<thead class=statusHeader><tr bgcolor=gainsboro><td align=center colspan=" & ColumnCount & "><b>Open Observation List</b></td><tr>" & strColumnHeaders & "</tr></thead>"
					end if
				else
					strColumnHeaders = "<thead><tr>" & strColumnHeaders & "</tr></thead>"
				end if
			end if

			'Build Section
			if trim(strSection) = "7" then 'Affected product List
				rs.open strSQl,cnSIO
				if rs.eof and rs.bof then
					response.write "Unable to find observations matching you search criteria.<BR><BR>"
				else
					strAffectedProductRows = ""
					set rs2 = server.CreateObject("ADODB.recordset")
					rs2.open "spListSIAffectedProductsAll " & trim(clng(rs("ObservationID"))) ,cnExcalibur
					if not (rs2.eof and rs2.bof) then
						laststatus = ""
						strProductList = ""
						do while not rs2.eof
							if laststatus <> rs2("state") & "" then
								if laststatus <> "" then
									strAffectedProductRows = strAffectedProductRows & mid(strProductList,2) & "</td></tr>"
									strProductList = ""
								end if
								strAffectedProductRows = strAffectedProductRows & "<tr><td valign=top nowrap>" & replace(rs2("state") & "","*","") & ":&nbsp;&nbsp;</td><td colspan=5>"
							end if
							laststatus = rs2("state") & ""
							if trim(lcase(rs2("product") & "")) = trim(lcase(rs("PrimaryProduct") & "")) then
								strProductList = strProductList & ", <b>" & replace(rs2("product") ," ","&nbsp;") & "</b>"
							else
								strProductList = strProductList & ", " & replace(rs2("product") ," ","&nbsp;")
							end if
							rs2.movenext
						loop
						strAffectedProductRows = strAffectedProductRows & mid(strProductList,2) & "</td></tr>"'
					end if
					rs2.close

					if trim(CurrentUserPartner) = "1" then
						response.Write "<b>Observation ID:</b> <a target=OTS href=""https://si.austin.hp.com/si/?ObjectType=6&Object=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a><BR>"
					else
						response.Write "<FONT face=verdana size=2><b>Observation ID:</b> <a target=OTS href=""https://prp-si.corp.hp.com/si/?ObjectType=6&Object=" & rs("observationid") & """>" & rs("ObservationID") & "</a></font><BR>"
					end if
					Response.Write "<TABLE cellspacing=0 Border=1 width=""100%""><TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Affected Products:</b></td></tr><tr><td colspan=3><table>" & strAffectedProductRows & "</table></TD></TR></table>"

				end if
			elseif trim(strSection) = "4" then
'				if currentuserid = 31 then
'					response.write strPageParams
'					response.Flush
'				end if
				if strPageParams <> "" then
					ParamArray4 = split(PageParamsArray(SectionCount),"^")
				end if

				'Set Defaults
				EndDateParam = formatdatetime(Now()-1,vbshortdate) 'cdate("3/4/2011")
				WeeksParam = 16
				if datediff("ww","1/1/2011",EndDateParam) < WeeksParam then
					WeeksParam = datediff("ww","1/1/2011",EndDateParam)
				end if
				LegendParam = 1 '0=none, 1=right, 2=bottom
				blnActivityGraphParam = true
				blnTotalBacklogParam = true
				strTitleParam = ""
				Widthparam = 2
				HeightParam = 2
				ItemsPerGridLineParam = 0
				strChartType = "LineMarkers"

				'Load values
				if strPageParams <> "" then
					for i = 0 to ubound(ParamArray4)
						if trim(ParamArray4(i)) <> "" then
							Select case i
								case 0
									EndDateParam = cdate(trim(ParamArray4(i)))
								case 1
									WeeksParam = clng(ParamArray4(i))
									if datediff("ww","1/1/2011",EndDateParam) < WeeksParam then
										WeeksParam = datediff("ww","1/1/2011",EndDateParam)
									end if
								case 2
									LegendParam = clng(ParamArray4(i))
								case 3
									blnActivityGraphParam = cbool(ParamArray4(i))
								case 4
									strTitleParam = trim(ParamArray4(i))
								case 5
									Widthparam = clng(ParamArray4(i))
								case 6
									Heightparam = clng(ParamArray4(i))
								case 7
									ItemsPerGridLineParam = clng(ParamArray4(i))
								case 10
									strChartType = ParamArray4(i)
								case 15
									blnTotalBacklogParam = CBool(ParamArray4(i))
							end select
						end if
					next
				end if
				response.write "<BR>"
			elseif trim(strSection) = "5" then
				Section5Title = "Observations By Priority"
				strSQl5 = "Select case when Priority in (0,1) then 1 else Priority end as Priority, count(case when state in (" & statesUI & ") then 1 end) as UI,count(case when state in (" & statesFIP & ") then 1 end) as FIP,count(case when state not in (" & statesUI & "," & statesFIP & ") then 1 end) as Retest, Count(1) as backlog " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQl5 = strSQl5 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQl5 = strSQl5 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQL5 = strSQL5 & " and status <> 'closed' "
					Section5Title = "Open Observations by Priority"
				end if
				strSQl5 = strSQl5 & "Group By case when Priority in (0,1) then 1 else Priority end " & _
									"order by Priority "
'				response.write strSQl5
'				response.write "<hr>"

				rs.open strSQL5,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=5><b>" & Section5Title & "</b></td></tr>"
					response.write "<td align=center><b>Priority</b></td>"
					response.write "<td align=center><b>Under Investigation *</b></td>"
					response.write "<td align=center><b>Fix in Progress **</b></td>"
					response.write "<td align=center><b>Retest ***</b></td>"
					response.write "<td align=center><b>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section5Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					if trim(rs("Priority")) = "1" then
						response.write "<tr>"
						response.write "<td align=center>P0/P1</td>"
						response.write "<td align=center>" & rs("UI") & "</td>"
						response.write "<td align=center>" & rs("FIP") & "</td>"
						response.write "<td align=center>" & rs("Retest") & "</td>"
						response.write "<td align=center>" & rs("backlog") & "</td>"
						response.write "</tr>"
					elseif trim(rs("Priority")) = "2" then
						response.write "<tr>"
						response.write "<td align=center>P2</td>"
						response.write "<td align=center>" & rs("UI") & "</td>"
						response.write "<td align=center>" & rs("FIP") & "</td>"
						response.write "<td align=center>" & rs("Retest") & "</td>"
						response.write "<td align=center>" & rs("backlog") & "</td>"
						response.write "</tr>"
					elseif trim(rs("Priority")) = "3" then
						response.write "<tr>"
						response.write "<td align=center>P3</td>"
						response.write "<td align=center>" & rs("UI") & "</td>"
						response.write "<td align=center>" & rs("FIP") & "</td>"
						response.write "<td align=center>" & rs("Retest") & "</td>"
						response.write "<td align=center>" & rs("backlog") & "</td>"
						response.write "</tr>"
					elseif trim(rs("Priority")) = "4" then
						response.write "<tr>"
						response.write "<td align=center>P4</td>"
						response.write "<td align=center>" & rs("UI") & "</td>"
						response.write "<td align=center>" & rs("FIP") & "</td>"
						response.write "<td align=center>" & rs("Retest") & "</td>"
						response.write "<td align=center>" & rs("backlog") & "</td>"
						response.write "</tr>"
					end if
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "6" then
				Section6Title = "Observations By Developer"
				strSQL6 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, DeveloperName " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQL6 = strSQL6 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQL6 = strSQL6 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQL6 = strSQL6 & " and status <> 'closed' "
					Section6Title = "Open Observations by Developer"
				end if
				strSQL6 = strSQL6 & "Group By DeveloperName " & _
									"order by DeveloperName "
'				response.write strSQL6
'				response.write "<hr>"

				rs.open strSQL6,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section6Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>Developer</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section6Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><BR><BR>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("Developername") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "28" then
				Section28Title = "Observations By Owner"
				strSQL28 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, OwnerName " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQL28 = strSQL28 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQL28 = strSQL28 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQL28 = strSQL28 & " and status <> 'closed' "
					Section28Title = "Open Observations by Owner"
				end if
				strSQL28 = strSQL28 & "Group By OwnerName " & _
									"order by OwnerName "
				response.write strSQL28
				response.write "<hr>"

				rs.open strSQL28,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section28Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>Owner</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section28Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><BR><BR>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("Ownername") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "8" then
				Section8Title = "Observations By Deliverable"
				strSQL8 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, Component " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQL8 = strSQL8 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQL8 = strSQL8 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQL8 = strSQL8 & " and status <> 'closed' "
					Section8Title = "Open Observations by Deliverable"
				end if
				strSQL8 = strSQL8 & "Group By Component " & _
									"order by Component "
'				response.write strSQL8
'				response.write "<hr>"

				rs.open strSQL8,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section8Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>Deliverable</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section8Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("Component") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "9" then
				Section9Title = "Observations By Sub System"
				strSQl9 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, SubSystem " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQl9 = strSQl9 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQl9 = strSQl9 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQl9 = strSQl9 & " and status <> 'closed' "
					Section9Title = "Open Observations by Sub System"
				end if
				strSQl9 = strSQl9 & "Group By SubSystem " & _
									"order by SubSystem "
'				response.write strSQl9
'				response.write "<hr>"

				rs.open strSQl9,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=9><b>" & Section9Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>SubSystem</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section9Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("SubSystem") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "10" then
				Section10Title = "Observations By State"
				strSQl10 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, State " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQl10 = strSQl10 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQl10 = strSQl10 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQl10 = strSQl10 & " and status <> 'closed' "
					Section10Title = "Open Observations by State"
				end if
				strSQl10 = strSQl10 & "Group By State " & _
									"order by State "
'				if currentuserid = 31 then
'					response.write strSQl10
'					response.write "<hr>"
'				end if
				rs.open strSQl10,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=9><b>" & Section10Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>State</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section10Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("State") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "13" then
				Section13Title = "Observations By Status"
				strSQl13 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, Status, case Status when 'Open' then 1 when 'Pending EA' then 2 else 3 end as statusid " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQl13 = strSQl13 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQl13 = strSQl13 & " where " & mid(strSQLFilters,5)
				end if
				strSQl13 = strSQl13 & "Group By Status " & _
									"order by Statusid "
'				response.write strSQl13
'				response.write "<hr>"

				rs.open strSQl13,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=9><b>" & Section13Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>Status</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section13Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("Status") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "12" then
				Section12Title = "Observations By Component PM"
				strSQl12 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, ComponentPMName " & _
					" " & strSQLTables & " "
				if trim(strSQLJoins) <> "" then
					strSQl12 = strSQl12 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
				else
					strSQl12 = strSQl12 & " where " & mid(strSQLFilters,5)
				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQl12 = strSQl12 & " and status <> 'closed' "
					Section12Title = "Open Observations by Component PM"
				end if
				strSQl12 = strSQl12 & "Group By ComponentPMName " & _
									"order by ComponentPMName "
'				if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'					response.write strSQl12
'					response.flush
'					response.write "<hr>"
'				end if

				rs.open strSQl12,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=9><b>" & Section12Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>Component PM</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section12Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("ComponentPMName") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "11" then
				Section11Title = "Observations By Core Team"
				strSQl11 = "Select count(case when state not in (" & statesUI & "," & statesFIP & ") and Priority IN (0,1) then 1 end) as Retest, count(case when state in (" & statesFIP & ") and Priority in(0,1) then 1 end) as FIP, count(case when state in (" & statesUI & ") and priority in (0,1) then 1 end) as UI, coalesce(sum(case when Priority in (0,1) then 1 end),0) as P1,coalesce(sum(case when Priority =2 then 1 end),0) as P2,coalesce(sum(case when Priority =3 then 1 end),0) as P3,coalesce(sum(case when Priority =4 then 1 end),0) as P4, Count(1) as backlog, dbo.[ufn_GetCoreTeamNameFromComponentName](o.Component) as CoreTeam " & _
					" " & strSQLTables & " "
'				if not blnExcaliburDataRequired then
'					if instr(strSQLTables,"ct.name as coreteam" )=0  then
''						strSQL11 = strSQL11 & ", (Select ct.id as CoreTeamID, ct.name as coreteam, v.id as VersionID, v.otspartnumber from prs.dbo.deliverablecoreteam ct with (NOLOCK), prs.dbo.deliverableroot r with (NOLOCK), prs.dbo.deliverableversion v with (NOLOCK) where v.deliverablerootid = r.id and ct.id = r.coreteamid union Select 0 as CoreTeamID, 'None' as CoreTeam,ID as versionid, OTSPartNumber from prs.dbo.OTSComponent with (NOLOCK)) ex "
'						strSQL11 = strSQL11 & ", (Select ct.id as CoreTeamID, ct.name as coreteam, v.id as VersionID from prs.dbo.deliverablecoreteam ct with (NOLOCK), prs.dbo.deliverableroot r with (NOLOCK), prs.dbo.deliverableversion v with (NOLOCK) where v.deliverablerootid = r.id and ct.id = r.coreteamid union Select 0 as CoreTeamID, 'None' as CoreTeam, ID as versionid from prs.dbo.OTSComponent with (NOLOCK) union Select 0 as CoreTeamID, 'None' as CoreTeam,0 as versionid) ex "
'					end if
'				end if
'				if not blnExcaliburDataRequired then
'					if trim(strSQLJoins) <> "" then
'						if instr(strSQLJoins,"ex.versionid")=0 then
''							strSQl11 = strSQl11 & " where " & mid(strSQLJoins,5) & " and ex.versionid = o.ExcaliburNumber " & strSQLFilters
'							strSQl11 = strSQl11 & " where " & mid(strSQLJoins,5) & " and ((o.DivisionID = 6 and ex.versionid = o.ExcaliburNumber) or (o.DivisionID <> 6 and ex.versionid = 0)) " & strSQLFilters
'						end if
'					else
''							strSQl11 = strSQl11 & " where ex.versionid = o.ExcaliburNumber and " & mid(strSQLFilters,5)
'						if instr(strSQLJoins,"ex.versionid")=0 then
'							strSQl11 = strSQl11 & " where ((o.DivisionID = 6 and ex.versionid = o.ExcaliburNumber) or (o.DivisionID <> 6 and ex.versionid = 0)) and " & mid(strSQLFilters,5)
'						end if
'					end if
'				else
					if trim(strSQLJoins) <> "" then
						strSQl11 = strSQl11 & " where " & mid(strSQLJoins,5) & " " & strSQLFilters
					else
						strSQl11 = strSQl11 & " where " & mid(strSQLFilters,5)
					end if
'				end if
				if not FilterContains(strSQLFilters,"status","closed") then
					strSQl11 = strSQl11 & " and status <> 'closed' "
					Section11Title = "Open Observations by Core Team"
				end if
				strSQl11 = strSQl11 & " Group By [dbo].[ufn_GetCoreTeamNameFromComponentName](o.Component) " & _
									" order by CoreTeam "

'				if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'					response.write strSQl11
'					response.flush
'					response.write "<hr>"
'				end if
				rs.open strSQl11,cnSIO,adOpenStatic
				if not (rs.eof and rs.bof) then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=9><b>" & Section11Title & "</b></td></tr>"
					response.write "<tr><td align=left><b><BR>CoreTeam</b></td>"
					response.write "<td align=center><b>Under Investigation<BR>P0/P1</b> *</td>"
					response.write "<td align=center><b>Fix In Progress<BR>P0/P1</b> **</td>"
					response.write "<td align=center><b>Retest<br>P0/P1</b> ***</td>"
					response.write "<td align=center><b><BR>P0/P1</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P2&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P3&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>&nbsp;&nbsp;P4&nbsp;&nbsp;</b></td>"
					response.write "<td align=center><b><BR>Total</b></td>"
					response.write "</tr>"
				else
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>" & Section11Title & "</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				end if
				do while not rs.EOF
					response.write "<tr>"
					response.write "<td align=left>" & rs("CoreTeam") & "</td>"
					response.write "<td align=center>" & rs("ui") & "</td>"
					response.write "<td align=center>" & rs("fip") & "</td>"
					response.write "<td align=center>" & rs("retest") & "</td>"
					response.write "<td align=center>" & rs("p1") & "</td>"
					response.write "<td align=center>" & rs("p2") & "</td>"
					response.write "<td align=center>" & rs("P3") & "</td>"
					response.write "<td align=center>" & rs("P4") & "</td>"
					response.write "<td align=center>" & rs("backlog") & "</td>"
					response.write "</tr>"
					rs.MoveNext
				loop
				if not (rs.eof and rs.bof) then
					response.write "</table><br><br>"
				end if
				rs.Close
			elseif trim(strSection) = "-1" then 'SQLCheck
				strCheckSQl = "Select Count(1) " & strSQlTables & " where "

				if trim(strSQLJoins) <> "" then
					strCheckSQl = strCheckSQl & " " & mid(strSQLJoins,5) & " "
				end if

				if trim(strSQLFilters) <> "" and trim(strSQLJoins) = "" then
					strCheckSQl = strCheckSQl & " " & mid(strSQLFilters,5)
				elseif trim(strSQLFilters) <> "" then
					strCheckSQl = strCheckSQl & strSQLFilters
				end if
				on error resume next
				rs.open strCheckSQl,cnSIO
				if err.number <> 0 then
					response.write "SyntaxError: " & err.description
				else
					response.write "SyntaxOK"
					response.write request.Form
				end if
				rs.Close
				on error goto 0
			elseif trim(strSection) = "-2" then 'History
				numObservations = 0
				numRows = 0
				currentObservationId = ""
				strHistorySQL = "select ObservationId,Log_Date,Action_Type,Updated_By,Log_Summary " & strSQlTables &_
					" , dbo.History h with (nolock) " & " where "
				if trim(strSQLJoins) <> "" then
					strHistorySQl = strHistorySQl & " " & mid(strSQLJoins,5) & " "
				end if

				if trim(strSQLFilters) <> "" and trim(strSQLJoins) = "" then
					strHistorySQl = strHistorySQl & " " & mid(strSQLFilters,5)
				elseif trim(strSQLFilters) <> "" then
					strHistorySQl = strHistorySQl & strSQLFilters
				end if
				strHistorySQl = strHistorySQl & " and h.object_id = observationid order by observationid, log_date desc; "
'				if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'					response.write "<p>" & server.HTMLEncode(strHistorySQl) & "</p>"
'					response.flush
'				end if
				rs.open strHistorySQl,cnSIO
				if rs.eof and rs.bof then
					response.write "Unable to find observations matching you search criteria.<BR><BR>"
				else
					response.write "<TABLE cellspacing=0 Border=1 width=""100%"">" &_
						"<tr>" &_
						"<td>Observation&nbsp;ID</td>" &_
						"<td>Date</td>" &_
						"<td>User</td>" &_
						"<td>Type</td>" &_
						"<td>Summary</td>" &_
						"</tr>"
					do while not rs.eof
						numRows = numRows + 1
						if rs("ObservationID") & "" <> currentObservationId & "" then
							numObservations = numObservations + 1
							currentObservationId = rs("ObservationID")
						end if
						response.write "<tr>" &_
							FormatDataCell( "observationid", rs("ObservationID") & "", xmlformat,"" ) &_
							FormatDataCell( "log_date", rs("Log_Date") & "", xmlformat,"" ) &_
							FormatDataCell( "updated_by", rs("Updated_By") & "", xmlformat,"" ) &_
							FormatDataCell( "action_type", rs("action_type") & "", xmlformat,"" ) &_
							FormatDataCell( "log_summary", rs("Log_Summary") & "", xmlformat,"" ) &_
							"</tr>"
						rs.movenext
					loop
					response.write "</table>"
					rs.close
					if xmlformat=0 then
						response.write "<BR><BR>Observations Displayed: " & numObservations &_
							"<br>History Records Displayed: " & numRows
					end if
				end if
			elseif trim(strSection) = "2" then 'Working Notes
				rs.open strSQl,cnSIO
				if rs.eof and rs.bof then
					response.write "Unable to find observations matching you search criteria.<BR><BR>"
				else
					set rs2 = server.CreateObject("ADODB.recordset")
					rs2.open "spGetSIWorkingNotes " & rs("ObservationID"),cnExcalibur
					strWorkingNotes = ""
					i = 0
					do while not rs2.eof
						if i > 0 then
							strWorkingNotes = strWorkingNotes & "<HR>"
						end if
						i=i+1
						strWorkingNotes = strWorkingNotes & rs2("Log Date") & " by " & rs2("User Name") & "<BR style=""mso-data-placement: same-cell""/>"
						strWorkingNotes = strWorkingNotes & rs2("New_Value") & ""
						rs2.movenext
					loop

					rs2.close
					set rs2 = nothing
						if trim(CurrentUserPartner) = "1" then
							response.Write "<b>Observation ID:</b> <a target=OTS href=""https://si.austin.hp.com/si/?ObjectType=6&Object=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a><BR>"
						else
							response.Write "<FONT face=verdana size=2><b>Observation ID:</b> <a target=OTS href=""https://prp-si.corp.hp.com/si/?ObjectType=6&Object=" & rs("observationid") & """>" & rs("ObservationID") & "</a></font><BR>"
						end if
					Response.Write "<TABLE cellspacing=0 Border=1 width=""100%""><TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Updates:</b></td></tr><tr><td colspan=3>" & strWorkingNotes & "</TD></TR></table>"
				end if
			elseif trim(strSection) = "1" then 'Detailed Tables
'				if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'					response.write strSQL
'					response.Flush
'				end if

				rs.open strSQl,cnSIO
				if rs.eof and rs.bof then
					response.write "Unable to find observations matching you search criteria.<BR><BR>"
				else
					do while not rs.EOF
						RowsDisplayed = RowsDisplayed + 1

						if trim(CurrentUserPartner) = "1" then
							response.Write "<b>Observation ID:</b> <a target=OTS href=""https://si.austin.hp.com/si/?ObjectType=6&Object=" & rs("ObservationID") & """>" & rs("ObservationID") & "</a><BR>"
						else
							response.Write "<FONT face=verdana size=2><b>Observation ID:</b> <a target=OTS href=""https://prp-si.corp.hp.com/si/?ObjectType=6&Object=" & rs("observationid") & """>" & rs("ObservationID") & "</a></font><BR>"
						end if

						if trim(rs("LongDescription") & "") = "" then
							strLongDescription = "&nbsp;"
						elseif trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
							strLongDescription = ReplaceAndHTMLEncodeFrom(rs("longdescription") & "")
						else
							strLongDescription = ShortenField(rs("longdescription")& "",clng(strSavedLargeFields))
						end if

						if trim(rs("StepsToReproduce") & "") = "" then
							strStepsToReproduce = "&nbsp;"
						elseif trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
							strStepsToReproduce = ReplaceAndHTMLEncodeFrom(rs("StepsToReproduce") & "")
						else
							strStepsToReproduce = ShortenField(rs("StepsToReproduce")& "",clng(strSavedLargeFields))
						end if

						Response.Write "<b>Short Description: " & Server.HTMLEncode(rs("ShortDescription")) & "</b><BR>"
						Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"" style=""border: solid 1px black""><TR>"
						Response.Write "<TD><TABLE width=""100%"" ><TR><TD nowrap>Date Opened:</TD><TD align=left>" & rs("DateOpened") & "</TD></TR><TR><TD>Days Open:</TD><TD align=left>" & rs("DaysOpen") & "</TD></TR><TR><TD>Days In Current State:<TD align=left>" & rs("DaysInState") & "</TD></TR><TR><TD>Days With Current Owner:<TD align=left>" & rs("DaysCurrentOwner") & "</TD></TR><TR><TD>Target Date:<TD align=left>" & rs("TargetDate") & "&nbsp;</TD></TR></table></TD>"
						Response.Write "<TD valign=top><TABLE width=""100%""><TR><TD nowrap>Product:</TD><TD align=left>" & rs("PrimaryProduct") & "</TD></TR><TR><TD>Sub&nbsp;System:</TD><TD align=left>" & rs("subsystem") & "</TD></TR><TR><TD>Component:</font><TD align=left>" & rs("component") & "</TD></TR><TR><TD>Version:</TD><TD align=left>" & rs("ComponentVersion") & "&nbsp;</TD></TR><TR><TD>Feature:</TD><TD align=left>" & rs("Feature") & "&nbsp;</TD></TR></table></TD>"
						Response.Write "<TD valign=top><TABLE width=""100%"" ><TR><TD nowrap>Priority:</TD><TD align=left>" & rs("Priority") & "</TD></TR><TR><TD>Status:</TD><TD align=left>" & rs("Status") & "</TD></TR><TR><TD>State:<TD align=left>" & rs("state") & "</TD></TR><TR><TD>Frequency:</TD><TD align=left>" & rs("Frequency") & "&nbsp;</TD></TR><TR><TD>Milestone:</TD><TD align=left>" & rs("GatingMilestone") & "&nbsp;</TD></TR></table></TD></TR>"
						Response.Write "<TR><TD colspan=3><TABLE width=""100%""><TR><TD align=center nowrap><u>Originator</u><BR>" & propercase(rs("OriginatorName")) & "<BR>" & rs("OriginatorGroup") & "</TD><TD align=center nowrap><u>Component PM</u><BR>" & propercase(rs("ComponentPMName")) & "<BR>" & rs("ComponentPMGroup") & "</TD><TD align=center nowrap><u>Developer</u><BR>" & propercase(rs("DeveloperName")) & "<BR>" & rs("DeveloperGroup") & "</TD><TD align=center nowrap><u>Owner</u><BR>" & propercase(rs("OwnerName")) & "<BR>" & rs("OwnerGroup") & "</TD></tr></table></TD></TR>"
'						Response.Write "<TR><TD colspan=3><table width=100% bgcolor=WhiteSmoke><TR><TD style=""BORDER-Bottom: LightGrey 1px solid;""><b>Long Description:</b></td></tr></table>" & strLongDescription & "</TD></TR>"
						Response.Write "<TR><TD bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;"" colspan=3><b>Long Description:</b></TD></TR>"
						response.write "<TR><TD colspan=3>" & strLongDescription & "</td></tr>"
'						Response.Write "<TR><TD colspan=3><table width=100% bgcolor=WhiteSmoke><TR><TD style=""BORDER-Bottom: LightGrey 1px solid;""><b>Steps to Reproduce:</b></td></tr></table>" & strStepsToReproduce &"</TD></TR>"
						Response.Write "<TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Steps to Reproduce:</b></td></tr>"
						response.write "<TR><TD colspan=3>" & strStepsToReproduce & "</td></tr>"

						if trim(rs("SourceSystem") & "") = "SI" then
							set rs2 = server.CreateObject("ADODB.recordset")
							rs2.open "spListSIAffectedProducts " & trim(clng(rs("ObservationID"))) ,cnExcalibur
							if not (rs2.eof and rs2.bof) then
								Response.Write "<TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Other Affected Products:</b></td></tr><tr><td colspan=3><table>"
								laststatus = ""
								strProductList = ""
								do while not rs2.eof
									if laststatus <> rs2("state") & "" then
										if laststatus <> "" then
											response.write mid(strProductList,2) & "</td></tr>"
											strProductList = ""
										end if
										response.write "<tr><td valign=top nowrap>" & replace(rs2("state") & "","*","") & ":&nbsp;&nbsp;</td><td colspan=5>"
									end if
									laststatus = rs2("state") & ""
									strProductList = strProductList & ", " & replace(rs2("product") ," ","&nbsp;")
									rs2.movenext
								loop
								response.write mid(strProductList,2) & "</td></tr></table></TD></TR>"'
							end if
							rs2.close
							set rs2 = nothing
'							set rs2 = server.CreateObject("ADODB.recordset")
'							rs2.open "spGetSIWorkingNotes " & rs("ObservationID"),cnExcalibur
'							strWorkingNotes = ""
'							i = 0
'							do while not rs2.eof
'								if i > 0 then
'									strWorkingNotes = strWorkingNotes & "<HR>"
'								end if
'								i=i+1
'								strWorkingNotes = strWorkingNotes & rs2("Log Date") & " by " & rs2("User Name") & "<BR style=""mso-data-placement: same-cell""/>"
'								strWorkingNotes = strWorkingNotes & rs2("New_Value") & ""
'								rs2.movenext
'							loop
'
'							rs2.close
'							set rs2 = nothing
						end if
'						Response.Write "<TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Updates:</b></td></tr><tr><td colspan=3>" & strWorkingNotes & "</TD></TR>"
						Response.Write "<TR><TD colspan=3 bgcolor=WhiteSmoke style=""BORDER-Bottom: LightGrey 1px solid;""><b>Updates:</b></td></tr><tr><td colspan=3>"
						if request("cboFormat")= "Excel" or request("cboFormat")= "Word" then
							response.write FormatUpdates(rs("Updates") & "",0,rs("Observationid"))
						else
							response.write FormatUpdates(rs("Updates") & "",strSavedLargeFields,rs("Observationid"))
						end if

						Response.Write "&nbsp;</TD></TR></TABLE><BR><BR>"
						rs.MoveNext
					loop
					rs.Close
				end if
			elseif trim(strSection) = "0" or trim(strSection) = "3" then 'Summary Table
'				if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then
'					response.write "<br/>columns=" & strSaveColumns & "<br/>"
'					response.write strSQL
'					response.Flush
'				end if

				if blnStatusReport then
					rs.open strSQlOpen,cnSIO
				else
					rs.open strSQl,cnSIO
				end if
				if rs.eof and rs.bof and ucase(request("cboFormat")) <> "XML" then
					response.write "<table width=""100%"" cellspacing=0 cellpadding=2 bgcolor=white bordercolor=black border=1 style=""border: solid 1px black"">"
					response.write "<tr bgcolor=gainsboro><td align=center colspan=8><b>Summary Report</b></td></tr>"
					response.write "<tr><td>No observations match your search criteria.</td></tr></table><br><br>"
				else
					if trim(strSection) <> "3" then
						if xmlformat = 2 then
							response.write "<root>" & vblf
						elseif xmlformat = 1 then
							response.write "<observations>" & vblf
						else
							if blnStatusReport then
								response.write "<table id=tblOTS border =1 cellpadding=2 cellspacing=0 bgcolor=white bordercolor=black>"
							else
								response.write "<table id=tblOTS border =1 cellpadding=2 cellspacing=0 bgcolor=ivory>"
							end if
							response.write strColumnheaders
						end if
					else
'						response.write "<form id=""frmEmail"" method=""post"" target=_blank action=""../common/sendemail_mattH.asp"">"
						response.write "<form id=""frmEmail"" method=""post"" target=frameSendEmail action=""../common/sendemail.asp"">"
						response.write "<input id=""cmdSend"" name=""cmdSend"" type=""button"" value=""Send Selected Emails"" onclick=""javascript:SendEmail();"" /><BR><BR>"
						response.write "<input type=""checkbox"" checked id=""chkCopyMe"" name=""chkCopyMe"" value=""1"">&nbsp;Copy me on these emails.</font><br><br>"
						response.write "<fieldset style=""background-color: lavender;"">"
						if trim(strSavedSubject) = "" then
							response.write "<table border=0 style=""width:100%""><tr><td><b>Subject: <font color=red>*</font></b></td><td width=""100%"" ><input style=""width:100%"" id=""txtSubject"" name=""txtSubject"" type=""text"" value=""" & server.HTMLEncode(strSavedSubject) & """></td></tr>"
						else
							response.write "<table border=0 style=""width:100%""><tr><td><b>Subject:</b></td><td width=""100%"" >" & server.HTMLEncode(strSavedSubject) & "<input style=""display:none"" id=""txtSubject"" name=""txtSubject"" type=""text"" value=""" & server.HTMLEncode(strSavedSubject) & """></td></tr>"
						end if
						if trim(strSavedNotes) = "" then
							response.write "<tr><td><b>Notes: <font color=red>*</font></b></td><td><input style=""width:100%"" id=""txtNotes"" name=""txtNotes"" type=""text"" value=""" & server.HTMLEncode(replace(strSavedNotes,vbcrlf,"<BR/>")) & """></td></tr>"
						else
							response.write "<tr><td><b>Notes:</b></td><td>" & server.HTMLEncode(replace(strSavedNotes,vbcrlf,"<BR/>")) & "<input style=""display:none"" id=""txtNotes"" name=""txtNotes"" type=""text"" value=""" & server.HTMLEncode(replace(strSavedNotes,vbcrlf,"<BR/>")) & """></td></tr>"
						end if
						response.write "<tr><td><b>Emails Sent:</b></td><td>Each person selected below will receive a single email containing only their observations.</td></tr>"
						response.write "<tr><td><b>Observations:</b></td><td>The table of selected observations owned by each person will be included in the body of the email.</td></tr>"
						response.write "<tr id=ClosedWarningRow style=""display:none""><td><b>Closed:</b></td><td><b>This report includes <font color=red>closed</font> observations.</b></td></tr></table>"
						response.write "</fieldset><BR>"
					end if
				end if

				strLastOwner = ""
				EmailTableCount=0
				ClosedCount = 0
				do while not rs.EOF
					if trim(strSection) = "3" and trim(strLastOwner) <> trim(rs("Owner")) then
						if strLastOwner <> "" then
							response.write "</table><br>"
							response.write "<textarea id=txtEmailTable" & trim(EmailTableCount) & " name=txtEmailTable" & trim(EmailTableCount) & " style=""display:none;width:100%""></textarea>"
						end if

						strManagerName = ""
						strManagerEmail = ""
						set rs2 = server.CreateObject("ADODB.recordset")
						rs2.Open "spGetManagerInfoByEmail '" & trim(rs("Owner") & "") & "'",cnExcalibur,adOpenStatic
						if not(rs2.eof and rs2.BOF) then
							strManagerName = rs2("Name") & ""
							strManagerEmail = lcase(rs2("Email") & "")
						end if
						rs2.Close
						set rs2 = nothing
						if trim(rs("OwnerID")) = "-1000" then
							response.write "<BR>These observations were owned by Remedy users that do not have an Sudden Impact account."
						else
							response.write "<BR><input style=""height:16px;"" checked id=""chkTo"" name=""chkTo"" type=""checkbox"" value=""" & trim(EmailTableCount) & """>&nbsp;<b>Send To:</b> " & longName(rs("OwnerName")) & " (" & trim(rs("OwnerGroup") & "") & ")" & "<br>"
							response.write "<input style=""display:none"" id=""txtTo" & trim(EmailTableCount) & """ name=""txtTo" & trim(EmailTableCount) & """ type=""text"" value=""" & trim(rs("Owner") & "") & """>"
							if trim(strManagerEmail) <> "" then
								response.write "<input style=""height:16px;"" id=""chkCC" & trim(EmailTableCount) & """ name=""chkCC" & trim(EmailTableCount) & """ type=""checkbox"" value=""" & trim(strManagerEmail) & """>&nbsp;<b>CC Manager:</b> " & longname(strManagerName) & "<br>" '& " - " & strManagerEmail
							end if
						end if
						EmailTableCount = EmailTableCount + 1
						response.write "<table id=tblOTS" & EmailTableCount & " style=""width:100%"" border =1 cellpadding=2 cellspacing=0 bgcolor=ivory>"
						if trim(strLastOwner) <> trim(rs("Owner")) then
							response.write strColumnHeaders
						end if
						strLastOwner = rs("Owner")
					end if
					if xmlformat=2 then
						response.write "<row "
					elseif xmlformat=1 then
						response.write "<observation id=""" & rs("ObservationID") & """>" & vblf
					else
						response.write "<tr>"
					end if
					if rs("Status")="Closed" then
						ClosedCount = ClosedCount + 1
					end if
					RowsDisplayed = RowsDisplayed + 1
					for each strColumn in ColumnArray
						Select case lcase(replace(strColumn," ",""))
						case "searchrank"
							response.write FormatDataCell( "searchrank", rs("searchrank") & "", xmlformat,"" )
						case "observationid"
							if xmlformat <> 0 then
								response.write formatdatacell("observationid",rs("observationid") & "",xmlformat,"" )
							else
								if trim(rs("ObservationID") & "" ) = "" then
									response.write "<td>&nbsp;</td>"
								else
									dim strUrl
									strUrl = createObsDetailUrlFor(CLng(rs("ObservationID")), CBool(trim(currentuserpartner)="1"), CBool(strPageParams="macro"))
									if request("cboFormat") = "Excel" or request("cboFormat") = "Word" then
										if strPageParams = "macro" then
											response.write "<td>=HYPERLINK(""" & strUrl & """,""" & rs("ObservationID") & """)</td>"
										else
											response.write "<td><a target=_blank href=""" & strUrl & """>" & rs("ObservationID") & "</a></td>"
										end if
									elseif trim(strSection) = "3" or strPageParams = "macro" then
										response.write "<td><a target=_blank href=""" & strUrl & """>" & rs("ObservationID") & "</a></td>"
									else
										response.write "<td><a onmousemove=""SaveMouseCoordinates();"" href=""javascript: ShowIDMenu(" & clng(rs("ObservationID")) & "," & clng(CurrentUserPartner) & ");"">" & rs("ObservationID") & "</a></td>"
									end if
								end if
							end if
						case "primaryproduct"
							response.write FormatDataCell( "primaryproduct", rs("PrimaryProduct") & "", xmlformat,"" )
						case "state"
							response.write FormatDataCell( "state", rs("State") & "", xmlformat,"" )
						case "component"
							response.write FormatDataCell( "component", rs("Component") & "", xmlformat,"" )
						case "priority"
							response.write FormatDataCell( "priority", rs("Priority") & "", xmlformat,"" )
						case "owner"
							response.write FormatDataCell( "owner", rs("OwnerName") & "", xmlformat,"" )
						case "failedfixes"
							response.write FormatDataCell( "failedfixes", rs("FailedFixes") & "", xmlformat,"")
						case "shortdescription"
							if request("cboFormat")= "Excel" or request("cboFormat")= "Word" then
								response.write "<td>" & Server.HTMLEncode(rs("ShortDescription") & "") & "</td>"
							else
								response.write FormatDataCell( "shortdescription", rs("ShortDescription") & "", xmlformat,"" )
							end if
						case "productfamily"
							response.write FormatDataCell( "productfamily", rs("ProductFamily") & "", xmlformat,"" )
						case "feature"
							response.write FormatDataCell( "feature", rs("Feature") & "", xmlformat,"" )
						case "approver"
							response.write FormatDataCell( "approver", rs("approvername") & "", xmlformat,"" )
						case "approveremail"
							response.write FormatDataCell( "approveremail", rs("Approver") & "", xmlformat,"" )
						case "approvergroup"
							response.write FormatDataCell( "approvergroup", rs("ApproverGroup") & "", xmlformat,"" )
						case "approverlocation"
							response.write FormatDataCell( "approverlocation", rs("approverlocation") & "", xmlformat,"" )
						case "approvermanager"
							response.write FormatDataCell( "approvermanager", rs("approvermanager") & "", xmlformat,"" )
						case "componentpm"
							response.write FormatDataCell( "componentpm", rs("ComponentPMname") & "", xmlformat,"" )
						case "componentpmemail"
							response.write FormatDataCell( "componentpmemail", rs("ComponentPM") & "", xmlformat,"" )
						case "componentpmgroup"
							response.write FormatDataCell( "componentpmgroup", rs("ComponentPMGroup") & "", xmlformat,"" )
						case "componentpmlocation"
							response.write FormatDataCell( "componentpmlocation", rs("componentpmlocation") & "", xmlformat,"" )
						case "componentpmmanager"
							response.write FormatDataCell( "componentpmmanager", rs("componentpmmanager") & "", xmlformat,"" )
						case "closedinversion"
							response.write FormatDataCell( "closedinversion", rs("ClosedInVersion") & "", xmlformat,"" )
						case "affectedproduct"
							if xmlformat=0 then
								if trim(rs("Affectedproduct") & "" ) = "" then
									if request("cboFormat") = "Excel" or request("cboFormat") = "Word" then
										response.write "<td>&nbsp;</td>"
									else
										response.write "<td><a href=""javascript: ShowAffectedWindow(" & rs("ObservationID") & ");"">View</a></td>"
									end if
								elseif trim(rs("AffectedProductCount")) = "1" then
									response.write "<td>" & rs("Affectedproduct") & "</td>"
								elseif request("cboFormat") = "Excel" or request("cboFormat") = "Word" then
									response.write "<td>[Multiple]</td>"
								else
									response.write "<td><a href=""javascript: ShowAffectedWindow(" & rs("ObservationID") & ");"">View</a></td>"
								end if
							else
								if trim(rs("AffectedProductCount")) = "1" then
									response.write FormatDataCell( "affectedproduct", rs("Affectedproduct") & "", xmlformat,"" )
								else
									response.write FormatDataCell( "affectedproduct", "[Multiple]", xmlformat,"" )
								end if
							end if
						case "affectedstate"
							if xmlformat=0 then
								if trim(rs("Affectedstate") & "") = "" then
									if request("cboFormat") = "Excel" or request("cboFormat") = "Word" then
										response.write "<td>&nbsp;</td>"
									else
										response.write "<td><a href=""javascript: ShowAffectedWindow(" & rs("ObservationID") & ");"">View</a></td>"
									end if
								elseif trim(rs("AffectedProductCount")) = "1" then
									response.write "<td>" & rs("Affectedstate") & "</td>"
								elseif request("cboFormat") = "Excel" or request("cboFormat") = "Word" then
									response.write "<td>[Multiple]</td>"
								else
									response.write "<td><a href=""javascript: ShowAffectedWindow(" & rs("ObservationID") & ");"">View</a></td>"
								end if
							else
								if trim(rs("AffectedProductCount")) = "1" then
									response.write FormatDataCell( "affectedstate", rs("Affectedstate") & "", xmlformat,"" )
								else
									response.write FormatDataCell( "affectedstate", "[Multiple]", xmlformat,"" )
								end if
							end if
						case "approvalcheck"
							response.write FormatDataCell( "approvalcheck", rs("approvalcheck") & "", xmlformat,"" )
						case "componenttestlead"
							response.write FormatDataCell( "componenttestlead", rs("componenttestleadname") & "", xmlformat,"" )
						case "componenttestleademail"
							response.write FormatDataCell( "componenttestleademail", rs("componenttestlead") & "", xmlformat,"" )
						case "componenttestleadgroup"
							response.write FormatDataCell( "componenttestleadgroup", rs("componenttestleadgroup") & "", xmlformat,"" )
						case "componenttestleadlocation"
							response.write FormatDataCell( "componenttestleadlocation", rs("componenttestleadlocation") & "", xmlformat,"" )
						case "componenttestleadmanager"
							response.write FormatDataCell( "componenttestleadmanager", rs("componenttestleadmanager") & "", xmlformat,"" )
						case "componenttype"
							response.write FormatDataCell( "componenttype", rs("componenttype") & "", xmlformat,"" )
						case "componentversion"
							response.write FormatDataCell( "componentversion", rs("componentversion") & "", xmlformat,"" )
						case "componentpartno"
							response.write FormatDataCell( "componentpartno", rs("componentpartno") & "", xmlformat,"" )
						case "coreteam"
							response.write FormatDataCell( "coreteam", rs("coreteam") & "", xmlformat,"" )
						case "countcomponentassignments"
							response.write FormatDataCell( "countcomponentassignments", rs("countcomponentassignments") & "", xmlformat,"" )
						case "countownerassignments"
							response.write FormatDataCell( "countownerassignments", rs("countownerassignments") & "", xmlformat,"" )
						case "originator"
							response.write FormatDataCell( "originator", rs("Originatorname") & "", xmlformat,"" )
						case "originatoremail"
							response.write FormatDataCell( "originatoremail", rs("Originator") & "", xmlformat,"" )
						case "originatorgroup"
							response.write FormatDataCell( "originatorgroup", rs("Originatorgroup") & "", xmlformat,"" )
						case "originatorlocation"
							response.write FormatDataCell( "originatorlocation", rs("originatorlocation") & "", xmlformat,"" )
						case "originatormanager"
							response.write FormatDataCell( "originatormanager", rs("originatormanager") & "", xmlformat,"" )
'						case "history"
'							if xmlformat=0 then
'								if trim(rs("History") & "") = "" then
'									response.write "<td>&nbsp;</td>"
'								else
'									response.write "<td>"
'									response.write FormatUpdates(rs("History") & "",0,rs("Observationid"))
'									response.write "&nbsp;</td>"
'								end if
'							else
'								response.write FormatDataCell( "history", FormatUpdatesXML(rs("History") & "",strSavedLargeFields,rs("Observationid")), xmlformat,"" )
'							end if
						case "updates"
							if xmlformat=0 then
								if trim(rs("Updates") & "") = "" then
									response.write "<td>&nbsp;</td>"
								else
									response.write "<td>"
									if request("cboFormat")= "Excel" or request("cboFormat")= "Word" then
										response.write FormatUpdates(rs("Updates") & "",0,rs("Observationid"))
									else
										response.write FormatUpdates(rs("Updates") & "",strSavedLargeFields,rs("Observationid"))
									end if
									response.write "&nbsp;</td>"
								end if
							else
								response.write FormatDataCell( "updates", FormatUpdatesXML(rs("Updates") & "",strSavedLargeFields,rs("Observationid")), xmlformat,"" )
							end if
						case "updates2"
							if xmlformat=0 then
								if trim(rs("Updates2") & "") = "" then
									response.write "<td>&nbsp;</td>"
								else
									response.write "<td>"
									if request("cboFormat")= "Excel" or request("cboFormat")= "Word" then
										response.write FormatUpdates(rs("Updates2") & "",0,rs("Observationid"))
									else
										response.write FormatUpdates(rs("Updates2") & "",strSavedLargeFields,rs("Observationid"))
									end if
									response.write "&nbsp;</td>"
								end if
							else
								response.write FormatDataCell( "updates2", FormatUpdatesXML(rs("Updates2") & "",strSavedLargeFields,rs("Observationid")), xmlformat,"" )
							end if
						case "customerimpact"
							if xmlformat=0 then
								if trim(rs("customerimpact") & "") = "" then
									response.write "<td>&nbsp;</td>"
								elseif trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write "<td>" & ReplaceAndHTMLEncodeFrom(rs("customerimpact") & "") & "</td>"
								else
									response.write "<td>" & ShortenField(rs("customerimpact")& "",clng(strSavedLargeFields)) & "</td>"
								end if
							else
								if trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write FormatDataCell( "customerimpact", rs("customerimpact") & "", xmlformat,"" )
								else
									response.write FormatDataCell( "customerimpact", ShortenFieldXML(rs("customerimpact") & "",clng(strSavedLargeFields)), xmlformat,"" )
								end if
							end if
						case "dateopened"
							if trim(rs("dateopened") & "") = "" then
								response.write FormatDataCell( "dateopened", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "dateopened", myFormatDateTime(rs("dateopened")) & "", xmlformat,"" )
								else
									response.write FormatDataCell( "dateopened", formatdatetime(rs("dateopened"),vbshortdate) & "", xmlformat,"" )
								end if
							end if
						case "dateclosed"
							if trim(rs("dateclosed") & "") = "" then
								response.write FormatDataCell( "dateclosed", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "dateclosed", myFormatDateTime(rs("dateclosed")) & "", xmlformat,"" )
								else
									response.write FormatDataCell( "dateclosed", formatdatetime(rs("dateclosed"),vbshortdate) & "", xmlformat,"" )
								end if
							end if
						case "datemodified"
							if trim(rs("datemodified") & "") = "" then
								response.write FormatDataCell( "datemodified", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "datemodified", myFormatDateTime(rs("datemodified")) & "", xmlformat,"" )
								else
									response.write FormatDataCell( "datemodified", formatdatetime(rs("datemodified"),vbshortdate) & "", xmlformat,"" )
								end if
							end if
						case "daysinstate"
							response.write FormatDataCell( "daysinstate", rs("daysinstate") & "", xmlformat,"align=center" )
						case "dayscurrentowner"
							response.write FormatDataCell( "dayscurrentowner", rs("dayscurrentowner") & "", xmlformat,"align=center" )
						case "daysopen"
							response.write FormatDataCell( "daysopen", rs("daysopen") & "", xmlformat,"align=center" )
						case "developer"
							response.write FormatDataCell( "developer", rs("developername") & "", xmlformat,"" )
						case "developeremail"
							response.write FormatDataCell( "developeremail", rs("developer") & "", xmlformat,"" )
						case "developergroup"
							response.write FormatDataCell( "developergroup", rs("developerGroup") & "", xmlformat,"" )
						case "developerlocation"
							response.write FormatDataCell( "developerlocation", rs("developerlocation") & "", xmlformat,"" )
						case "developermanager"
							response.write FormatDataCell( "developermanager", rs("developermanager") & "", xmlformat,"" )
						case "division"
							response.write FormatDataCell( "division", rs("division") & "", xmlformat,"" )
						case "eadate"
							if trim(rs("eadate") & "") = "" then
								response.write FormatDataCell( "eadate", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "eadate", myFormatDateTime(rs("eadate")), xmlformat,"" )
								else
									response.write FormatDataCell( "eadate", formatdatetime(rs("eadate"),vbshortdate), xmlformat,"" )
								end if
							end if
						case "eanumber"
							response.write FormatDataCell( "eanumber", rs("eanumber") & "", xmlformat,"" )
						case "eastatus"
							response.write FormatDataCell( "eastatus", rs("eastatus") & "", xmlformat,"" )
						case "earliestproductmilestone"
							if trim(rs("EarliestProductMilestone") & "") = "" then
								response.write FormatDataCell( "earliestproductmilestone", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "earliestproductmilestone", myFormatDateTime(rs("EarliestProductMilestone")), xmlformat,"" )
								else
									response.write FormatDataCell( "earliestproductmilestone", formatdatetime(rs("EarliestProductMilestone"),vbshortdate), xmlformat,"" )
								end if
							end if
						case "frequency"
							response.write FormatDataCell( "frequency", rs("frequency") & "", xmlformat,"" )
						case "gatingmilestone"
							response.write FormatDataCell( "gatingmilestone", rs("gatingmilestone") & "", xmlformat,"" )
						case "impacts"
							response.write FormatDataCell( "impacts", rs("impacts") & "", xmlformat,"" )
						case "implementationcheck"
							response.write FormatDataCell( "implementationcheck", rs("implementationcheck") & "", xmlformat,"" )
						case "lastmodifiedby"
							response.write FormatDataCell( "lastmodifiedby", rs("lastmodifiedby") & "", xmlformat,"" )
						case "lastreleasetested"
							response.write FormatDataCell( "lastreleasetested", rs("lastreleasetested") & "", xmlformat,"" )
						case "localization"
							response.write FormatDataCell( "localization", rs("localization") & "", xmlformat,"" )
						case "longdescription"
							if xmlformat = 0 then
								if trim(rs("longdescription") & "") = "" then
									response.write "<td>&nbsp;</td>"
								elseif trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write "<td>" & ReplaceAndHTMLEncodeFrom(rs("longdescription") & "") & "</td>"
								else
									response.write "<td>" & ShortenField(rs("longdescription")& "",clng(strSavedLargeFields)) & "</td>"
								end if
							else
								if trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write FormatDataCell( "longdescription", rs("longdescription") & "", xmlformat,"" )
								else
									response.write FormatDataCell( "longdescription", ShortenFieldXML(rs("longdescription") & "",clng(strSavedLargeFields)), xmlformat,"" )
								end if
							end if
						case "odms"
							response.write FormatDataCell( "odms", rs("odms") & "", xmlformat,"" )
						case "onboard"
							response.write FormatDataCell( "onboard", rs("onboard") & "", xmlformat,"" )
						case "owneremail"
							response.write FormatDataCell( "owneremail", rs("owner") & "", xmlformat,"" )
						case "ownergroup"
							response.write FormatDataCell( "ownergroup", rs("ownergroup") & "", xmlformat,"" )
						case "ownerlocation"
							response.write FormatDataCell( "ownerlocation", rs("ownerlocation") & "", xmlformat,"" )
						case "ownermanager"
							response.write FormatDataCell( "ownermanager", rs("ownermanager") & "", xmlformat,"" )
						case "productsegment"
							response.write FormatDataCell( "productsegment", rs("productsegment") & "", xmlformat,"" )
						case "producttestlead"
							response.write FormatDataCell( "producttestlead", rs("producttestleadname") & "", xmlformat,"" )
						case "producttestleademail"
							response.write FormatDataCell( "producttestleademail", rs("producttestlead") & "", xmlformat,"" )
						case "producttestleadgroup"
							response.write FormatDataCell( "producttestleadgroup", rs("producttestleadgroup") & "", xmlformat,"" )
						case "producttestleadlocation"
							response.write FormatDataCell( "producttestleadlocation", rs("producttestleadlocation") & "", xmlformat,"" )
						case "producttestleadmanager"
							response.write FormatDataCell( "producttestleadmanager", rs("producttestleadmanager") & "", xmlformat,"" )
						case "productpm"
							response.write FormatDataCell( "productpm", rs("productpmname") & "", xmlformat,"" )
						case "productpmemail"
							response.write FormatDataCell( "productpmemail", rs("productpm") & "", xmlformat,"" )
						case "productpmgroup"
							response.write FormatDataCell( "productpmgroup", rs("productpmgroup") & "", xmlformat,"" )
						case "productpmlocation"
							response.write FormatDataCell( "productpmlocation", rs("productpmlocation") & "", xmlformat,"" )
						case "productpmmanager"
							response.write FormatDataCell( "productpmmanager", rs("productpmmanager") & "", xmlformat,"" )
						case "referencenumber"
							response.write FormatDataCell( "referencenumber", rs("referencenumber") & "", xmlformat,"" )
						case "releasefiximplemented"
							response.write FormatDataCell( "releasefiximplemented", rs("releasefiximplemented") & "", xmlformat,"" )
						case "reviewed"
							response.write FormatDataCell( "reviewed", rs("reviewed") & "", xmlformat,"" )
						case "sapartnumber"
							response.write FormatDataCell( "sapartnumber", rs("sapartnumber") & "", xmlformat,"" )
						case "severity"
							response.write FormatDataCell( "severity", rs("severityname") & "", xmlformat,"" )
						case "status"
							response.write FormatDataCell( "status", rs("status") & "", xmlformat,"" )
						case "stepstoreproduce"
							if xmlformat = 0 then
								if trim(rs("stepstoreproduce") & "") = "" then
									response.write "<td>&nbsp;</td>"
								elseif trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write "<td>" & ReplaceAndHTMLEncodeFrom(rs("stepstoreproduce") & "") & "</td>"
								else
									response.write "<td>" & ShortenField(rs("stepstoreproduce")& "",clng(strSavedLargeFields)) & "</td>"
								end if
							else
								if trim(strSavedLargeFields) = "" or trim(strSavedLargeFields) = "0" then
									response.write FormatDataCell( "stepstoreproduce", rs("stepstoreproduce") & "", xmlformat,"" )
								else
									response.write FormatDataCell( "stepstoreproduce", ShortenFieldXML(rs("stepstoreproduce") & "",clng(strSavedLargeFields)), xmlformat,"" )
								end if
							end if
						case "subsystem"
							response.write FormatDataCell( "subsystem", rs("subsystem") & "", xmlformat,"" )
						case "suppliers"
							response.write FormatDataCell( "suppliers", rs("suppliers") & "", xmlformat,"" )
						case "supplierversion"
							response.write FormatDataCell( "supplierversion", rs("supplierversion") & "", xmlformat,"" )
						case "targetdate"
							if trim(rs("targetdate") & "") = "" then
								response.write FormatDataCell( "targetdate", "", xmlformat,"" )
							else
								if (strPageParams = "macro") then
									response.write FormatDataCell( "targetdate", myFormatDateTime(rs("targetdate")), xmlformat,"" )
								else
									response.write FormatDataCell( "targetdate", formatdatetime(rs("targetdate"),vbshortdate), xmlformat,"" )
								end if
							end if
						case "testprocedure"
							response.write FormatDataCell( "testprocedure", rs("testprocedure") & "", xmlformat,"" )
						case "testescape"
							response.write FormatDataCell( "testescape", rs("testescape") & "", xmlformat,"" )
						case "tester"
							response.write FormatDataCell( "tester", rs("testername") & "", xmlformat,"" )
						case "testeremail"
							response.write FormatDataCell( "testeremail", rs("tester") & "", xmlformat,"" )
						case "testergroup"
							response.write FormatDataCell( "testergroup", rs("testergroup") & "", xmlformat,"" )
						case "testerlocation"
							response.write FormatDataCell( "testerlocation", rs("testerlocation") & "", xmlformat,"" )
						case "testermanager"
							response.write FormatDataCell( "testermanager", rs("testermanager") & "", xmlformat,"" )
						end select
					next
					if xmlformat=2 then
						response.write " />" & vblf
					elseif xmlformat=1 then
						response.write "</observation>" & vblf
					else
						response.write "</tr>"
					end if
					rs.MoveNext
				loop
				rs.Close
'				response.write "</tr>"
				if xmlformat=0 then
					response.write "</table><BR><BR>"
				end if
				if trim(strSection) = "3" then
					response.write "<textarea id=txtEmailTable" & trim(EmailTableCount) & " name=txtEmailTable" & trim(EmailTableCount) & " style=""display:none;width:100%""></textarea>"
'					response.write "<input id=""Text1"" name=""Text1"" type=""text"" value=""Test"">"
					response.write "</form>"
				end if
			end if
			SectionCount = SectionCount + 1
		next

'		response.write "<BR><BR>" & strSQL & "<BR>"
		if PageSections <> "2" and PageSections <> "7" and trim(request("txtReportSections")) <> "-1" and trim(request("txtReportSections")) <> "-2" then
			if PageSections = "3" then
				response.write "<input id=""txtClosedCount"" style=""display:none"" type=""text"" value=""" & ClosedCount & """>"
			elseif blnStatusReport then
				response.write "<br><hr>"
				response.write "<table><tr><td><font face=verdana size=""1"">* - </td><td nowrap>""Under Investigation"" numbers include: " & Server.HTMLEncode(statesUI) & "</td></tr>"
				response.write "<tr><td><font face=verdana size=""1"">** - </td><td nowrap>""Fix in Progress"" numbers include: " & Server.HTMLEncode(statesFIP) & "</td></tr>"
				response.write "<tr><td nowrap><font face=verdana size=""1"">*** - </td><td nowrap>""Retest"" numbers include all other open statuses.</font></td></tr>"
				response.write "</table><hr><br>"
				response.write "Report Generated " & Now & " (UTC-" & getTimezoneOffset()/60 & ")"
			end if
			if xmlformat=0 and not blnStatusReport and not (strPageParams = "macro" and request("cboFormat") = "Excel") then
				response.write "<BR><BR>Observations Displayed: " & RowsDisplayed
			end if
		end if
	end if

	function ReplaceAndHTMLEncodeFrom(someText)
		dim sameCellBreak
		dim workingText
		sameCellBreak = "<BR style=""mso-data-placement: same-cell""/>"
		workingText = someText
		workingText = replace(workingText, chr(160), " ")
		workingText = Server.HTMLEncode(workingText)
		workingText = replace(workingText, vbcr, sameCellBreak)
		workingText = replace(workingText, vblf, sameCellBreak)
		workingText = replace(workingText, vbcrlf, sameCellBreak)
		ReplaceAndHTMLEncodeFrom = workingText
	end function

	function createObsDetailUrlFor(observationID, isHP, isMacro)
		if CBool(isMacro) then
			if CBool(isHP) then
				createObsDetailUrlFor = "https://si.austin.hp.com/si/?ObjectType=6&Object=" & CLng(observationID)
			else
				createObsDetailUrlFor = "https://prp-si.corp.hp.com/si/?ObjectType=6&Object=" & CLng(observationID)
			end if
		else
			if CBool(isHP) then
'				createObsDetailUrlFor = "http://pulsarweb.usa.hp.com/Excalibur/search/ots/report_mattH.asp?txtReportSections=1&txtObservationID=" & CLng(observationID)
				createObsDetailUrlFor = "http://" & Application("Excalibur_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & CLng(observationID)
			else
'				createObsDetailUrlFor = "https://pulsarweb-pro.prp.ext.hp.com/Excalibur/search/ots/report_mattH.asp?txtReportSections=1&txtObservationID=" & CLng(observationID)
				createObsDetailUrlFor = "https://" & Application("Excalibur_ODM_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & CLng(observationID)
			end if
		end if
	end function

	function ListPrep(strText, Fieldname, ListType)
		dim strList
		dim strItem
		dim ItemArray
		dim SubItemArray
		dim strListProd, strListProdVer
		' special utf-8 character is making it through from default.asp wrongly
		' "ï¿½" 0xEFBFBD should be "â" 0xE28093
		' this problem showed up in the component field, but it likely could happen for any of them
		ItemArray = split(scrubsql(replace(strText, "ï¿½", "â")),",")
		strList = ""
		strListProd = ""
		strListProdVer = ""
		for each strItem in ItemArray
			if FieldName = "GatingMilestone" and trim(stritem)="Not Specified" then 'Text Items
				strList = strList & ",'' "
			elseif FieldName = "Component" then
				strList = strList & ",'" & replace(replace(trim(stritem),"'","''"),"|",",") & "' "
			elseif FieldName = "ProductAndVersion" then
				SubItemArray = split(strItem, "||")
				if ubound(SubItemArray) = 1 then
					strListProd = strListProd & ",'" & replace(trim(SubItemArray(0)), "'", "''") & "' "
					strListProdVer = strListProdVer & ",'" & replace(trim(SubItemArray(1)), "'", "''") & "' "
				end if
			elseif ListType = 0 then 'Text Items
				strList = strList & ",'" & replace(trim(stritem),"'","''") & "' "
			else 'Numeric Items
				strList = strList & "," & clng(stritem) & " "
			end if
		next
		if FieldName = "Assigned" then
			ListPrep = " and (OwnerID in (" & mid(strList,2) & ") or DeveloperID in (" & mid(strList,2) & ") or ComponentPMID in (" & mid(strList,2) & ") or OriginatorID in (" & mid(strList,2) & ") or ProductPMID in (" & mid(strList,2) & ") or ComponentTestLeadID in (" & mid(strList,2) & ") or ProductTestLeadID in (" & mid(strList,2) & ") or ApproverID in (" & mid(strList,2) & ") or TesterID in (" & mid(strList,2) & ") ) "
		elseif FieldName = "PrimaryProduct" and instr(strList,"Any Functional Test")> 0 then
			ListPrep = " and (" & FieldName & " in (" & mid(strList,2) & ") or ProductFamily like 'Func Tst%' ) "
		elseif lcase(FieldName) = "productandversion" then
			ListPrep = " and (PrimaryProduct in (" & mid(strListProdVer,2) & ") and ProductFamily in (" & mid(strListProd,2) & ")) "
		elseif lcase(FieldName) = "coreteamid" then
			ListPrep = " and dbo.[ufn_GetCoreTeamIDFromComponentName](o.Component) in (" & mid(strList,2) & ")  "
		else
			ListPrep = " and " & FieldName & " in (" & mid(strList,2) & ") "
		end if
	end function

	function ItemPrep(strText, Fieldname, ItemType)
		if Fieldname="Status" and strtext = "Not Closed" then
			ItemPrep = " and " & FieldName & " <> 'Closed' "
		elseif ItemType = 0 then 'Text Items
			ItemPrep = " and " & FieldName & " = '" & scrubsql(replace(strText,"'","''")) & "' "
		elseif ItemType = 2 and strText = "Not Specified" then
			ItemPrep = " and " & FieldName & " is null "
		else
			ItemPrep = " and " & FieldName & " = '" & scrubsql(trim(strText)) & "' "
		end if
	end function

	function DaysPrep(strTextDays, strTextRange, Fieldname, CompareType)
		if clng(CompareType) = 1 then 'Less Than
			DaysPrep = " and " & Fieldname & " < " & clng(strTextDays) & " "
		elseif clng(CompareType) = 2 then 'Equal To
			DaysPrep = " and " & Fieldname & " = " & clng(strTextDays) & " "
		elseif clng(CompareType) = 3 then 'Greater Than
			DaysPrep = " and " & Fieldname & " > " & clng(strTextDays) & " "
		elseif clng(CompareType) = 4 then 'Range
			if clng(strTextDays) > clng(strTextRange) then
				DaysPrep = " and " & Fieldname & " between " & clng(strTextRange) & " and " & clng(strTextDays) & " "
			else
				DaysPrep = " and " & Fieldname & " between " & clng(strTextDays) & " and " & clng(strTextRange) & " "
			end if
		else
			DaysPrep = ""
		end if
	end function

	function DatePrep(strTextDays, strTextRange1, strTextRange2, Fieldname, CompareType)
		dim strTemp

		if clng(CompareType) = 1 and FieldName <> "TargetDate" then 'Less Than - Not Target
			DatePrep = " and datediff(dd," & Fieldname & ",Getutcdate()) < " & clng(strTextDays) & " "
		elseif clng(CompareType) = 2 and FieldName <> "TargetDate" then 'Equal To - Not Target
			DatePrep = " and datediff(dd," & Fieldname & ",Getutcdate()) = " & clng(strTextDays) & " "
		elseif clng(CompareType) = 3 and FieldName <> "TargetDate"then 'Greater Than - Not Target
			DatePrep = " and datediff(dd," & Fieldname & ",Getutcdate()) > " & clng(strTextDays) & " "
		elseif clng(CompareType) = 1 and FieldName = "TargetDate" then 'Less Than - Target
			DatePrep = " and datediff(dd,Getutcdate()," & Fieldname & ") < " & clng(strTextDays) & " "
		elseif clng(CompareType) = 2 and FieldName = "TargetDate" then 'Equal To - Taregt
			DatePrep = " and datediff(dd,Getutcdate()," & Fieldname & ") = " & clng(strTextDays) & " "
		elseif clng(CompareType) = 3 and FieldName = "TargetDate"then 'Greater Than - Target
			DatePrep = " and datediff(dd,Getutcdate()," & Fieldname & ") > " & clng(strTextDays) & " "
		elseif clng(CompareType) = 4 then 'Range
			if strTextRange1 = "" then
				strTextRange1 = "1/1/1970"
			end if
			if strTextRange2 = "" then
				strTextRange2 = Date()
			end if
			if datediff("d",strTextRange1,strTextRange2) < 0 then
				strTemp = strTextRange1
				strTextRange1 = strTextRange2
				strTextRange2 = strTemp
			end if

			strTextRange2 = Dateadd("d",1,strTextRange2)

			DatePrep = " and " & Fieldname & " between '" & cdate(strTextRange1) & "' and '" & cdate(strTextRange2) & "' "
		else
			DatePrep = ""
		end if
	end function

	function GetValue(strFieldname)
		dim strTemp
		if ProfileData <> "" then
			if instr("&" & profileData,"&" & strFieldname & "=") > 0 then
				strTemp = mid(ProfileData,instr("&" & profileData,"&" & strFieldname & "=")+len(strFieldname) + 1)
				strTemp = left(strTemp & "&",instr(strTemp & "&","&")-1)
				getValue = urldecode(strTemp)
			else
				GetValue=""
			end if
		elseif trim(request(strFieldname)) <> "" then
			GetValue = request(strFieldname)
		else
			GetValue= ""
		end if
	end function

	function URLDecode(byVal encodedstring)
		Dim strIn, strOut, intPos, strLeft,strRight, intLoop
		strIn = encodedstring : strOut = "" : intPos = Instr(strIn, "+")
		Do While intPos
			strLeft = "" : strRight = ""
			If intPos > 1 then strLeft = Left(strIn, intPos - 1)
			If intPos < len(strIn) then strRight = Mid(strIn, intPos + 1)
			strIn = strLeft & " " & strRight
			intPos = InStr(strIn, "+")
			intLoop = intLoop + 1
		Loop
		intPos = InStr(strIn, "%")
		Do while intPos
			If intPos > 1 then strOut = strOut & Left(strIn, intPos - 1)
			strOut = strOut & Chr(CInt("&H" & mid(strIn, intPos + 1, 2)))
			If intPos > (len(strIn) - 3) then
				strIn = ""
			Else
				strIn = Mid(strIn, intPos + 3)
			End If
			intPos = InStr(strIn, "%")
		Loop
		URLDecode = strOut & strIn
	end function

	set rs = nothing
	cnExcalibur.Close
	set cnExcalibur = nothing
	cnSIO.Close
	set cnSIO = nothing

	function GetSortName (strColumn)
		Select Case trim(lcase(strColumn))
		Case "e.fullname"
			GetSortName = "Owner"
		Case "e2.fullname"
			GetSortName = "PM"
		Case "e4.fullname"
			GetSortName = "developerfullname"
		Case else
			if instr(strColumn,".")> 0 then
				GetSortName = mid(strColumn,instr(strColumn,".")+1)
			else
				GetSortName = strColumn
			end if
		end select
	end function

	function ProperCase(sString)
		Dim lTemp
		Dim sTemp, sTemp2
		Dim x
		sString = LCase(sString)
		if Len(sString) Then
			sTemp = Split(sString, " ")
			lTemp = UBound(sTemp)
			For x = 0 To lTemp
				sTemp2 = sTemp2 & UCase(Left(sTemp(x), 1)) & Mid(sTemp(x), 2) & " "
			Next
			ProperCase = trim(sTemp2)
		Else
			ProperCase = sString
		End if
	end function

	function FormatUpdates(strUpdates, MaxUpdateSize, ObservationID)
		dim strOutput
		dim RowArray
		dim ValueArray
		dim RowCount
		dim i
		dim MaxSize
		if trim(MaxUpdateSize) = "" then
			MaxSize=0
		else
			MaxSize=clng(MaxUpdateSize)
		end if

		strOutput = ""
		RowCount = 0
		RowArray = split(strUpdates,chr(4))
		for i = 0 to ubound(RowArray)
			if trim(RowArray(i)) <> "" then
'				if RowCount >= MaxUpdates and MaxUpdates <> 0 then
				if len(strOutput) > MaxSize + (175 * RowCount) and MaxSize <> 0 then
'					strOutput = strOutput & "<BR style=""mso-data-placement: same-cell""/><BR style=""mso-data-placement: same-cell""/><a target=""_blank"" href=""report_mattH.asp?txtReportSections=2&txtObservationID=" & ObservationID & """>View All Updates</a>"
					strOutput = strOutput & "<BR style=""mso-data-placement: same-cell""/><BR style=""mso-data-placement: same-cell""/><a target=""_blank"" href=""report.asp?txtReportSections=2&txtObservationID=" & ObservationID & """>View All Updates</a>"
					exit for
				end if
				ValueArray = split(RowArray(i),chr(3))

				if i > 0 then
					strOutput = strOutput & "<BR style=""mso-data-placement: same-cell""/><BR style=""mso-data-placement: same-cell""/>"
				end if
				if ubound(ValueArray) < 2 and strPageParams <> "macro" then
					strOutput = strOutput & ReplaceAndHTMLEncodeFrom(rowarray(i))
				else
					if ubound(ValueArray) = 2 and strPageParams <> "macro" then
						strOutput = strOutput & "<b>" & ReplaceAndHTMLEncodeFrom(ValueArray(0)) & " by " & ReplaceAndHTMLEncodeFrom(ValueArray(1)) & _
							"</b><BR style=""mso-data-placement: same-cell""/>" & ReplaceAndHTMLEncodeFrom(ValueArray(2))
					else
						for j = 0 to ubound(ValueArray)
							if j = 0 then
								strOutput = strOutput & "<b>" & formatdatetime(ValueArray(j))
							else
								if j = ubound(ValueArray) then
									strOutput = strOutput & "</b><BR style=""mso-data-placement: same-cell""/>"
								else
									strOutput = strOutput & " - "
								end if
								strOutput = strOutput & ReplaceAndHTMLEncodeFrom(ValueArray(j))
							end if
						next
					end if
				end if
				RowCount=RowCount+1
			end if
		next
		FormatUpdates = strOutput
	end function


	function FormatUpdatesXML(strUpdates, MaxUpdateSize, ObservationID)
		dim strOutput
		dim RowArray
		dim ValueArray
		dim RowCount
		dim i
		dim MaxSize
		if trim(MaxUpdateSize) = "" then
			MaxSize=0
		else
			MaxSize=clng(MaxUpdateSize)
		end if

		strOutput = ""
		RowCount = 0
		RowArray = split(strUpdates,chr(4))
		for i = 0 to ubound(RowArray)
			if trim(RowArray(i)) <> "" then
				if len(strOutput) > MaxSize + (175 * RowCount) and MaxSize <> 0 then
					strOutput = strOutput & "[Truncated]"
					exit for
				end if
				ValueArray = split(RowArray(i),chr(3))
				if i > 0 then
					strOutput = strOutput & vbcrlf & vbcrlf
				end if
'				strOutput = strOutput & ValueArray(0) & " by " & ValueArray(1) & vbcrlf & ValueArray(2)
				for j = 0 to ubound(ValueArray)
					select case(j)
						case 0
							strOutput = strOutput & Server.HTMLEncode(ValueArray(j))
						case 1
							strOutput = strOutput & " by " & Server.HTMLEncode(ValueArray(j))
						case 2
							strOutput = strOutput & vbcrlf & Server.HTMLEncode(ValueArray(j))
					end select
				next
				RowCount=RowCount+1
			end if
		next
		FormatUpdatesXML = strOutput
	end function

	function myFormatDateTime(strDateTime)
		dim strYear, strMonth, strDay, strTime
		myFormatDateTime = ""
		if (strDateTime <> "" ) then
			strYear = DatePart("yyyy", strDateTime)
			strMonth = DatePart("m", strDateTime)
			strDay = DatePart("d", strDateTime)
			strTime = FormatDateTime(strDateTime, vbShortTime)
			myFormatDateTime = Server.HTMLEncode(strYear & "-" & strMonth & "-" & strDay & " " & strTime)
		end if
	end function

	function FormatDataCell(strFieldName, strData,strFormat, strCellParam)
		select case clng(strFormat)
		case 2
'			FormatDataCell = strFieldname & "=""" & replace(replace(replace(replace(replace(replace(replace(strData,"<","&lt;"),">","&gt;"),"&","&amp;"),"'","&apos;"),"""","&quot;"), "", "&quot;"), "", "&quot;") & """ "
			FormatDataCell = strFieldname & "=""" & Server.HTMLEncode(strData) & """ "
		case 1
			if trim(strData) = "" then
				FormatDataCell = "<" & strfieldname & " />" & vblf
			else
'				FormatDataCell = "<" & strfieldname & ">" & replace(replace(replace(replace(replace(replace(replace(strData,"<","&lt;"),">","&gt;"),"&","&amp;"),"'","&apos;"),"""","&quot;"), "", "&quot;"), "", "&quot;") & "</" & strfieldname & ">" & vblf
				FormatDataCell = "<" & strfieldname & ">" & Server.HTMLEncode(strData) & "</" & strfieldname & ">" & vblf
			end if
		case else
			if trim(strData) = "" then
				FormatDataCell = "<td " & strCellParam & ">&nbsp;</td>" & vblf
			else
				FormatDataCell = "<td " & strCellParam & ">" & replace(strData,vbcr,"<br style=""mso-data-placement:same-cell;""/>") & "</td>" & vblf
			end if
		end select
	end function

	function ShortenField(strText,MaxLength)
		if len(strText) > MaxLength then
			ShortenField = ReplaceAndHTMLEncodeFrom(left(strText,MaxLength)) & "...&nbsp;&nbsp;<b>[truncated]</b>"
		else
			ShortenField = ReplaceAndHTMLEncodeFrom(strText)
		end if
	end function

	function ShortenFieldXML(strText,MaxLength)
		if len(strText) > MaxLength then
			ShortenFieldXML = Server.HTMLEncode(left(strText,MaxLength)) & "... [truncated]"
		else
			ShortenFieldXML = Server.HTMLEncode(strText)
		end if
	end function

	function ScrubSQL(strWords)
		dim badChars
		dim newChars
		dim i

'		strWords=replace(strWords,"'","''")

		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update")
		newChars = strWords

		for i = 0 to uBound(badChars)
			newChars = replace(newChars, badChars(i), "")
		next

		ScrubSQL = newChars
	end function

	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(") then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
	end function

	function FilterContains(strFilter,strField, strValue)
		dim re

		Set re = new regexp

		re.Pattern = "[\s\(.]" & strField & "(\s*)=(\s*)'" & strValue & "'|[\s\(.]" & strField & "(\s+)in(\s*)(\()(\s*)'" & strValue & "'|[\s\(.]" & strField & "(\s+)like(\s+)'%?" & strValue & ""
		re.IgnoreCase = true

		FilterContains = re.Test(strFilter)
	end function

	function IsArrayEmpty(myArray)
		dim element, result
		result = true
		for each element in myArray
			result = false
			exit for
		next
		IsArrayEmpty = result
	end function

	if XMLFormat = 0 then
	%>
	<div id="mnuPopup" style="display: none; position: absolute; width: 2px; height: 2px; left: 0px; top: 0px; padding: 0px; background: white; border: 1px solid gainsboro; z-index: 100">
	</div>
	<%if trim(request("txtReportSections")) = "3" then%>
	<iframe style="display: none" id="frameSendEmail"></iframe>
	<%end if%>
</body>
</html>
	<%elseif XMLFormat = 1 then
		response.write "</observations>" & vblf
	else
		response.write "</root>" & vblf
	end if%>
