<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Custom Report</title>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
		var blnDragging;
		var SourceCell;
		var SourceTable;
		var SourceIndex;
		var DestinationTable;
		var DestinationCell;
		var MSIE = navigator.userAgent.indexOf('MSIE') >= 0 ? true : false;
		var navigatorVersion = navigator.appVersion.replace(/.*?MSIE (\d\.\d).*/g, '$1') / 1;

		function window_onload() {
			var VerticalNodes = AvailableRows.getElementsByTagName("TD");
			for (i = 0; i < VerticalNodes.length; i++) {
				VerticalNodes[i].onmouseup = EnableDropZoneRows;
				VerticalNodes[i].onmousemove = EnableDropZoneRows;
				VerticalNodes[i].onmousedown = NodeMouseDown;
				if (VerticalNodes[i].className != "Pusher")
					VerticalNodes[i].style.cursor = 'hand';
			}

			VerticalNodes = SelectedRows.getElementsByTagName("TD");
			for (i = 0; i < VerticalNodes.length; i++) {
				VerticalNodes[i].onmouseup = EnableDropZoneRows;
				VerticalNodes[i].onmousemove = EnableDropZoneRows;
				VerticalNodes[i].onmousedown = NodeMouseDown;
				if (VerticalNodes[i].className != "Pusher")
					VerticalNodes[i].style.cursor = 'hand';
			}

			document.documentElement.onmousemove = moveDragMe;
			document.documentElement.onmouseup = dropDragMe;
			SourceTable = AvailableRowTable;
			document.onselectstart = function () { return false; };
			blnDragging = false;
			DestinationTable = AvailableRows;
		}

		function dropDragMe() {
			var NewCell;
			var ConfigureSpan;
			var ParamSpan;

			DestinationTable.style.backgroundColor = "white";
			if (blnDragging && typeof (Target) == "unknown") {
				DestinationRow = SourceTable.insertRow(SourceIndex);
				DestinationCell = DestinationRow.insertCell(0);
				if (SourceTable.id == "AvailableRows")
					DestinationCell.innerHTML = "<div id=Target class=Dragging FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
				else
					DestinationCell.innerHTML = "<div id=Target class=Dragging style=\"height:" + DragMe.style.height + ";width:" + DragMe.style.width + ";\" FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
			}
			if (blnDragging && typeof (Target) != "undefined") {
				Target.innerHTML = DragMe.innerHTML;
				Target.className = "";
				Target.id = SourceCell.id;
				DragMe.style.display = "none";
				DragMe.innerHTML = "";
				blnDragging = false;
				NewCell = DestinationCell;
				NewCell.onmousedown = NodeMouseDown;
				NewCell.onmouseenter = EnableDropZoneRows;
				NewCell.onmousemove = EnableDropZoneRows;
				NewCell.onmouseup = EnableDropZoneRows;

				NewCell.style.cursor = 'hand';
				if (DestinationTable.id == "AvailableRows") {
					ConfigureSpan = document.getElementById("Configure" + SourceCell.id.substring(4));
					ConfigureSpan.style.display = "none";
					ParamSpan = document.getElementById("Params" + SourceCell.id.substring(4));
					ParamSpan.innerText = "";
					document.all("Node" + SourceCell.id.substring(4)).style.color = "black";
				}
				else if (SourceCell.id.substring(4) != "4" && SourceCell.id.substring(4) != "14" && SourceCell.id.substring(4) != "15" && SourceCell.id.substring(4) != "16" && SourceCell.id.substring(4) != "17" && SourceCell.id.substring(4) != "18" && SourceCell.id.substring(4) != "19" && SourceCell.id.substring(4) != "20" && SourceCell.id.substring(4) != "21" && SourceCell.id.substring(4) != "22" && SourceCell.id.substring(4) != "23" && SourceCell.id.substring(4) != "24" && SourceCell.id.substring(4) != "25" && SourceCell.id.substring(4) != "26" && SourceCell.id.substring(4) != "27") {
					ConfigureSpan = document.getElementById("Configure" + SourceCell.id.substring(4))
					ConfigureSpan.style.display = "none";
				}
				else {
					ConfigureSpan = document.getElementById("Configure" + SourceCell.id.substring(4))
					ConfigureSpan.style.display = "";
					ParamSpan = document.getElementById("Params" + SourceCell.id.substring(4));
					if (ParamSpan.innerText == "")
						document.all("Node" + SourceCell.id.substring(4)).style.color = "black";
					else
						document.all("Node" + SourceCell.id.substring(4)).style.color = "green";
				}
				DestinationCell = undefined;
			}
		}

		function moveDragMe() {
			if (blnDragging) {
				var st = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
				var sl = Math.max(document.body.scrollLeft, document.documentElement.scrollLeft);
				DragMe.style.left = (event.clientX + sl) + 'px';
				DragMe.style.top = (event.clientY + st) + 'px';
			}
		}

		function NodeMouseDown() {
			if (event.srcElement.id.substring(0, 4) == "Node") {
				SourceCell = event.srcElement;
				SourceTable = event.srcElement;

				while (SourceTable.tagName != "TABLE" && SourceTable.tagName != "") {
					SourceTable = SourceTable.parentNode;
				}

				blnDragging = true;
				DragMe.innerHTML = event.srcElement.innerHTML;
				event.srcElement.innerHTML = "";
				event.srcElement.className = "Dragging";
				var st = Math.max(document.body.scrollTop, document.documentElement.scrollTop);
				var sl = Math.max(document.body.scrollLeft, document.documentElement.scrollLeft);
				DragMe.style.left = (event.clientX + sl) + 'px';
				DragMe.style.top = (event.clientY + st) + 'px';
				DragMe.style.width = GetFieldWidth(event.srcElement.id); //event.srcElement.FieldWidth;
				DragMe.style.height = GetFieldHeight(event.srcElement.id); //event.srcElement.FieldHeight;
				DragMe.style.display = "";
				SourceIndex = SourceCell.parentNode.parentNode.rowIndex;
				SourceTable.deleteRow(SourceCell.parentNode.parentNode.rowIndex);
			}
		}

		function GetFieldWidth(ElementID) {
			return 600;
		}

		function GetFieldHeight(ElementID) {
			return 50;
		}

		function EnableDropZoneRows() {
			var MyTable;
			var MyRow;
			var MyCell;
			var tmpIndex = -1;
			var DestinationRow;
			if (blnDragging) {
				MyTable = event.srcElement
				while (MyTable.tagName != "TABLE" && MyTable.tagName != "") {
					if (MyTable.tagName == "TD")
						MyCell = MyTable;
					else if (MyTable.tagName == "TR")
						MyRow = MyTable;
					MyTable = MyTable.parentNode;
				}

				MyTable.style.backgroundColor = "Lavender";
				if (DestinationTable.id != MyTable.id) {
					DestinationTable.style.backgroundColor = "white";
					if (typeof (DestinationCell) != "undefined")
						DestinationTable.deleteRow(DestinationCell.parentNode.rowIndex);
				}
				else if (MyCell != DestinationCell) {
					if (typeof (DestinationCell) != "undefined") {
						tmpIndex = DestinationCell.parentNode.rowIndex;
						DestinationCell.parentNode.parentNode.deleteRow(DestinationCell.parentNode.rowIndex);
					}
				}
				DestinationTable = MyTable;
				if (tmpIndex != -1 && tmpIndex <= MyRow.rowIndex && MyCell.className != "Pusher") {
					DestinationRow = MyTable.insertRow(MyRow.rowIndex + 1);
					DestinationCell = DestinationRow.insertCell(MyCell.cellIndex);
				}
				else {
					DestinationRow = MyTable.insertRow(MyRow.rowIndex);
					DestinationCell = DestinationRow.insertCell(MyCell.cellIndex);
				}
				if (MyTable.id == "AvailableRows")
					DestinationCell.innerHTML = "<div id=Target class=Dragging FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
				else
					DestinationCell.innerHTML = "<div id=Target class=Dragging style=\"height:" + DragMe.style.height + ";width:" + DragMe.style.width + ";\" FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
			}
		}

		function SaveLayout() {
			//get Selected Sections
			var LayoutArray = SelectedRows.getElementsByTagName("DIV");
			var strResults = "";
			var OutArray = new Array();

			for (i = 0; i < LayoutArray.length; i++) {
				if (LayoutArray[i].id.substr(0, 4) == "Node")
					if (strResults == "")
						strResults = LayoutArray[i].id.substring(4);
					else
						strResults = strResults + "," + LayoutArray[i].id.substring(4);
			}
			OutArray[0] = strResults;
			strResults = "";

			//get Parameters for Selected Sections
			var LayoutArray2 = SelectedRows.getElementsByTagName("SPAN");

			for (i = 0; i < LayoutArray2.length; i++) {
				if (LayoutArray2[i].id.substr(0, 6) == "Params")
					strResults = strResults + "|" + LayoutArray2[i].innerText;
			}

			if (strResults == "")
				OutArray[1] = "";
			else
				OutArray[1] = strResults.substring(1);
			return OutArray;
		}

		function cmdFinish_onclick() {
			var OutArray = new Array();

			OutArray = SaveLayout();

			if (navigator.appName != "Microsoft Internet Explorer" && navigator.appName != "Internet Explorer" && navigator.appName != "IE")
				if (typeof (window.parent.opener) != "undefined") {
					window.parent.opener.frmMain.action = "Report.asp"
					window.parent.opener.frmMain.target = "_blank"
					window.parent.opener.frmMain.txtReportSections.value = OutArray[0];
					window.parent.opener.frmMain.txtReportSectionParameters.value = OutArray[1];
					window.parent.opener.frmMain.submit();
				}

			window.returnValue = OutArray;
			top.close();
		}

		function cmdCancel_onclick() {
			top.close();
		}

		function AddProfile() {
			var strLayout = SaveLayout();
			if (strLayout[0] == "") {
				alert("Select at least one report section before saving your report.");
				return;
			}
			var strID = new Array();
			txtReturnValue.value = "";
			strID = window.showModalDialog("ProfileProperties.asp?ReportType=1", "", "dialogWidth:655px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
			if (typeof (strID) != "undefined" || txtReturnValue.value != "") {
				frmMain.txtProfileUpdateType.value = "1";
				frmMain.txtProfileType.value = "7";
				if (txtReturnValue.value != "")
					frmMain.txtNewProfileName.value = txtReturnValue.value;
				else
					frmMain.txtNewProfileName.value = strID[0];
				frmMain.txtNewTodayLink.value = "0";
				frmMain.txtNewReportFormat.value = "0";
				frmMain.txtProfileUpdateID.value = "0";
				frmMain.txtPageLayout.value = strLayout[0];
				frmMain.txtFieldFilters.value = strLayout[1];
				frmMain.target = "ProfileFrame";
				frmMain.action = "UpdateProfile.asp"
				frmMain.submit();
			}
		}

		function UpdateProfile() {
			var strLayout = SaveLayout();
			if (strLayout[0] == "") {
				alert("Select at least one report section before saving your report.");
				return;
			}
			if (confirm("Are you sure you want to save changes to this report?")) {
				frmMain.txtProfileUpdateType.value = "2";
				frmMain.txtNewProfileName.value = cboProfile.options[cboProfile.selectedIndex].text;
				frmMain.txtNewTodayLink.value = 0;
				frmMain.txtNewReportFormat.value = 0;
				frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
				frmMain.txtPageLayout.value = strLayout[0];
				frmMain.txtFieldFilters.value = strLayout[1];
				frmMain.target = "ProfileFrame";
				frmMain.action = "UpdateProfile.asp"
				frmMain.submit();
			}
		}

		function RenameProfile() {
			var strNewName;
			strNewName = window.prompt("Enter new name for this report.", cboProfile.options[cboProfile.selectedIndex].text);

			if (strNewName != null) {
				frmMain.txtNewProfileName.value = strNewName;
				frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
				frmMain.txtProfileUpdateType.value = "3";
				frmMain.target = "ProfileFrame";
				frmMain.action = "UpdateProfile.asp"
				frmMain.submit();
			}
		}

		function ShareProfile() {
			var strResult;
			strResult = window.showModalDialog("ProfileShare.asp?ID=" + cboProfile.options[cboProfile.selectedIndex].value, "", "dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
		}

		function RemoveProfile() {
			if (window.confirm("Are you sure you want to stop receiving this shared report?")) {
				frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].SharingID;
				frmMain.txtNewProfileName.value = "";
				frmMain.txtProfileUpdateType.value = "5";
				frmMain.target = "ProfileFrame";
				frmMain.action = "UpdateProfile.asp"
				frmMain.submit();
			}
		}

		function DeleteProfile() {
			if (window.confirm("Are you sure you want to delete this profile?")) {
				frmMain.txtProfileUpdateID.value = cboProfile.options[cboProfile.selectedIndex].value;
				frmMain.txtNewProfileName.value = "";
				frmMain.txtProfileUpdateType.value = "4";
				frmMain.target = "ProfileFrame";
				frmMain.action = "UpdateProfile.asp"
				frmMain.submit();
			}
		}

		function ProfileSaved(strType, strID, strResult, strError) {
			if (strType == "3") {
				if (strResult == "0")
					alert("Error Renaming Profile: " + strError);
				else {
					cboProfile.options[cboProfile.selectedIndex].text = frmMain.txtNewProfileName.value;
				}
			}
			else if (strType == "4") {
				if (strResult == "0")
					alert("Error Deleting Profile: " + strError);
				else {
					cboProfile.options[cboProfile.selectedIndex] = null;
					cboProfile.selectedIndex = 0;
					//alert("Profile Deleted.");
					cboProfile_onchange();
				}
			}
			else if (strType == "5") {
				if (strResult == "0")
					alert("Error Removing Profile: " + strError);
				else {
					cboProfile.options[cboProfile.selectedIndex] = null;
					cboProfile.selectedIndex = 0;
					//alert("Profile Removed.");
					cboProfile_onchange();
				}
			}
			else if (strType == "1") {
				if (strResult == "0")
					alert("Error Adding Profile: " + strError);
				else {
					cboProfile.options[cboProfile.length] = new Option(frmMain.txtNewProfileName.value, strResult);
					window.location.href = "CustomReport.asp?ProfileID=" + strResult + "&RunReportOK=" + txtRunReportOK.value;

				}
			}
			else {
				if (strResult == "0")
					alert("Error Updating Profile: " + strError);
				else {
					cboProfile.options[cboProfile.selectedIndex].text = frmMain.txtNewProfileName.value
					alert("Profile Updated.");
				}
			}
			frmMain.txtNewProfileName.value = "";
			frmMain.txtProfileUpdateID.value = "";
			frmMain.txtProfileUpdateType.value = "";
			frmMain.txtNewTodayLink.value = "";
			frmMain.txtNewReportFormat.value = "";
		}

		function cboProfile_onchange() {
			var strColumns;
			var strProducts;
			var strBuffer;
			var i;
			var strHeader;

			ProfileOptionsAdd.style.display = "none";
			ProfileOptionsUpdate.style.display = "none";
			ProfileOptionsDelete.style.display = "none";
			ProfileOptionsRename.style.display = "none";
			ProfileOptionsOwner.style.display = "none";
			ProfileOptionsRemove.style.display = "none";
			ProfileOptionsShare.style.display = "none";


			FilterLoadingMessage.style.display = "";
			FilterLoadingMessage.style.width = FilterLoadingMessage.scrollWidth + 10;
			FilterLoadingMessage.style.height = FilterLoadingMessage.scrollHeight;
			FilterLoadingMessage.style.left = 200;
			FilterLoadingMessage.style.top = 76;
			if (cboProfile.selectedIndex > 0) {
				window.location.href = "customreport.asp?ProfileID=" + cboProfile.options[cboProfile.selectedIndex].value + "&RunReportOK=" + txtRunReportOK.value;
			}
			else {
				window.location.href = "customreport.asp?RunReportOK=" + txtRunReportOK.value;
			}
		}

		function ConfigureSection(ID) {
			var strID = "";
			var ConfigureType = "";
			if (ID == 4 || ID == 14 || ID == 21)
				ConfigureType = 1;
			else if (ID == 22 || ID == 23 || ID == 24 || ID == 25 || ID == 26 || ID == 27)
			    ConfigureType = 3;
			else
				ConfigureType = 2;

			strID = window.showModalDialog("ConfigureReportSections.asp?txtID=" + ID + "&TypeID=" + ConfigureType + "&txtParams=" + document.all("Params" + ID).innerText.replace("&", "and"), "", "dialogWidth:530px;dialogHeight:450px;edge: Sunken;center:Yes; help: No;maximize:No;resizable: Yes;status: No")
			if (typeof (strID) != "undefined") {
				document.all("Params" + ID).innerText = strID;
				if (strID == "")
					document.all("Node" + ID).style.color = "black";
				else
					document.all("Node" + ID).style.color = "green";
			}
		}
//-->
	</script>
	<style type="text/css">
		td
		{
			font-family: Verdana;
			font-size: xx-small;
		}
		body
		{
			font-family: Verdana;
			font-size: xx-small;
			background-color: #e6e6fa; /* lavender */
		}
		
		#MainDragArea, #Step2Panel
		{
			/* Main container for this script */
			width: 100%;
			height: 120px;
			border: 1px solid #317082;
			background-color: #FFF;
			-moz-user-select: none;
		}
		#AvailableBox div
		{
			margin: 2px;
			border: 1px solid black;
			background-color: #EEE;
			width: 100%;
			padding: 2px;
		}
		
		#AvailableBox div.Dragging
		{
			margin: 2px;
			border: 1px solid black;
			background-color: white;
			width: 100%;
			padding: 2px;
		}
		
		#DropZone div
		{
			margin: 2px;
			border: 1px solid black;
			background-color: #EEE;
			padding: 2px;
		}
		
		#DropZone div.Dragging
		{
			margin: 2px;
			border: 1px solid black;
			background-color: white;
			width: 100%;
			padding: 2px;
		}
		
		#DragMe
		{
			margin: 2px;
			border: 1px solid black;
			background-color: #EEE;
			width: 150px;
			padding: 2px;
			position: absolute;
		}
		#dragContent1
		{
			/* Drag container */
			position: absolute;
			width: 150px;
			height: 20px;
			display: none;
			margin: 0px;
			padding: 0px;
			z-index: 2000;
		}
		A:visited
		{
			color: blue;
		}
		A:link
		{
			color: blue;
		}
		A:hover
		{
			color: red;
		}
		TD.HeaderButton
		{
			font-size: xx-small;
			font-family: Verdana;
			font-weight: bold;
			color: White;
		}
	</style>
</head>
<body onload="window_onload()">
	<%
	dim cn, rs, strSQL, cm, p, j
	dim strProfileOptions
	dim strProfile
	dim FieldPropertyArray
	dim FieldAttributes
	dim FieldAttributeArray
	dim FieldArray
	dim ValueArray
	dim strField
	dim strMobileConsumerChecked
	dim strMobileCommercialChecked
	dim strMobileFunctionalChecked
	dim strDTOChecked
	dim FieldFilterValues
	dim UserSettingArray
	dim strAvailableColumnIDs
	dim strDeveloperIDList
	dim strComponentPMIDList
	dim strSubSystemIDList
	dim strCoreTeamIDList
	dim strTypeIDList
	dim strSelectedDeveloperList
	dim strSelectedSubSystemList
	dim strSelectedCoreTeamList
	dim strSelectedTypeList
	dim strSelectedComponentPMList
	dim CurrentDomain, CurrentUser, CurrentUserID, CurrentUserDivision, CurrentUserPartner
	dim strRow1Fields,strRow2Fields,strRow3Fields,strRow4Fields,strRow5Fields,strRow6Fields, strAvailableFields
	dim blnProfileFound
	dim blnProfileCanEdit
	dim blnProfileCanDelete
	dim blnProfileCanRemove
	dim strProfilePrimaryOwner
	dim strProfilePageLayout
	dim strProfileFilters
	dim SelectedRowFields
	dim FilterArray

	strAvailableColumnIDs = ",5,9,6,28,8,11,12,10,13,0,4,14,15,16,17,18,19,20,21,22,23,24,25,26,27,"

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	cn.CommandTimeout = 180

	'Get User
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

	set cm=nothing

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserDivision = rs("Division") & ""
		CurrentUserPartner = rs("PartnerID") & ""
	else
		Response.Redirect "../Excalibur.asp"
	end if
	rs.Close

	blnProfileFound = false
	blnProfileCanEdit = false
	blnProfileCanDelete = false
	blnProfileCanRemove = true
	strProfilePrimaryOwner = ""
	strProfilePageLayout = ""
	strProfileFilters = ""

	if trim(request("ProfileID")) <> "" then
		rs.open "spGetReportProfile " & clng(request("ProfileID")),cn,adOpenStatic
		if not(rs.eof and rs.bof) then
			strProfile=trim(rs("ID"))
			strProfilePageLayout=rs("PageLayout")
			strProfileFilters = rs("SelectedFilters") & ""
		end if
		rs.Close
	end if

	rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	strProfileOptions = ""
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<Option selected SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = true
			blnProfileCanDelete = true
		else
			strProfileOptions = strProfileOptions & "<Option SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<Option selected SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) & """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = cbool(rs("CanEdit"))
			blnProfileCanDelete= cbool(rs("CanDelete"))
			strProfilePrimaryOwner = shortname(rs("PrimaryOwner"))
		else
			strProfileOptions = strProfileOptions & "<Option SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) & """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	do while not rs.EOF
		if strProfile = trim(rs("ID")) then
			strProfileOptions = strProfileOptions & "<Option selected CanRemove=0 SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) & """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			blnProfileFound = true
			blnProfileCanEdit = cbool(rs("CanEdit"))
			blnProfileCanDelete= cbool(rs("CanDelete"))
			strProfilePrimaryOwner = shortname(rs("PrimaryOwner"))
			blnProfileCanRemove = false
		else
			strProfileOptions = strProfileOptions & "<Option CanRemove=0 SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")) & """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

	if strProfileOptions <> "" then
		strProfileOptions = "<option selected></option>" & strProfileOptions
	end if




	SelectedRowFields = ""
	FieldArray = split(strProfilePageLayout,",")
	FilterArray = split(strProfileFilters,"|")
	for i = 0 to ubound(FieldArray)
		if trim(Fieldarray(i)) <> "" then
			FieldAttributes = LookupFieldAttributes(Fieldarray(i))
			if FieldAttributes <> "" then
				FieldAttributeArray = split(FieldAttributes,"|")
				if instr("," & replace(strProfilePageLayout," ","") & ",","," & replace(Fieldarray(i)," ","") & ",") > 0 then
					if strProfileFilters = "" then
						strFilters = ""
					else
						if i > ubound(FilterArray) then
							strFilters = ""
						else
							strFilters = trim(Filterarray(i))
						end if
					end if
					if strFilters="" then
						DivFontColor= "black"
					else
						DivFontColor= "green"
					end if

					if trim(Fieldarray(i)) = "4" or trim(Fieldarray(i)) = "14" or trim(Fieldarray(i)) = "15" or trim(Fieldarray(i)) = "16" or trim(Fieldarray(i)) = "17" or trim(Fieldarray(i)) = "18" or trim(Fieldarray(i)) = "19" or trim(Fieldarray(i)) = "20" or trim(Fieldarray(i)) = "21" or trim(Fieldarray(i)) = "22" or trim(Fieldarray(i)) = "23" or trim(Fieldarray(i)) = "24" or trim(Fieldarray(i)) = "25" or trim(Fieldarray(i)) = "26" or trim(Fieldarray(i)) = "27" then
						SelectedRowFields = SelectedRowFields & "<tr><td><div style=""color:" & DivFontColor & ";height:50;"" FieldHeight=50 FieldWidth=600 id=Node" & trim(Fieldarray(i)) & ">" & FieldAttributeArray(0) & "<span style=""display:none"" id=""Params" & trim(Fieldarray(i)) & """>" & strFilters & "</span><span id=""Configure" & trim(Fieldarray(i)) & """>&nbsp;-&nbsp;<a href=""javascript: ConfigureSection(" & trim(Fieldarray(i)) & ");"">Configure</a></span></div></td></tr>"
					else
						SelectedRowFields = SelectedRowFields & "<tr><td><div style=""color:" & DivFontColor & ";height:50;"" FieldHeight=50 FieldWidth=600 id=Node" & trim(Fieldarray(i)) & ">" & FieldAttributeArray(0) & "<span style=""display:none"" id=""Params" & trim(Fieldarray(i)) & """>" & strFilter & "</span><span style=""display:none"" id=""Configure" & trim(Fieldarray(i)) & """>&nbsp;-&nbsp;<a href=""javascript: ConfigureSection(" & trim(Fieldarray(i)) & ");"">Configure</a></span></div></td></tr>"
					end if
				end if
			end if
		end if
	next


	AvailableRowFields = ""
	FieldArray = split(strAvailableColumnIDs,",")
	for i = 0 to ubound(FieldArray)
		if trim(Fieldarray(i)) <> "" then
			FieldAttributes = LookupFieldAttributes(Fieldarray(i))
			if FieldAttributes <> "" then
				FieldAttributeArray = split(FieldAttributes,"|")
				if instr("," & replace(strProfilePageLayout," ","") & ",","," & replace(Fieldarray(i)," ","") & ",") = 0 then
					AvailableRowFields = AvailableRowFields & "<tr><td><div FieldHeight=50 FieldWidth=600 id=Node" & trim(Fieldarray(i)) & ">" & FieldAttributeArray(0) & "<span style=""display:none"" id=""Params" & trim(Fieldarray(i)) & """></span><span style=""display:none"" id=""Configure" & trim(Fieldarray(i)) & """>&nbsp;-&nbsp;<a href=""javascript: ConfigureSection(" & trim(Fieldarray(i)) & ");"">Configure</a></span></div></td></tr>"
				end if
			end if
		end if
	next


	dim ProfileDisplayUpdateLink, ProfileDisplayDeleteLink , ProfileDisplayRenameLink,ProfileDisplayRemoveLink, ProfileDisplayOwnerLink, ProfileDisplayShareLink
	if strProfile = "" then
		ProfileDisplayUpdateLink = "none"
		ProfileDisplayDeleteLink = "none"
		ProfileDisplayRenameLink = "none"
		ProfileDisplayRemoveLink = "none"
		ProfileDisplayOwnerLink = "none"
		ProfileDisplayShareLink = "none"
	else
			if blnProfileCanEdit then
				ProfileDisplayUpdateLink = ""
				ProfileDisplayRenameLink = ""
			else
				ProfileDisplayUpdateLink = "none"
				ProfileDisplayRenameLink = "none"
			end if

			if blnProfileCanDelete then
				ProfileDisplayDeleteLink = ""
			else
				ProfileDisplayDeleteLink = "none"
			end if

			if strProfilePrimaryOwner = "" then
				ProfileDisplayRemoveLink = "none"
				ProfileDisplayOwnerLink = "none"
				ProfileDisplayShareLink = ""
			else
				if blnProfileCanRemove then
					ProfileDisplayRemoveLink = ""
				else
					ProfileDisplayRemoveLink = "none"
				end if
				ProfileDisplayOwnerLink = ""
				ProfileDisplayShareLink = "none"

			end if
	end if

	%>
	<span style="font-size: small; font-family: verdana; font-weight: bold">Custom Observation Reports</span><br />
	<br />
	<table border="0" width="100%">
		<tr>
			<td style="white-space: nowrap">
				<b>Saved&nbsp;Reports:&nbsp;</b>
				<select id="cboProfile" name="cboProfile" style="width: 400px" onchange="return cboProfile_onchange()">
					<%=strProfileOptions%>
				</select>
				<span id="ProfileOptionsAdd" style="font-size: xx-small; font-family: verdana"><a href="javascript:AddProfile();">Add</a> </span><span style="display: <%=ProfileDisplayUpdateLink%>; font-size: xx-small; font-family: verdana" id="ProfileOptionsUpdate"><a href="javascript:UpdateProfile();">Update</a> </span><span style="display: <%=ProfileDisplayDeleteLink%>; font-size: xx-small; font-family: verdana" id="ProfileOptionsDelete"><a href="javascript:DeleteProfile();">Delete</a> </span><span style="display: <%=ProfileDisplayRenameLink%>; font-size: xx-small; font-family: verdana" id="ProfileOptionsRename"><a href="javascript:RenameProfile();">Rename</a> </span><span style="display: <%=ProfileDisplayRemoveLink%>; font-size: xx-small; font-family: verdana" id="ProfileOptionsRemove"><a href="javascript:RemoveProfile();">Remove</a> </span><span style="display: <%=ProfileDisplayShareLink%>; font-size: xx-small; font-family: verdana" id="ProfileOptionsShare"><a href="javascript:ShareProfile();">Share</a> </span><span style="display: <%=ProfileDisplayOwnerLink%>; font-size: xx-small; font-family: verdana; font-weight: bold; color: black" id="ProfileOptionsOwner">Report Owner:
					<%=strProfilePrimaryOwner%>
				</span>
			</td>
		</tr>
		<tr>
			<td colspan="8">
				<hr />
			</td>
		</tr>
	</table>
	<div id="step1">
		<div id="MainDragArea" style="height: 360px">
			<table cellpadding="0" cellspacing="0">
				<tr>
					<td id="AvailableBox" valign="top" style="height: 100%">
						<div style="border: none; height: 354px; overflow-y: scroll">
							<table id="AvailableRowTable" style="height: 100%; background-color: white; border-color: #dcdcdc" <%'gainsboro%> border="1" cellpadding="0" cellspacing="0" width="250px">
								<tr style="height: 15px">
									<td style="white-space: nowrap; background-color: #317082; color: White; font-weight: bold">
										Available Report Sections
									</td>
								</tr>
								<tr>
									<td valign="top" style="white-space: nowrap">
										<table id="AvailableRows" width="100%" style="height: 100%" cellpadding="0" cellspacing="0" border="0">
											<%=AvailableRowFields%>
											<tr style="height: 100%">
												<td class="Pusher">
													&nbsp;
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
					</td>
					<td id="DropZone" valign="top" style="width: 100%">
						<div style="border: none; height: 354px; overflow-y: scroll">
							<table id="SelectedRowTable" style="height: 100%; background-color: white; border-color: #dcdcdc" <%'gainsboro%> border="1" cellpadding="0" cellspacing="0" width="600px">
								<tr style="height: 15px">
									<td style="background-color: #317082; white-space: nowrap; color: White; font-weight: bold">
										Custom Report Layout
									</td>
								</tr>
								<tr>
									<td valign="top" style="white-space: nowrap">
										<table id="SelectedRows" style="width: 100%; height: 100%" cellpadding="0" cellspacing="0" border="0">
											<%=SelectedRowFields%>
											<tr style="height: 100%">
												<td class="Pusher">
													&nbsp;
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
					</td>
				</tr>
			</table>
		</div>
	</div>
	<hr />
	<table border="0" cellspacing="1" cellpadding="1" style="float: right">
		<tr>
			<%
			dim strDisplayType
			strDisplayType="inline"
			if trim(request("RunReportOK")) = "0" then
				strDisplayType="none"
			end if
			%>
			<td style="display: <%=strDisplayType%>">
				<input type="button" value="Run Report" id="cmdRun" name="cmdRun" onclick="return cmdFinish_onclick();" />
			</td>
			<td>
				<input type="button" value="Close" id="cmdClose" name="cmdClose" onclick="return cmdCancel_onclick();" />
			</td>
		</tr>
	</table>
	<div id="DragMe" style="display: none">
	</div>
	<%
	set rs = nothing
	cn.Close
	set cn=nothing


	function LookupFieldAttributes(ID)
		select case clng(ID)
		case 5
			LookupFieldAttributes = "Observations by Priority"
		case 9
			LookupFieldAttributes = "Observations by Sub System"
		case 11
			LookupFieldAttributes = "Observations by Core Team"
		case 12
			LookupFieldAttributes = "Observations by Component PM"
		case 10
			LookupFieldAttributes = "Observations by State"
		case 13
			LookupFieldAttributes = "Observations by Status"
		case 8
			LookupFieldAttributes = "Observations by Deliverable"
		case 6
			LookupFieldAttributes = "Observations by Developer"
		case 28
			LookupFieldAttributes = "Observations by Owner"
		case 4,14
			LookupFieldAttributes = "Weekly Backlog Graph"
		case 15,16,17,18,19,20
			LookupFieldAttributes = "Weekly Backlog Group Graph"
		case 21
			LookupFieldAttributes = "Weekly Observation Counts Graph"
		case 22,23,24
			LookupFieldAttributes = "Current Risk Observations by Group Graph"
		case 25,26,27
			LookupFieldAttributes = "Average Days Open by Group Graph"
		case 0
			LookupFieldAttributes = "Summary Report"
		end select

	end function

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
	%>
	<form id="frmMain" method="post" action="UpdateProfile.asp">
	<input id="txtPageLayout" name="txtPageLayout" style="display: none" type="text" value="" />
	<input id="txtFieldFilters" name="txtFieldFilters" style="display: none" type="text" value="" />
	<input id="txtUserID" name="txtUserID" style="display: none" type="text" value="<%=CurrentUserID%>" />
	<input style="display: none" id="txtNewProfileName" name="txtNewProfileName" type="text" value="" />
	<input style="display: none" id="txtProfileType" name="txtProfileType" type="text" value="" />
	<input style="display: none" id="txtNewTodayLink" name="txtNewTodayLink" type="text" value="" />
	<input style="display: none" id="txtNewReportFormat" name="txtNewReportFormat" type="text" value="" />
	<input style="display: none" id="txtProfileUpdateID" name="txtProfileUpdateID" type="text" value="" />
	<input style="display: none" id="txtProfileUpdateType" name="txtProfileUpdateType" type="text" value="" />
	</form>
	<iframe style="display: none; width: 100%; height: 300px" id="ProfileFrame" name="ProfileFrame"></iframe>
	<div id="FilterLoadingMessage" style="display: none; position: absolute; background: #FFFFCC; width: 2px; height: 2px; left: 0px; top: 0px; padding: 10px; background: cornsilk; border: 2px ridge gainsboro; z-index: 100; font-family: verdana; font-size: x-small; font-weight: bold; color: #000080">
		Loading&nbsp;Report.&nbsp;&nbsp;Please&nbsp;Wait...
	</div>
	<input style="display: none" id="txtRunReportOK" name="txtRunReportOK" type="text" value="<%=trim(request("RunReportOK"))%>" />
	<input style="display: none" id="txtReturnValue" name="txtReturnValue" type="text" value="" />
</body>
</html>
