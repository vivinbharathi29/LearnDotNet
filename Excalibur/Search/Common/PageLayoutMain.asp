<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
	<title>Page Layout</title>
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
		var VerticalNodes = AvailableColumns.getElementsByTagName("TD");
		for (i = 0; i < VerticalNodes.length; i++) {
			VerticalNodes[i].onmouseup = EnableDropZoneRows;
			VerticalNodes[i].onmousemove = EnableDropZoneRows;
			VerticalNodes[i].onmousedown = NodeMouseDown;
			if (VerticalNodes[i].className != "Pusher")
				VerticalNodes[i].style.cursor = 'hand';
		}
		var RowNodes = ReportRows.getElementsByTagName("TD");
		for (i = 0; i < RowNodes.length; i++) {
			RowNodes[i].onmouseup = EnableDropZone;
			RowNodes[i].onmousemove = EnableDropZone;
			RowNodes[i].onmousedown = NodeMouseDown;
			if (RowNodes[i].className != "Pusher")
				RowNodes[i].style.cursor = 'hand';

		}
		document.documentElement.onmousemove = moveDragMe;
		document.documentElement.onmouseup = dropDragMe;
		SourceTable = AvailableColumnsTable;
		document.onselectstart = function () { return false; };
		blnDragging = false;
		DestinationTable = AvailableColumns;
	}

	function dropDragMe() {
		var NewCell;
		DestinationTable.style.backgroundColor = "white";
		if (blnDragging && typeof (Target) == "unknown") {
			if (SourceTable.id == "AvailableColumns") {
				DestinationRow = SourceTable.insertRow(SourceIndex);
				DestinationCell = DestinationRow.insertCell(0);
				DestinationCell.innerHTML = "<div id=Target class=Dragging FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
			}
			else {
				DestinationCell = SourceTable.rows[0].insertCell(SourceIndex);
				DestinationCell.innerHTML = "<div id=Target class=Dragging style=\"height:100;width:" + DragMe.style.width + ";\" FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
			}
		}
		if (blnDragging && typeof (Target) != "undefined") {
			Target.innerHTML = DragMe.innerHTML;
			Target.className = "";
			Target.id = SourceCell.id;
			DragMe.style.display = "none";
			blnDragging = false;
			NewCell = DestinationCell;//.parentNode.insertCell(DestinationCell.parentNode.cells.length-1);
			NewCell.onmousedown = NodeMouseDown;
			if (DestinationTable.id == "AvailableColumns") {
				NewCell.onmouseenter = EnableDropZoneRows;
				NewCell.onmouseupEnableDropZoneRows;
			}
			else {
				NewCell.onmouseenter = EnableDropZone;
				NewCell.onmouseup = EnableDropZone;
			}
			NewCell.onmouseout = CancelDropZone;
			// if (SourceTable.id == "AvailableColumns")
			//     SourceTable.deleteRow(SourceCell.parentNode.parentNode.rowIndex);
			//else
			//     SourceTable.rows(0).deleteCell(SourceCell.parentNode.cellIndex);
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

	function GetFieldWidth(ElementID) {
		if (ElementID == "Node7")
			return 205;
		else if (ElementID == "Node58")
			return 310;
		else
			return 100;
	}

	function GetFieldHeight(ElementID) {
		return 100;
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
			DragMe.style.width = GetFieldWidth(event.srcElement.id);
			DragMe.style.height = GetFieldHeight(event.srcElement.id);
			DragMe.style.display = "";
			if (SourceTable.id == "AvailableColumns") {
				SourceIndex = SourceCell.parentNode.parentNode.rowIndex;
				SourceTable.deleteRow(SourceCell.parentNode.parentNode.rowIndex);
			}
			else {
				SourceIndex = SourceCell.parentNode.cellIndex;
				SourceTable.rows(0).deleteCell(SourceCell.parentNode.cellIndex);
			}
		}
	}

	function EnableDropZone() {
		var MyTable;
		var MyRow;
		var MyCell;
		var tmpIndex = -1;
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
				window.status = DestinationTable.id
				if (typeof (DestinationCell) != "undefined")
					if (DestinationTable.id == "AvailableColumns")
						DestinationTable.deleteRow(DestinationCell.parentNode.rowIndex);
					else
						DestinationCell.parentNode.deleteCell(DestinationCell.cellIndex);
				//DestinationCell.innerHTML = "";
			}
			else if (MyCell != DestinationCell) {
				if (typeof (DestinationCell) != "undefined") {
					tmpIndex = DestinationCell.cellIndex
					DestinationCell.parentNode.deleteCell(DestinationCell.cellIndex);
				}
			}
			DestinationTable = MyTable;
			if (tmpIndex != -1 && tmpIndex <= MyCell.cellIndex && MyCell.className != "Pusher")
				DestinationCell = MyCell.parentNode.insertCell(MyCell.cellIndex + 1);
			else
				DestinationCell = MyCell.parentNode.insertCell(MyCell.cellIndex);
			//window.status = MyTable.id;//Date() + "_M" + MyCell.cellIndex + "_D" + DestinationCell.cellIndex+ "_T" + tmpIndex;
			DestinationCell.innerHTML = "<div id=Target class=Dragging style=\"height:100;width:" + DragMe.style.width + ";\" FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
			DestinationCell.style.cursor = 'hand';
			//if (MyTable.id != "AvailableColumns")
			MyCell.parentElement.height = parseInt(DragMe.style.height.substring(0, DragMe.style.height.length - 2)) + 8 + "px";
			// }
		}
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
			//if (MyTable.id != "AvailableColumns")
			//    alert(MyTable.id);
			MyTable.style.backgroundColor = "Lavender";
			if (DestinationTable.id != MyTable.id) {
				DestinationTable.style.backgroundColor = "white";
				if (typeof (DestinationCell) != "undefined")
					DestinationCell.parentNode.deleteCell(DestinationCell.cellIndex);
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
			//window.status = MyTable.id;//Date() + "_M" + MyCell.cellIndex + "_D" + DestinationCell.cellIndex+ "_T" + tmpIndex;
			DestinationCell.innerHTML = "<div id=Target class=Dragging FieldHeight=" + DragMe.style.height + " FieldWidth=" + DragMe.style.width + ">&nbsp;</div>";
		}
	}

	function CancelDropZone() {
		// DestinationTable.backgroundColor = "White";
	}

	function CancelDropZone_old() {
		return;
		var MyTable;
		MyTable = event.srcElement
		while (MyTable.tagName != "TABLE" && MyTable.tagName != "") {
			MyTable = MyTable.parentElement;
		}
		MyTable.style.backgroundColor = "white";
		//   alert(MyTable.id);
	}

	//	var count=1;
	//	var oTable = document.getElementById('row2');
	//    var oCells = oTable.rows.item(0).cells;
	//	    var CellsLength = oCells.length;
	//	    for (var j=0; j < CellsLength; j++) 
	//	    {
	//		    oCells.item(j).innerHTML = count++;
	//	    }

	function SaveLayout() {
		var strResults = "";
		var strFinalResults = "";
		var LayoutArray = row1.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strResults = "";
		var LayoutArray = row2.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strResults = "";
		var LayoutArray = row3.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strResults = "";
		var LayoutArray = row4.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strResults = "";
		var LayoutArray = row5.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strResults = "";
		var LayoutArray = row6.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":380:134";
				else if (LayoutArray[i].id.substring(4) == "58")
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":580:134";
				else
					strResults = strResults + "," + LayoutArray[i].id.substring(4) + ":180:134";
		}
		if (strResults.substring(1) != "")
			strFinalResults = strFinalResults + "|" + strResults.substring(1);
		strFinalResults = "21:100%:0|0|2:100%:0|-|0|" + strFinalResults.substring(1) + "|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3";
		frmMain.txtLayout.value = strFinalResults;
		frmMain.txtFieldFilters.value = BuildContentString();
		frmMain.submit();
	}

	function cmdFinish_onclick() {
	    if (!chkMobileCommercial.checked && !chkMobileConsumer.checked && !chkMobileFunctional.checked && !chkDTO.checked)
	        alert("You must select at least one set of products to display.");
	    else {
	        SaveLayout();
	        parent.window.parent.ClosePropertiesDialog();
	    }
	}
	

	function cboLayouts_onchange() {
		if (cboLayouts.options[cboLayouts.selectedIndex].value == "")
			window.location.href = window.location.pathname;
		else
			window.location.href = window.location.pathname + "?LayoutID=" + cboLayouts.options[cboLayouts.selectedIndex].value;
	}

	function DisplayFilterContentSection() {
		var blnFound = false;
		var LayoutArray = AvailableColumns.getElementsByTagName("DIV");
		for (i = 0; i < LayoutArray.length; i++) {
			if (LayoutArray[i].id.substr(0, 4) == "Node")
				if (LayoutArray[i].id.substring(4) == "7")
					blnFound = true;
		}
		if (blnFound)
			ComponentSection.style.display = "none";
		else
			ComponentSection.style.display = "";
	}

	function cmdNext_onclick() {
		DisplayFilterContentSection();
		step1.style.display = "none";
		step2.style.display = "";
		cmdPrevious.disabled = false;
		cmdNext.disabled = true;
		cmdFinish.disabled = false;
	}

	function cmdPrevious_onclick() {
		step1.style.display = "";
		step2.style.display = "none";
		cmdPrevious.disabled = true;
		cmdNext.disabled = false;
		cmdFinish.disabled = true;
	}

	function ShowSubSystem() {
		if (event.srcElement.checked)
			SubSystemRow.style.display = "none";
		else
			SubSystemRow.style.display = "";
		SubSystemRow.scrollIntoView(false);
	}

	function ShowComponent() {
		if (event.srcElement.checked)
			ComponentRow.style.display = "none";
		else
			ComponentRow.style.display = "";
		ComponentRow.scrollIntoView(false);
	}

	function ShowOwner() {
		if (event.srcElement.checked)
			OwnerRow.style.display = "none";
		else
			OwnerRow.style.display = "";
		OwnerRow.scrollIntoView(false);
	}

	function ShowDeveloper() {
		if (event.srcElement.checked)
			DeveloperRow.style.display = "none";
		else
			DeveloperRow.style.display = "";
		DeveloperRow.scrollIntoView(false);
	}

	function ShowComponentPM() {
		if (event.srcElement.checked)
			ComponentPMRow.style.display = "none";
		else
			ComponentPMRow.style.display = "";
		ComponentPMRow.scrollIntoView(false);
	}

	function LookupValues(strField) {
		var strResult;
		var myField;
		var myText;
		if (strField == "SubSystem") {
			myField = txtSubSystemIDs;
			myText = divComponentSubSystemsSelected;
		}
		else if (strField == "CoreTeam") {
			myField = txtCoreTeamIDs;
			myText = divComponentCoreTeamSelected;
		}
		else if (strField == "Type") {
			myField = txtTypeIDs;
			myText = divComponentTypeSelected;
		}
		else if (strField == "Developer") {
			myField = txtDeveloperIDs;
			myText = divComponentDeveloperSelected;
		}
		else if (strField == "ComponentPM") {
			myField = txtComponentPMIDs;
			myText = divComponentComponentPMSelected;
		}
		strResult = window.showModalDialog("BuildSQLLookup.asp?txtField=" + strField + "&ReturnFormat=1&SelectedValues=" + myField.value, "", "dialogWidth:310px;dialogHeight:380px;edge: Raised;center:Yes; help: No;resizable: yes;status: No");
		if (typeof (strResult) != "undefined") {
			myField.value = strResult[0];
			myText.innerText = strResult[1];
		}
	}

	function BuildContentString() {
		var strOutput = "";
		var strNextSet = "";
		if (chkMobileConsumer.checked)
			strOutput = strOutput + " or MobileConsumer=1"
		if (chkMobileCommercial.checked)
			strOutput = strOutput + " or MobileCommercial=1"
		if (chkMobileFunctional.checked)
			strOutput = strOutput + " or MobileFunctional=1"
		if (chkDTO.checked)
			strOutput = strOutput + " or DTO=1"
		if (strOutput != "")
			strOutput = " and (" + strOutput.substring(4) + ") ";
		strNextSet = "";
		if (ComponentRow.style.display != "none" && !chkAllComponents.checked) {
			if (txtCoreTeamIDs.value != "")
				strNextSet = strNextSet + " or CoreteamID in (" + txtCoreTeamIDs.value + ") "
			if (txtSubSystemIDs.value != "")
				strNextSet = strNextSet + " or (LinkType=1 and LinkID in (" + txtSubSystemIDs.value + ")) "
			if (txtTypeIDs.value != "")
				strNextSet = strNextSet + " or (LinkType=2 and LinkID in (" + txtTypeIDs.value + ")) "
			if (txtDeveloperIDs.value != "")
				strNextSet = strNextSet + " or (LinkType=3 and LinkID in (" + txtDeveloperIDs.value + ")) "
			if (txtComponentPMIDs.value != "")
				strNextSet = strNextSet + " or (LinkType=4 and LinkID in (" + txtComponentPMIDs.value + ")) "
		}
		if (strNextSet != "")
			strOutput = strOutput + "| and (" + strNextSet.substring(4) + ") "
		return strOutput;
	}

	function cmdCancel_onclick() {
	    window.parent.close();
	    parent.window.parent.ClosePropertiesDialog();
	}
	function CloseIframeDialog() {
	    var iframeName = window.name;
	    if (iframeName != '') {
	        parent.window.parent.ClosePropertiesDialog();
	    } else {
	        window.close();
	    }
	}
	//-->
	</script>
	<style>
		td
		{
			FONT-FAMILY: Verdana;
			FONT-SIZE: xx-small;
		}
		body
		{
			FONT-FAMILY: Verdana;
			FONT-SIZE: xx-small;
			background-color: Lavender;
		}
		#MainDragArea, #Step2Panel
		{ /* Main container for this script */
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
		{ /* Drag container */
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
			COLOR: blue;
		}
		A:link
		{
			COLOR: blue;
		}
		A:hover
		{
			COLOR: red;
		}
	</style>
</head>
<body onload="window_onload()">
	<%
	dim cn, rs, strSQL, cm, p, j
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

	strAvailableColumnIDs = ",54,55,50,51,57,31,7,5,52,46,47,27,9,6,10,11,12,4,29,14,26,16,58,15,28,18,45,48,49,20,22,24,13,8," 

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

	if request("LayoutID") = "" then
		rs.open "spGetEmployeeUserSettings " & clng(currentuserid) & ",8"
		if not (rs.eof and rs.bof) then
			strProfilePageLayout = trim(rs("Setting") & "")
		end if
		rs.Close
	elseif request("LayoutID") = "1" then 'Classic
		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:150:134,54:180:134,27:180:134,20:180:134,22:170:134,31:150:134|28:150:134,55:180:134,14:180:134,7:380:134,12:150:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|42:100%:3|40:100%:0|3:100%:3"
	elseif request("LayoutID") = "2" then 'All Fields
'        strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:170:134,27:170:134,11:170:134,14:170:134,26:170:134,8:170:134,7:380:134|54:180:134,55:180:134,50:170:134,51:170:134,18:170:134,45:170:134,4:170:134,29:170:134,20:170:134|28:170:134,5:170:134,52:170:134,9:170:134,6:170:134,24:170:134,13:170:134,22:170:134|12:170:134,46:170:134,47:170:134,48:170:134,49:170:134,10:170:134,15:170:134,31:170:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:180:134,54:180:134,20:180:134,11:180:134,14:180:134,26:180:134,31:180:134|28:180:134,55:180:134,12:180:134,8:180:134,27:180:134,7:380:134|15:180:134,22:180:134,10:180:134,9:180:134,6:180:134,5:180:134,52:180:134|4:180:134,29:180:134,24:180:134,13:180:134,18:180:134,45:180:134|46:180:134,47:180:134,48:180:134,49:180:134,50:180:134,51:180:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
	elseif request("LayoutID") = "3" then 'Empty layout
		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0||-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
	elseif request("LayoutID") = "4" then 'Popular - Mobile
		'strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:180:134,14:180:134,12:180:134,20:180:134,11:180:134,10:180:134,31:180:134|28:180:134,54:180:134,55:180:134,8:180:134,27:180:134,22:180:134,7:380:134|4:180:134,18:180:134,9:180:134,5:180:134,24:180:134,48:180:134,46:180:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:180:134,14:180:134,12:180:134,20:180:134,7:380:134,31:180:134|28:180:134,54:180:134,55:180:134,8:180:134,27:180:134,22:180:134,11:180:134|4:180:134,18:180:134,9:180:134,5:180:134,24:180:134,48:180:134,46:180:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
	elseif request("LayoutID") = "5" then 'Popular - DTO
		 strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|58:580:134,14:180:134,4:180:134,12:180:134,20:180:134,31:180:134|54:180:134,55:180:134,8:180:134,22:180:134,11:180:134,7:380:134|18:180:134,9:180:134,5:180:134,24:180:134,48:180:134,46:180:134,50:180:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|43:100%:0|42:100%:3|40:100%:0|3:100%:3"
	end if
	if strProfilePageLayout = "" then 'Default to classic
		strProfilePageLayout = "21:100%:0|0|2:100%:0|-|0|16:150:134,54:180:134,27:180:134,20:180:134,22:170:134,31:150:134|28:150:134,55:180:134,14:180:134,7:380:134,12:150:134|-|0|44:120:0,17:240:0,39:300:0|30:120:0,32:220:0,33:200:0|25:120:0,37:220:0,34:200:0|19:120:0,36:220:0,35:200:0|23:120:0,38:220:0,56:150:0|-|0|41:100%:0|42:100%:3|40:100%:0|3:100%:3"
	end if
	LayoutRows = split(strProfilePageLayout,"|")

	'response.write ">" & strProfilePageLayout & "<br/><br/>"
	for i = 5 to ubound(LayoutRows)
		strRowFields = ""
		if left(LayoutRows(i),1) = "-" then
			exit for
		end if
		FieldArray = split(LayoutRows(i),",")
		for each strField in FieldArray
			FieldPropertyArray = split(strField,":")
			if trim(FieldPropertyArray(0))="7" then
				strRowFields = strRowFields & "<td><div FieldHeight=100 FieldWidth=205 id=Node7 style=""width:205px;height:100px"">Component</div></td>"
			elseif trim(FieldPropertyArray(0))="58" then
				FieldAttributes = LookupFieldAttributes(FieldPropertyArray(0))
				if FieldAttributes <> "" then
					FieldAttributeArray = split(FieldAttributes,"|")
					strRowFields = strRowFields & "<td><div FieldHeight=100 FieldWidth=" & trim(FieldAttributeArray(1)) & " id=Node" & trim(FieldPropertyArray(0)) & " style=""width:" & trim(FieldAttributeArray(1)) & "px;height:100px"">" & FieldAttributeArray(0) & "</div></td>"
				end if
			else
				FieldAttributes = LookupFieldAttributes(FieldPropertyArray(0))
				if FieldAttributes <> "" then
					FieldAttributeArray = split(FieldAttributes,"|")
					strRowFields = strRowFields & "<td><div FieldHeight=100 FieldWidth=100 id=Node" & trim(FieldPropertyArray(0)) & " style=""width:100px;height:100px"">" & FieldAttributeArray(0) & "</div></td>"
				end if
			end if
			strAvailableColumnIDs = replace(strAvailableColumnIDs,"," & trim(FieldPropertyArray(0)) & ",",",")
		next
		if i = 5 then
			strRow1Fields = strRowFields
		elseif i = 6 then
			strRow2Fields = strRowFields
		elseif i = 7 then
			strRow3Fields = strRowFields
		elseif i = 8 then
			strRow4Fields = strRowFields
		elseif i = 9 then
			strRow5Fields = strRowFields
		elseif i = 10 then
			strRow6Fields = strRowFields
		end if
	next

	AvailableRowFields = ""
	FieldArray = split(strAvailableColumnIDs,",")
	for i = 0 to ubound(FieldArray)
		if trim(Fieldarray(i)) <> "" then
			if trim(Fieldarray(i)) = "7" then
				AvailableRowFields = AvailableRowFields & "<tr><td><div FieldHeight=100 FieldWidth=205 id=Node7>Components</div></td></tr>"
			else
				FieldAttributes = LookupFieldAttributes(Fieldarray(i))
				if FieldAttributes <> "" then
					FieldAttributeArray = split(FieldAttributes,"|")
					AvailableRowFields = AvailableRowFields & "<tr><td><div FieldHeight=100 FieldWidth=100 id=Node" & trim(Fieldarray(i)) & ">" & FieldAttributeArray(0) & "</div></td></tr>"
				end if
			end if
		end if
	next

	'Load Division User Settings
	strMobileConsumerChecked = " "
	strMobileCommercialChecked = " "
	strMobileFunctionalChecked = " "
	strDTOChecked = " "
	strDeveloperIDList = ""
	strComponentPMIDList = ""
	strSubSystemIDList = ""
	strCoreTeamIDList = ""
	strTypeIDList = ""
	strSelectedDeveloperList = "All Developers"
	strSelectedComponentPMList = "All Component PMs"
	strSelectedSubSystemList = "All Sub Systems"
	strSelectedCoreTeamList = "All Core Teams"
	strSelectedTypeList = "All Component Types"

	rs.open "spGetEmployeeUserSettings " & clng(currentuserid) & ",9" ,cn,adOpenStatic
	if rs.eof and rs.bof then
		if trim(CurrentUserDivision) = "1" then
			strMobileConsumerChecked = " checked "
			strMobileCommercialChecked = " checked "
			strMobileFunctionalChecked = " checked "
			strDTOChecked = " "
		else
			strMobileConsumerChecked = " "
			strMobileCommercialChecked = " "
			strMobileFunctionalChecked = " "
			strDTOChecked = " checked "
		end if
		rs.Close
	else
		UserSettingArray = split(rs("setting") & "","|")
		rs.Close
		if instr(UserSettingArray(0),"MobileConsumer=1") <> 0 then
			strMobileConsumerChecked = " checked "
		end if
		if instr(UserSettingArray(0),"MobileCommercial=1") <> 0 then
			strMobileCommercialChecked = " checked "
		end if
		if instr(UserSettingArray(0),"MobileFunctional=1") <> 0 then
			strMobileFunctionalChecked = " checked "
		end if
		if instr(UserSettingArray(0),"DTO=1") <> 0 then
			strDTOChecked = " checked "
		end if

		if ubound(UserSettingArray) = 0 then
			DisplayComponentRow = "none"
			strAllComponentsChecked = "checked"
		elseif trim(UserSettingArray(1)) = "" then
			DisplayComponentRow = "none"
			strAllComponentsChecked = "checked"
		else
			DisplayComponentRow = ""
			strAllComponentsChecked = ""

			FieldFilterValues = split (mid(replace(lcase(UserSettingArray(1)),"coreteamid in","0|"),5),"or")
			for i = 0 to ubound(FieldFilterValues)
				valuearray=split(replace(replace(replace(replace(replace(lcase(FieldFilterValues(i)),")","") ,"(","") ," ",""),"andlinkidin","|"),"linktype=","") ,"|")
				if ubound(valuearray) = 1 then
					if trim(valuearray(0)) = "0" then
						strSQL =  "Select Name from prs.dbo.deliverablecoreteam with (NOLOCK) where id in (" & scrubsql(valuearray(1)) & ") order by name;"
						strSelectedCoreTeamList = ""
						rs.open strSQL,cn,adOpenStatic
						do while not rs.EOF
							if rs("Name") & "" = "None" then
								strSelectedCoreTeamList = strSelectedCoreTeamList & "; No Core Team Assigned" 
							else
								strSelectedCoreTeamList = strSelectedCoreTeamList & "; " & rs("Name")
							end if
							rs.MoveNext
						loop
						rs.Close
						if strSelectedCoreTeamList <> "" then
							strSelectedCoreTeamList = mid(strSelectedCoreTeamList,3)
						end if
						strCoreTeamIDList = valuearray(1)
					elseif trim(valuearray(0)) = "1" then
						strSQl = "Select Name from HOUSIREPORT01.SIO.dbo.list_subsystem with (NOLOCK) where id in (" & scrubsql(valuearray(1)) & ") order by name;"
						strSelectedSubSystemList = ""
						rs.open strSQL,cn,adOpenStatic
						do while not rs.EOF
							strSelectedSubSystemList = strSelectedSubSystemList & "; " & rs("Name")
							rs.MoveNext
						loop
						rs.Close
						if strSelectedSubSystemList <> "" then
							strSelectedSubSystemList = mid(strSelectedSubSystemList,3)
						end if
						strSubSystemIDList = valuearray(1)
					elseif trim(valuearray(0)) = "2" then
						strSQl = "Select Name from HOUSIREPORT01.SIO.dbo.list_type with (NOLOCK) where id in (" & scrubsql(valuearray(1)) & ") order by name;"
						strSelectedTypeList = ""
						rs.open strSQL,cn,adOpenStatic
						do while not rs.EOF
							strSelectedTypeList = strSelectedTypeList & "; " & rs("Name")
							rs.MoveNext
						loop
						rs.Close
						if strSelectedTypeList <> "" then
							strSelectedTypeList = mid(strSelectedTypeList,3)
						end if
						strTypeIDList = valuearray(1)
					elseif trim(valuearray(0)) = "3" then
						strSQl = "Select Name from HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) where userid in (" & scrubsql(valuearray(1)) & ") order by name;"
						strSelectedDeveloperList = ""
						rs.open strSQL,cn,adOpenStatic
						do while not rs.EOF
							strSelectedDeveloperList = strSelectedDeveloperList & "; " & rs("Name")
							rs.MoveNext
						loop
						rs.Close
						if strSelectedDeveloperList <> "" then
							strSelectedDeveloperList = mid(strSelectedDeveloperList,3)
						end if
						strDeveloperIDList = valuearray(1)
					elseif trim(valuearray(0)) = "4" then
						strSQl = "Select Name from HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) where userid in (" & scrubsql(valuearray(1)) & ") order by name;"
						strSelectedComponentPMList = ""
						rs.open strSQL,cn,adOpenStatic
						do while not rs.EOF
							strSelectedComponentPMList = strSelectedComponentPMList & "; " & rs("Name")
							rs.MoveNext
						loop
						rs.Close
						if strSelectedComponentPMList <> "" then
							strSelectedComponentPMList = mid(strSelectedComponentPMList,3)
						end if
						strComponentPMIDList = valuearray(1)
					end if
				end if
			next
		end if
	end if

	'Load Field Filter Settigns
	%>
	<span style="font: bold small verdana">Default Page Layout</span>
	<br/>
	<br/>
	<div id="step1">
		<table style="width: 100%; border: none; border-spacing: 0; border-collapse: collapse; padding: 0">
			<tr>
				<td style="font: x-small verdana">Step 1: Drag fields to create a custom default page layout.<br/><br/></td>
				<td style="vertical-align: bottom; float: right">
					<div style="font: x-small verdana; margin-bottom: 4px; float: right">
						Start with:
		<select id="cboLayouts" onchange="javascript:cboLayouts_onchange();">
			<option selected value="">My Default Layout</option>
			<%if trim(request("LayoutID")) = "1" then%>
			<option selected value="1">Classic Layout</option>
			<%else%>
			<option value="1">Classic Layout</option>
			<%end if%>
			<%if trim(request("LayoutID")) = "4" then%>
			<option selected value="4">Popular Fields - Mobile</option>
			<%else%>
			<option value="4">Popular Fields - Mobile</option>
			<%end if%>
			<%if trim(request("LayoutID")) = "5" then%>
			<option selected value="5">Popular Fields - DTO</option>
			<%else%>
			<option value="5">Popular Fields - DTO</option>
			<%end if%>
			<%if trim(request("LayoutID")) = "3" then%>
			<option selected value="3">Empty Layout</option>
			<%else%>
			<option value="3">Empty Layout</option>
			<%end if%>
		</select>
					</div>
				</td>
			</tr>
		</table>
		<div id="MainDragArea" style="height: 400px">
			<table style="border: none; border-spacing: 0; border-collapse: collapse; padding: 0">
				<tr>
					<td id="AvailableBox" style="vertical-align: top; Height: 100%">
						<div style="border: none; height: 394px; overflow-y: scroll">
							<table id="AvailableColumnsTable" style="Height: 100%; width: 175px; border: none; border-spacing: 0; border-collapse: collapse; padding: 0; background-color: white">
								<tr style="height: 15px">
									<td style="white-space: nowrap; background-color: #317082; color: white; font-weight: bold">Available Fields</td>
								</tr>
								<tr>
									<td style="vertical-align: top; white-space: nowrap">
										<table id="AvailableColumns" style="width: 100%; height: 100%; border: none; border-spacing: 0; border-collapse: collapse; padding: 0">
											<%=AvailableRowFields%>
											<tr style="Height: 100%">
												<td class="Pusher">&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
					</td>
					<td id="DropZone" style="vertical-align: top; width: 100%">
						<div style="border: none; height: 394px; overflow-y: scroll">
							<table style="border: none; border-spacing: 0; border-collapse: collapse; padding: 0; height: 800px; width: 100%">
								<tr style="height: 15px">
									<td style="white-space: nowrap; background-color: #317082; font-weight: bold; color: white">Page Layout - Listbox Section</td>
								</tr>
								<tr>
									<td style="vertical-align: top" id="ReportRows">
										<table id="row1" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow1Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
											</tr>
										</table>
										<table id="row2" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow2Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
											</tr>
										</table>
										<table id="row3" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow3Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
											</tr>
										</table>
										<table id="row4" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow4Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
											</tr>
										</table>
										<table id="row5" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow5Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
											</tr>
										</table>
										<table id="row6" style="width: 100%; border: 1px solid DarkGray; border-collapse: collapse; padding: 0; background-color: white; margin: 10px">
											<tr style="height: 30px">
												<%=strRow6Fields%>
												<td class="Pusher" style="width: 100%" onmouseup="javascript:EnableDropZone();" onmousemove="javascript:EnableDropZone();" onmouseout="javascript:CancelDropZone();">&nbsp;</td>
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
	<div id="step2" style="display: none">
		<table style="width: 100%; border-collapse: collapse">
			<tr>
				<td style="font: x-small verdana">Step 2: Choose what items are displayed for each field.<br/><br/></td>
			</tr>
		</table>
		<div id="Step2Panel" style="overflow-y: scroll; margin-left: 10px; height: 400px">
			<div style="padding: 2px; margin: 3px; background: #EEE; border: 1px solid black">
				<span style="text-decoration: underline">Products</span> <span style="color: red; font-weight: bold">*</span>
				<br/>
				<input id="chkDTO" <%=strDTOChecked%> type="checkbox" />
				DTO Products<br/>
				<input id="chkMobileConsumer" <%=strMobileConsumerChecked%> type="checkbox" />
				Mobile Products - Consumer<br/>
				<input id="chkMobileCommercial" <%=strMobileCommercialChecked%> type="checkbox" />
				Mobile Products - Commercial<br/>
				<input id="chkMobileFunctional" <%=strMobileFunctionalChecked%> type="checkbox" />
				Mobile Products - Functional Test<br/>
				<div id="ComponentSection">
					<br/>
					<span style="text-decoration: underline">Components</span><br/>
					<table style="border: none; width: 100%">
						<tr>
							<td>
								<input id="chkAllComponents" <%=strAllComponentsChecked%> type="checkbox" onclick="javascript: ShowComponent();" />
								All components on selected products.
							</td>
						</tr>
						<tr id="ComponentRow" style="display: <%=DisplayComponentRow%>">
							<td style="width: 100%">
								<fieldset style="background-color: White; margin-left: 17px; margin-right: 15px;">
									<legend></legend>
									<div style="margin: 5px 0px 5px 2px; color: green">Components are filtered by the selected products automatically.  Choose any other filters you wish to apply to this field.</div>
									<table style="width: 100%; border: none; padding: 2px; border-spacing: 0px; margin-bottom: 2px;">
										<tr>
											<td style="white-space: nowrap; text-decoration: underline; font-weight: bold">Filter</td>
											<td style="width: 100%; font-weight: bold; text-decoration: underline">Items Selected</td>
										</tr>
										<tr>
											<td style="white-space: nowrap; vertical-align: top"><a href="javascript:LookupValues('SubSystem');">Sub System</a></td>
											<td style="width: 100%" id="divComponentSubSystemsSelected"><%=strSelectedSubSystemList%></td>
											<td>
												<input style="display: none" id="txtSubSystemIDs" name="txtSubSystemIDs" type="text" value="<%=strSubSystemIDList%>"/></td>
										</tr>
										<tr>
											<td style="white-space: nowrap; vertical-align: top"><a href="javascript:LookupValues('CoreTeam');">Core Team</a></td>
											<td style="width: 100%" id="divComponentCoreTeamSelected"><%=strSelectedCoreTeamList%></td>
											<td>
												<input style="display: none" id="txtCoreTeamIDs" name="txtCoreTeamIDs" type="text" value="<%=strCoreTeamIDList%>"/></td>
										</tr>
										<tr>
											<td style="white-space: nowrap; vertical-align: top"><a href="javascript:LookupValues('Type');">Component Type</a>&nbsp;&nbsp;&nbsp;</td>
											<td style="width: 100%" id="divComponentTypeSelected"><%=strSelectedTypeList%></td>
											<td>
												<input style="display: none" id="txtTypeIDs" name="txtTypeIDs" type="text" value="<%=strTypeIDList%>"/></td>
										</tr>
										<tr>
											<td style="white-space: nowrap; vertical-align: top"><a href="javascript:LookupValues('Developer');">Developers</a></td>
											<td style="width: 100%" id="divComponentDeveloperSelected"><%=strSelectedDeveloperList%> </td>
											<td>
												<input style="display: none" id="txtDeveloperIDs" name="txtDeveloperIDs" type="text" value="<%=strDeveloperIDList%>"/></td>
										</tr>
										<tr>
											<td style="white-space: nowrap; vertical-align: top"><a href="javascript:LookupValues('ComponentPM');">Component PMs</a></td>
											<td style="width: 100%" id="divComponentComponentPMSelected"><%=strSelectedComponentPMList%></td>
											<td>
												<input style="display: none" id="txtComponentPMIDs" name="txtComponentPMIDs" type="text" value="<%=strComponentPMIDList%>"/></td>
										</tr>
									</table>
								</fieldset>
							</td>
						</tr>
					</table>
					<br/>
					<br/>
				</div>
			</div>
		</div>
	</div>
	<hr/>
	<table style="border: none; border-spacing: 0; border-collapse: collapse; float: right">
		<tr>
			<td>
				<input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" onclick="return cmdCancel_onclick()"/></td>
			<td>&nbsp;&nbsp;</td>
			<td>
				<input type="button" disabled value="<< Previous" id="cmdPrevious" name="cmdPrevious" onclick="return cmdPrevious_onclick()"/></td>
			<td>
				<input type="button" value="Next >>" id="cmdNext" name="cmdNext" onclick="return cmdNext_onclick()"/></td>
			<td>&nbsp;&nbsp;</td>
			<td>
				<input type="button" disabled value="Finish" id="cmdFinish" name="cmdFinish" onclick="return cmdFinish_onclick()"/></td>
		</tr>
	</table>
	<div id="DragMe" style="display: none"></div>
	<%
	set rs = nothing
	cn.Close
	set cn=nothing

	function LookupFieldAttributes(ID)
		select case clng(ID)
		case 54
			LookupFieldAttributes = "Affected Product|180"
		case 55
			LookupFieldAttributes = "Affected State|180"
		case 50
			LookupFieldAttributes = "Approver|180"
		case 51
			LookupFieldAttributes = "Approver Group|180"
		case 31
			LookupFieldAttributes = "Columns|180"
		case 5
			LookupFieldAttributes = "Component PM|180"
		case 52
			LookupFieldAttributes = "Component PM Group|180"
		case 46
			LookupFieldAttributes = "Comp. Test Lead|180"
		case 47
			LookupFieldAttributes = "Comp. Test Lead Group|180"
		case 27
			LookupFieldAttributes = "Core Team|180"
		case 9
			LookupFieldAttributes = "Developer|180"
		case 6
			LookupFieldAttributes = "Developer Group|180"
		case 10
			LookupFieldAttributes = "Feature|180"
		case 11
			LookupFieldAttributes = "Frequency|180"
		case 12
			LookupFieldAttributes = "Gating Milestone|180"
		case 4
			LookupFieldAttributes = "Originator|180"
		case 29
			LookupFieldAttributes = "Originator Group|180"
		case 14
			LookupFieldAttributes = "Owner|180"
		case 26
			LookupFieldAttributes = "Owner Group|180"
		case 16
			LookupFieldAttributes = "Primary Product|180"
		case 58
			LookupFieldAttributes = "Product and Version|310"
		case 15
			LookupFieldAttributes = "Product Family|180"
		case 28
			LookupFieldAttributes = "Product Group|180"
		case 18
			LookupFieldAttributes = "Product PM|180"
		case 45 
			LookupFieldAttributes = "Product PM Group|180"
		case 48
			LookupFieldAttributes = "Prod. Test Lead|180"
		case 49
			LookupFieldAttributes = "Prod. Test Lead Group|180"
		case 20
			LookupFieldAttributes = "State|180"
		case 22
			LookupFieldAttributes = "Sub System|180"
		case 24
			LookupFieldAttributes = "Tester|180"
		case 57
			LookupFieldAttributes = "Assigned To|180"
		case 13
			LookupFieldAttributes = "Tester Group|180"
		case 8
			LookupFieldAttributes = "Type|180"
		end select
	end function

		function ScrubSQL(strWords) 
		dim badChars 
		dim newChars 
		dim i
	'	strWords=replace(strWords,"'","''")
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		ScrubSQL = newChars 
	end function 
	%>
	<form id="frmMain" method="post" action="PageLayoutSave.asp">
		<input id="txtLayout" name="txtLayout" style="display: none" type="text" value=""/>
		<input id="txtFieldFilters" name="txtFieldFilters" style="display: none" type="text" value=""/>
		<input id="txtUserID" name="txtUserID" style="display: none" type="text" value="<%=CurrentUserID%>"/>
	</form>
</body>
</html>
