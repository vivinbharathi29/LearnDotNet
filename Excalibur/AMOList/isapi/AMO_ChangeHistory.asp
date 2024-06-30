<%@ Language=VBScript %>
<% Option Explicit 
	
    Response.Buffer = False
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"

    Server.ScriptTimeout = 6000
%>
<!------------------------------------------------------------------- 
'Description: AMO DATA
'----------------------------------------------------------------- //-->    
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataAVL.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
Call ValidateSession
dim sHeader, sHelpFile, strTabName, sErr, sName, oRsBusSeg, sPrevBusSegID
dim strSortField, strSortOrder, strPrevTab, strSincePublication, strLastDays
dim strUpdateDateFrom, strUpdateDateTo, strBlindFrom, strBlindTo, strDiscFrom, strDiscTo, strBuildFilter
dim nNumRequest, nNumTotalRequest, nMode
dim lngCount
dim bCreate, bView, bUpdate, bDelete
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim oSvr, oErr, oRs, oRsAMOCategory, oRsAMO, oRsChangetype, sHPPartNo, strError
dim oRsUpdateDates, oRsUpdaters
dim bShowcheckbox
dim sModuleIDs
dim arrColors
dim sSearchFilter, slastpublishSCMID, stimerange, sOriginalSelectedIDs
dim sPrevAMOCategoryID, sPrevChangeType, sPrevPartNo, sPrevUpdater, sPrevUpdatedate, sPrevREvID


arrColors = Array("white", "#CCCCCC")

strTabName = "Change History"


'get permissions
GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete


nMode = Request.QueryString("nMode")

if sErr = "" then
	strLastDays = Request.Form("txtLastdays")
end if

if sErr = "" then
	'create overall page security variables.
	bCreate = bAMOCreate
	bUpdate = bAMOUpdate
	bDelete = false
	bView = bAMOView
	
	sHelpFile = ""
	sHeader = "After Market Option List"
end if

if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	'set oRsAMOCategory = oSvr.AMO_ChangeHistory_AMOCategories(Application("REPOSITORY"))
    set oRsAMOCategory = GetMOLCategory(33)	
	if (oRsAMOCategory is nothing) then
		strError = "Recordset Error: oRsAMOCategory, AMO_ChangeHistory.asp"
	    Response.Write(strError)
        Response.End()
	end if
end if

if sErr = "" then
	set oRsChangetype = oSvr.AMO_ChangeHistory_AllChangeTypes(Application("REPOSITORY"))
	'if (oRsChangetype is nothing) then
		'strError = "Recordset Error: oRsChangetype, AMO_ChangeHistory.asp"
	    'Response.Write(strError)
        'Response.End()
	'end if
end if

if sErr = "" then
	set oRsUpdaters = oSvr.AMO_ChangeHistory_AllUpdaters(Application("REPOSITORY"))
	'if (oRsUpdaters is nothing) then
		'strError = "Recordset Error: oRsUpdaters, AMO_ChangeHistory.asp"
	    'Response.Write(strError)
        'Response.End()
	'end if
end if

if len(Request.Form("lbxChangeType")) > 0 and Request.Form("lbxChangeType") <> "0" then
	sPrevChangeType = Request.Form("lbxChangeType")
	sSearchFilter = sSearchFilter & " and  AC.ChangeTypeID= " & sPrevChangeType
end if 



if len(Request.Form("lbxAMOCategory")) > 0  and Request.Form("lbxAMOCategory") <> "0" then
	sPrevAMOCategoryID = Request.Form("lbxAMOCategory")
	sSearchFilter =sSearchFilter &  " and FC.FeatureCategoryID= " & sPrevAMOCategoryID
end if 


if len(Request.Form("lbxUpdaters")) >  0 then	
	sPrevUpdater = Request.Form("lbxUpdaters")
	sSearchFilter =sSearchFilter &  " and  AC.Updater= '" & sPrevUpdater & "'"
end if 


strUpdateDateFrom = Request.Form("txtUpdateDateFrom")
strUpdateDateTo = Request.Form("txtUpdateDateTo")
if strUpdateDateFrom <> "" and strUpdateDateTo <> "" then
	sSearchFilter = sSearchFilter & " and ( AC.TimeChanged >= convert(datetime,'"  & strUpdateDateFrom & "') and AC.TimeChanged <= dateadd(day,1, convert(datetime,'"  & strUpdateDateTo & "')))"
elseif strUpdateDateFrom <> "" and strUpdateDateTo = "" then
	sSearchFilter = sSearchFilter & " and (AC.TimeChanged >= convert(datetime,'"  & strUpdateDateFrom & " '))"
elseif strUpdateDateFrom = "" and strUpdateDateTo <> "" then
	sSearchFilter = sSearchFilter & " and (AC.TimeChanged <=  dateadd(day,1, convert(datetime,'"  & strUpdateDateTo & "')))"
end if
				

if len(Request.Form("txtLastdays")) > 0 then
	sSearchFilter  = sSearchFilter & " and datediff(day, AC.timechanged, '" & date() & "' )< " & Request.Form("txtLastdays")
	stimerange = Request.Form("txtLastdays")
	
else
	sSearchFilter  = sSearchFilter & " and datediff(day, AC.timechanged, '" & date() & "' )< 1" 
	stimerange = 1
end if



if len(Request.Form("txtHPPartNo")) > 0 then
		sSearchFilter  = sSearchFilter & " and AO.BluePartNo like '" & Request.Form("txtHPPartNo") & "'"
		sHPPartNo = Request.Form("txtHPPartNo")
end if 
	


sPrevBusSegID = Request.Form("lbxBusSeg")
if sPrevBusSegID = 0 then
	sSearchFilter = sSearchFilter & "|" 
else
	sSearchFilter = sSearchFilter & "|" & sPrevBusSegID
end if

if sErr = "" then
	set oRs = oSvr.AMO_ChangeHistory(Application("REPOSITORY"), sSearchFilter)	
    '--NOTE: got to add code that will display filter when no records are returned
	if oRs is nothing then
		strError = "Recordset Error: oRs, AMO_ChangeHistory.asp"
	    Response.Write(strError)
        Response.End()
	end if
end if

if sErr = "" then
	strSortOrder = Request.Form("sortorder")
	strSortField = Request.Form("sortfield")
	
	if strSortField = "" then
		strSortField = "TimeChanged"
	end if
	
	
	if strSortOrder = "" then
		strSortOrder = "DESC"
	end if

	if oRs.RecordCount > 0 then
		oRs.Sort = strSortField & " " & strSortOrder
	end if
	
end if
%>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sHeader%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sHeader%> - Change History</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/wizard%20style.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    //var newWindow;
    //var oPopup = window.createPopup();
    //var SelectedRow;


    function BtnSave_Onlick() {
        thisform.action = "AMO_Save.asp?nMode=6";
        thisform.submit();
    }


    function SortField(fieldname, sortorder) {
        var nAMOCategoryID = -1
        var ASCENDING = "<%= ASCENDING %>"
        var DESCENDING = "<%= DESCENDING %>"
        thisform.sortfield.value = fieldname
        if (fieldname == "<%= strSortField %>") {
            // same field so flip the sort order
            if (sortorder == ASCENDING) {
                thisform.sortorder.value = DESCENDING
            } else {
                thisform.sortorder.value = ASCENDING
            }
        } else {
            // new field so start the sort order as ascending
            thisform.sortorder.value = ASCENDING
        }
        if (thisform.lbxAMOCategory != null)
            nAMOCategoryID = thisform.lbxAMOCategory.value;
        thisform.action = "AMO_ChangeHistory.asp";
        return thisform.submit();
    }


    function ApplyFilter() {
        if (ValidateFilter()) {
            thisform.action = "AMO_ChangeHistory.asp?nMode=1";
            return thisform.submit();
        }
    }



    function btnDeselectAll_Click() {
        if (document.all("chkHideFromSCM").length > 1) {
            for (var i = 0; i < document.all("chkHideFromSCM").length; i++) {
                document.all("chkHideFromSCM").item(i).checked = false
            }
        } else {
            document.all("chkHideFromSCM").checked = false
        }
    }

    function btnSelectAll_Click() {
        if (document.all("chkHideFromSCM").length > 1) {
            for (var i = 0; i < document.all("chkHideFromSCM").length; i++) {
                document.all("chkHideFromSCM").item(i).checked = true;
            }
        } else {
            document.all("chkHideFromSCM").checked = true;
        }
    }

    function ValidateFilter() {

        if (thisform.txtLastdays.value.length > 0) {
            if (!isPositiveInteger(thisform.txtLastdays.value, false)) {
                alert("Please enter a valid number.");
                return false;
            }
        }


        if (!checkDate(thisform.txtUpdateDateFrom, "Update Date (From)", true))
            return false;

        if (!checkDate(thisform.txtUpdateDateTo, "Update Date (To)", true))
            return false;


        if (thisform.txtUpdateDateFrom.value != '' && thisform.txtUpdateDateTo.value != '') {
            var startdate = newDate(thisform.txtUpdateDateFrom.value)
            var enddate = newDate(thisform.txtUpdateDateTo.value)
            if ((startdate.getTime() - enddate.getTime()) > 0) {
                alert("The update Date (To) is older than the PHweb (General) Availability (GA) (From), please re-enter the dates.");
                thisform.txtUpdateDateFrom.focus();
                return (false);
            }
        }


        return true;
    }

    function changeSum(nChangeHistoryID, sNewValue) {
        var sNewAVReasonID;
        var sNewAVReason;

        sNewAVReasonID = thisform.newAVReassonID.value;
        sNewAVReason = thisform.newAVReasson.value;
        if (sNewValue.length > 256) {
            alert("Change reason is limited to 256 characters");
            return;
        }
        // if the unique field is not in the string then add it
        if (sNewAVReasonID == "" || sNewAVReasonID.indexOf("," + nChangeHistoryID + ",") == -1) {
            sNewAVReasonID += "," + nChangeHistoryID + ",";
            thisform.newAVReassonID.value = sNewAVReasonID;
            sNewAVReason += "|" + sNewValue + "|";
            thisform.newAVReasson.value = sNewAVReason;
        }
    }
    //-->
</SCRIPT>
</HEAD>

<BODY LANGUAGE=javascript bgcolor="#FFFFFF">
<!-- #include file="../library/includes/popup.inc" -->
<h1 class="page-title"><%=sHeader%></h1>
<FORM name=thisform method=post>
<% 'insert the header, global navigation and overview links

'steven fix #5532 
'Response.write BuildHelpNoLine("AMO: " & sName, sHelpfile)
'Response.write BuildHelpNoLine("After Market Option List" & sName, sHelpfile)
Response.Write ""
WriteTabs strTabName

if sErr <> "" then
	Response.Write sErr
else
	%>
	<div ID=erroutputArea></div>
	<TABLE border=0 cellPadding=2 cellSpacing=2 width=100% border="2" style="font-family:Verdana,Sans-Serif !important;">
		<%
		
			'set oErr = GetMOLCategory(oRsBusSeg, 28)
	        set oRsBusSeg = GetMOLCategory(34)	
	        if oRsBusSeg is Nothing then
		        Response.Write("Recordset error: oRsBusSeg")
		        Response.End()
	        end if

			Response.Write "<tr><td align=left width=""20%""><h3>Business Segment</h3></td>" & vbCrLf
			response.write "<td width=""20%"">"
			Call Lbx_GetHTML5write("lbxBusSeg", false, 1, 0, _
					oRsBusSeg, "SegmentName", "BusinessSegmentID", cint(sPrevBusSegID), true, "", false)
			response.write "</td></tr>" & vbCrLf	
			
			Response.Write "<tr><td align=left width=""20%""><h3>AMO Category</h3></td>" & vbCrLf
			response.write "<td width=""20%"">"
			Call Lbx_GetHTML5write("lbxAMOCategory", false, 1, 0, _
					oRsAMOCategory, "Name", "FeatureCategoryID", cint(sPrevAMOCAtegoryID), true, "", false)
			response.write "</td>" & vbCrLf	
			Response.Write "<td align=right width='20%'><h3>Updater</h3></td>" & vbCrLf
			response.write "<td>"
			Call Lbx_GetHTML5write("lbxUpdaters", false, 1, 0, _
					oRsUpdaters, "Updater", "Updater", sPrevUpdater, true, "", false)
			response.write "</td>" & vbCrLf
			response.write "</tr>" & vbCrLf
			
			Response.Write "<tr><td align=left width=""20%""><h3>Change Type</h3></td>" & vbCrLf
			response.write "<td width=""20%"">"
			Call Lbx_GetHTML5write("lbxChangeType", false, 1, 0, _
					oRsChangetype, "Description", "StatusID", cint(sPrevChangeType), true, "", false)
			response.write "</td>" & vbCrLf
			Response.Write "<td align=right width=""20%""><h3>Show History Within Last</h3></td>" & vbCrLf
			
			response.write "<td><input type=""text"" name=""txtLastdays"" id=""txtLastdays"" size=""8"" maxlength=""8""  value=" & stimerange & "> Days</td></tr>"
			
			response.write "</tr>" & vbCrLf
			
			
			%>
			<tr>
					<td align=left width="20%"><h3>Update Date:</h3></td>
					<td colspan=2>From&nbsp;<input id="txtUpdateDateFrom" name="txtUpdateDateFrom" value="<%= strUpdateDateFrom %>" class="filter-dateselection" style="HEIGHT: 22px;" size=10 maxLength="10" >
					To&nbsp;<input id="txtUpdateDateTo" name="txtUpdateDateTo" value="<%= strUpdateDateTo %>" class="filter-dateselection" style="HEIGHT: 22px; " size=10 maxLength="10">
					 (MM/DD/YYYY)
					</td></tr>						
			<%
			
			Response.Write "<tr><td align=left width=""20%""><h3>HP PartNo</h3></td>" & vbCrLf
			response.write "<td colspan='2'><input type=""text"" name=""txtHPPartNo"" id=""txtHPPartNo"" size=""20"" maxlength=""20""  value=" & sHPPartNo & "><FONT size=1><EM>Leave blank or enter value with wildcard(*)</EM></FONT></td>"
			
		if oRs.RecordCount > 0 then
			nNumTotalRequest = oRs.RecordCount
			bShowcheckbox = True
		
		end if
		if oRs.RecordCount <= 0 then
			bShowcheckbox = False
		else
			strBuildFilter = ""	
			if IsODM = 1 then	'hide cost items from ODM
				strBuildFilter = "FieldMask <> 32 and FieldMask <> 64 and FieldMask <> 131072"
			end if
			oRs.Filter = strBuildFilter
			nNumRequest = oRs.RecordCount
		end if
		%>
		
		
		<td align=left><INPUT type='button' value='Filter List' id=btnApply name=btnApply LANGUAGE=javascript onclick="return ApplyFilter();"></td></tr>
		<%if nNumRequest = 1000 then %>
			<tr><td colspan='3'><FONT size=1 color='red'><EM>Warning : Limit results to 1000 latest records</EM></FONT></td></tr>
		<% end if %>
	</table>
	
	<TABLE border=0 cellPadding=2 cellSpacing=2 width="100%">
		<colgroup width="50%"><colgroup width="50%">
		<TR><TD colspan=2 align=left><hr size=1 width=100%></TD></TR>
		<%
		if nNumRequest > 0 then
			%>
			<tr>
				<td width="50%">
				<%
				if bShowcheckbox and bUpdate then
					response.write "<INPUT id='btnSelAll' name='btnSelAll' style='width:90' type='button' value='Select All' LANGUAGE='javascript' onclick=""return btnSelectAll_Click()"">&nbsp;"
					response.write "<INPUT id='btnDelSelAll' name='btnDeSelAll' style='width:90' type='button' value='Deselect All' LANGUAGE='javascript' onclick=""return btnDeselectAll_Click()"">"
				end if
				response.write "&nbsp;</td>"
				response.write "<td width=""50%"" align=left>"
				if bUpdate then
					response.write "<INPUT id='btnSave' name='btnSave' type='button' value='Save' LANGUAGE='javascript' onclick=""return BtnSave_Onlick()"">&nbsp;"
				end if 
				response.write "</td>" & vbCrLf
				%>
			</tr>
			<tr>
				<td colspan=2>
				<%
				Response.Write "<table align='center' border=1 WIDTH=100% CELLSPACING=1 CELLPADDING=1>" & vbCRLF
			
				'Row 1; Heading
				Response.Write  "<tr class=tblrow-pulsar>" & vbCRLF
				if bShowcheckbox and bUpdate then
					Response.Write  "<th>Hide From SCM</th>" & vbCrLf
				end if	
				Response.Write  "<th width=100><a onclick='SortField(""ChangeType"", """ & strSortOrder & """)' class='BlueHand'>Change Type"
					call PutArrow( "ChangeType", strSortField, strSortOrder)
					response.write "</a></th>" & vbCrLf
				Response.Write  "<th width=100><a onclick='SortField(""AMOCategory"", """ & strSortOrder & """)' class='BlueHand'>AMO Category"
					call PutArrow( "AMOCategory", strSortField, strSortOrder)
					response.write "</a></th>" & vbCrLf
				Response.Write  "<th width=100><a onclick='SortField(""HPPartNo"", """ & strSortOrder & """)' class='BlueHand'>HP PartNo"
					call PutArrow( "HPPartNo", strSortField, strSortOrder)
					response.write "</a></th>" & vbCrLf
				Response.Write  "<th>Short Description</th>" & vbCrLf
				Response.Write  "<th>Change Description</th>" & vbCrLf
				Response.Write  "<th>Reason<br><span style='font-size:10px;font-style:italic;'>256 max. characters</span></th>" & vbCrLf
				Response.Write  "<th>Updated Date</th>" & vbCrLf
				Response.Write  "<th>Updater</th>" & vbCrLf
				Response.Write  "</tr>" & vbCRLF
			
				'Change History Data
				oRs.MoveFirst
				lngCount = 0

				do while not oRs.EOF
					Response.Write  "<tr align='left' bgcolor=""" & arrColors(lngCount mod 2) & """>" & vbCRLF
			
					'Select check box 
					if bShowcheckbox and bUpdate then
						Response.Write "<td align=""center""><input type='checkbox' NAME='chkHideFromSCM' value=""" & oRs("ChangeHistoryID") & """ "
						if oRs("HideFromSCM") = 1 then 
							response.write "checked"
							sOriginalSelectedIDs = sOriginalSelectedIDs & oRs("ChangeHistoryID").Value & ","
						end if 
						response.write "></td>" & vbCrLf
					end if 

			
					'Change Type
					Response.Write  "<td align=center>" & oRs("ChangeType").Value & "&nbsp;</td>" & vbCRLF

					'AMO Category
					Response.Write  "<td>" & oRs("AMOCategory").Value & "&nbsp;</td>" & vbCRLF

					'Part Number
					Response.Write  "<td align=center>" & oRs("HPPartNo").Value & "&nbsp;</td>" & vbCRLF

					'Short Description
					Response.Write  "<td>" & oRs("ShortDescription").Value & "&nbsp;</td>" & vbCRLF
					
					'Change Description
					Response.Write  "<td>" & server.htmlencode (trim(oRs("ChangeDescription").Value)) & "&nbsp;</td>" & vbCRLF

					'Reason
					Response.Write  "<td align=center><textarea cols=40 rows=3 name=txtReason id=txtReason style='font-family:Arial;	font-size:12px;'    onBlur='changeSum(" & oRs("ChangeHistoryID") & " , this.value);' >" & server.htmlencode (oRs("Reason")) & "</textarea>"
					Response.Write "</td>" & vbCRLF
					'updated date
					Response.Write  "<td>" & oRs("TimeChanged").Value & "&nbsp;</td>" & vbCRLF
					'updater
					Response.Write  "<td>" & oRs("Updater").Value & "&nbsp;</td>" & vbCRLF
			
					lngCount = lngCount + 1
					oRs.MoveNext
				loop
				if len(sOriginalSelectedIDs) > 1 then
					sOriginalSelectedIDs = left(sOriginalSelectedIDs, len(sOriginalSelectedIDs)-1)
				end if 
				response.write "</tr>" & vbCrLf
				Response.Write  "</table>" & vbCRLF
				%>
			</td>
		</tr>
		<%
		else
			Response.Write "<tr><td colspan=2>"
			Response.Write "There was no Change History found.<br>"
			Response.Write "</td></tr>"
		end if 
		%>
	</TABLE>

	<input type="Hidden" name="sortfield" value="<%= strSortField %>">
	<input type="Hidden" name="sortorder" value="<%= strSortOrder %>">
	<input type="Hidden" name="newAVReasson" value="">
	<input type="Hidden" name="newAVReassonID" value="">
	<input type="Hidden" name="OriginalSelectedIDs" value="<%=sOriginalSelectedIDs%>">
    <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />

	<%
	set oRs = nothing
	set oSvr = nothing
end if 'no error


%>

</FORM>
</BODY>
</HTML>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->