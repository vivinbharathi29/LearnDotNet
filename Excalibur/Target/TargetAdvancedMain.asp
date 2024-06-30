<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
' --- GLOBAL & OPTIONAL INCLUDES: ---%>
<!--#INCLUDE FILE="../includes/oConnect.asp"-->

<%
'*************************************************************************************
'* Description	: SQL recordset connection(s) from the root deliverable tables 
'* Creator		: Harris, Valerie
'* Created		: 01/22/2016 - PBI 15594 / Task 15982
'*************************************************************************************
Dim oErrors		'OUR ERROR OBJECT
Dim sErrorMessage

'-------------------------------------------------------------------------------------
'* Purpose		: Return details for the selected Root Deliverable.
'* Inputs		: Product ID, Root ID, Version ID.
'-------------------------------------------------------------------------------------
Dim oRSRootDetails	   'Recordset object.
Sub GetRootDetails(bOpen, ProductID, RootID, VersionID)	  
	On Error Resume Next
	Dim qsRootDetails  'Query string.
	If bOpen=True Then
		Set oRSRootDetails = Server.CreateObject("ADODB.Recordset")
		Set oRSRootDetails.ActiveConnection = oConnect
		Set oErrors = oRSRootDetails.ActiveConnection.Errors
		qsRootDetails = "EXECUTE usp_TargetDeliverable_GetRootDetails "
		qsRootDetails = qsRootDetails & ""& ProductID &", "
		qsRootDetails = qsRootDetails & ""& RootID &", "
		qsRootDetails = qsRootDetails & ""& VersionID &""
		oRSRootDetails.Open qsRootDetails, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
		If oErrors.Count > 0 Then		'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.")
            Response.End()
        End If
	ElseIf bOpen=False Then		        'CLOSE RECORDSET
		If Not (oRSRootDetails Is Nothing) Then
			oRSRootDetails.Close
			Set oRSRootDetails = Nothing
		End If
	End If
End Sub

'--- INSTANTIATE OBJECTS: ---
Call OpenDBConnection(PULSARDB(), True)			'Open database connection, oConnect.

'--- DECLARE LOCAL VARIABLES: ---
Dim iProductID						'INTEGER				
Dim iRootID						    'INTEGER
Dim iVersionID						'INTEGER
Dim sRootLanguage					'STRING
Dim bGlobal					        'BOOLEAN
Dim iRowCount                       'INTEGER
Dim inactiveRowCount                'INTEGER
Dim iColSpan                        'INTEGER
Dim iCount                          'INTEGER
Dim sVersionStatus                  'STRING
Dim sJSON                           'STRING
Dim sTypeId							'INTEGER

'--- DEFINE LOCAL VARIABLES: ---
iProductID = Request("ProductID")
iRootID = Request("RootID")
iVersionID = Request("VersionID")


'---Get Root Deliverable's typeid: --- 
Call GetRootDetails(True, iProductID, iRootID, iVersionID)			
	If Not (oRSRootDetails Is Nothing) Then
		If Not oRSRootDetails.EOF Then 
			sTypeId = oRSRootDetails("typeid")
			If (sTypeId = "1") Then 'if deliverable is HW then set global = true. PBI 205228
				bGlobal = True
			Else
                bGlobal = False
            End If
		Else
			bGlobal = True
		End If
	End If
Call GetRootDetails(False, Empty, Empty, Empty)

Call OpenDBConnection(PULSARDB(), False)	    'Close database connection, oConnect
	%>
<HTML data-browser="" data-version="">
<HEAD>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="0">
<title>Target Advanced Main</title>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<link href="../style/targetadvanced.css" type="text/css" rel="stylesheet">
<STYLE>
    A:visited
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }
</STYLE>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../includes/client/json2.js"></script>
<script type="text/javascript" src="../includes/client/json_parse.js"></script>
<script src="scripts/targetadvancedmain.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    var CountIDforReleases;
function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function ModImages(DelID, ProdID, VerID, Fusion, Pulsar){
	var url;	
	if (Pulsar == 1) {
	    url = 'Pulsar/ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID + '&VersionID=' + VerID;	    
	} else if (Fusion == 1) {
	    url = 'Fusion/ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID + '&VersionID=' + VerID;	   
	} else {
	    url = 'ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID + '&VersionID=' + VerID + '&pulsarplusDivId=Deliverables';
	}

	modalDialog.open({ dialogTitle: 'Change Image', dialogURL: '' + url + '', dialogHeight: 550, dialogWidth: 700, dialogArguments: 'Image', dialogArgumentsName: 'change_image_cell', dialogResizable: true, dialogDraggable: true });

    //save Version ID for results function: ---
	globalVariable.save(VerID, 'version_id');
}

function ModImagesResult(strID) {
        if (typeof (strID) != "undefined") {
            if (strID == "") {
                document.all("Image" + globalVariable.get('version_id')).innerText = "ALL";
            } else {
                document.all("Image" + globalVariable.get('version_id')).innerText = strID;
            }
        }
}


function ModDist(DelID,ProdID,VerID){
	var url;
	
	url = 'ChangeDistribution.asp?ProductID=' + ProdID + '&RootID=' + DelID + '&VersionID=' + VerID;
	modalDialog.open({ dialogTitle: 'Modify Distribution', dialogURL: '' + url + '', dialogHeight: 550, dialogWidth: 650, dialogResizable: true, dialogDraggable: true });

    //save Version ID for results function: ---
	globalVariable.save(VerID, 'version_id');
}

function ModDistResult(strID) {  
        var ResultArray;
        var iVersionID = globalVariable.get('version_id');
        if (typeof (strID) != "undefined") {
            ResultArray = strID.split("|")
            if (ResultArray[0] == "") {
                document.all("Dist" + iVersionID).innerText = "Not Specified";
            } else {
                document.all("Dist" + iVersionID).innerText = ResultArray[0];
            }
            if (ResultArray[1] == "") {
                document.all("Image" + iVersionID).innerText = "ALL";
            } else {
                document.all("Image" + iVersionID).innerText = ResultArray[1];
            }
        }
}

function window_onload() {
    //Instantiate modalDialog load
    modalDialog.load();
}

function release_onclick(CountID, ProdID, DelID, VerID)
{
    CountIDforReleases = CountID;
    var TargetedReleases = document.getElementById("TargetedReleases_" + CountID).value;
    var url = 'TargetReleaseEdit.asp?ProductID=' + ProdID + '&RootID=' + DelID + '&VersionID=' + VerID;
    OpenPopUp(url, 500, 600, "Edit Release", true, false, false, "divReleases", "ifReleases");
}

function CloseTargetReleasePopup(refresh, targetedReleases, targetedReleaseIDs) {
    $("#ifReleases").attr("src", "");
    $("#ifReleases").contents().find("body").html('');
    $("#divReleases").dialog("close");
    $("#divReleases").dialog('destroy');    
}
function SetTargetInfomation(targetedReleases, targetedReleaseIDs) {
    document.getElementById("TargetedReleases_" + CountIDforReleases).value = targetedReleaseIDs;
    if (targetedReleases.length > 0) {
        document.getElementById("aRelease_" + CountIDforReleases).innerHTML = targetedReleases;
        var dd = document.getElementById('cboStatus_' + CountIDforReleases);
        for (var i = 0; i < dd.options.length; i++) {
            if (dd.options[i].text === "Targeted") {
                dd.selectedIndex = i;
                break;
            }
        }
    }       
}

function OpenPopUp(link, newHeight, newWidth, title, noScrollBar, hideCloseButton, Resizable, divID, ifrID) {
    var $divPopup = $('#' + divID);
    $divPopup.dialog({
        height: newHeight,
        width: newWidth,
        modal: true,
        title: title,
        resizable: Resizable,
        draggable: true,
        open: function (event, ui) {
            if (hideCloseButton)
                $(this).parent().children().children('.ui-dialog-titlebar-close').hide();
            else
                $(this).parent().children().children('.ui-dialog-titlebar-close').show();

            if (noScrollBar)
                $divPopup.css('overflow', 'hidden');
        },
        close: function (event, ui) {
            //everytime the jquery dialog is closed trigger this event to clear the iframe so when dialogue is called again it will show blank first then load with the url
            $("#" + ifrID).attr("src", "");
        }
    });

    loadIframe(ifrID, link);
}
function loadIframe(iframeName, url) {
    var $iframe = $('#' + iframeName);
    $iframe.attr("width", "100%");
    $iframe.attr("height", "100%");
    if ($iframe.length) {
        $iframe.attr('src', url);
        return false;
    }
    return true;
}

function selMultiTargetOnchange(object, IsStableConsistent, language, PartNo, PMAlert)
{
    if (language == "Global") {
        if ($('table#tblTarget .select-option[value="Targeted"][option:selected]').length > 1) {
            if (IsStableConsistent == "False") {
                alert("Another Version of this same component is already targeted.  Only one Version of a Component can be Targeted at one time.");
                $("#txtKeepAllTargeted").val("0");
                if (PMAlert == "1")
                    $("#" + object).val('New').prop('selected', true);
                else
                    $("#" + object).val('Available').prop('selected', true);
            }
            else {
                if (!window.confirm("Another Version(s) of this same component is already targeted.  Do you want to keep the previous version(s) Targeted or only Target this Version?  \n\n Click 'OK' to Keep ALL Versions Targeted or Click 'Cancel' to ONLY Target this Version")) {
                    $("#txtKeepAllTargeted").val("0");
                    $('table#tblTarget .select-option[value="Targeted"][option:selected]').each(function (i,obj) {
                        if ($(this).attr("id") != object){
                            if (PMAlert == "1")
                                $(this).val('New').prop('selected', true);
                            else
                                $(this).val('Available').prop('selected', true);
                        }
                    });
                }
                else 
                    $("#txtKeepAllTargeted").val("1");
            }
        }
    }
    else {
        if ($('table#tblTarget .select-option[value="Targeted"][data-PN^="' + PartNo.substring(0, 9) + '"][option:selected]').length > 1) {
            if (IsStableConsistent == "False") {
                alert("Another Version of this same component is already targeted.  Only one Version of a Component can be Targeted at one time.");
                $("#txtKeepAllTargeted").val("0");
                if (PMAlert == "1")
                    $("#" + object).val('New').prop('selected', true);
                else
                    $("#" + object).val('Available').prop('selected', true);
            }
            else {
                if (!window.confirm("Another Version(s) of this same component is already targeted.  Do you want to keep the previous version(s) Targeted or only Target this Version?  \n\n Click 'OK' to Keep ALL Versions Targeted or Click 'Cancel' to ONLY Target this Version")) {
                    $("#txtKeepAllTargeted").val("0");
                    $('table#tblTarget .select-option[value="Targeted"][option:selected]').each(function (i, obj) {
                        if ($(this).attr("id") != object) {
                            if (PMAlert == "1")
                                $(this).val('New').prop('selected', true);
                            else
                                $(this).val('Available').prop('selected', true);
                        }
                    });
                }
                else 
                    $("#txtKeepAllTargeted").val("1");
            }
        }
    }
}
//-->
</SCRIPT>
</HEAD>
<BODY  onload="window_onload();" bgcolor="Ivory">
<%

	dim strDelName
	dim strprodname
	dim cn
	dim rs
	dim blnLoadFailed
	dim strChanges
	dim strVersion
	dim strOTS
	dim strStatus
	dim i
	dim strImages
	dim  strType
    dim isFusion
    dim isPulsar
    dim enableCboStatus
    dim StableConsistent
    isPulsar = 0
	dim strNoOfReleases : strNoOfReleases = 0
	
	
	blnLoadFailed = false	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")

	if not blnLoadFailed then
		rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProdName = ""
			blnLoadFailed = true
            isFusion = 0
            isPulsar = 0
		else
			strprodName = rs("name") & ""
            if rs("Fusion") then
                isFusion = 1
            else
                isFusion = 0
            end if
            if rs("FusionRequirement") then
                isPulsar = 1
            else
                isPulsar = 0
            end if
            StableConsistent = rs("StableConsistent")
		end if
		
		rs.Close
	end if

	rs.Open "spGetDeliverableRootName " & clng(request("RootID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strDelName = ""
		strType = ""
		blnLoadFailed = true
	else
		strDelName = rs("name") & ""
		strType = rs("TypeID") & ""
		strActive = rs("Active") & ""
		if rs("Active") = 0 then
			blnLoadFailed = true
		end if

        If strType <> "1" Then
            iColspan = 13
        Else
            iColspan = 10
        End If
	end if
	
	rs.Close

if blnLoadFailed then
	if trim(strActive) = "0" then
		Response.Write "<BR><BR><font size=2 face=verdana>This deliverable is inactive.  Please remove it from the Requirements tab or contact the developer to have it reactivated.</font>"
	else
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to load advanced targeting information</font>"
	end if
else

	rs.Open "spListVersions4Targeting " & clng(request("ProductID")) & "," & clng(request("RootID")),cn,adOpenForwardOnly

	if rs.EOF and rs.BOF then
		Response.Write "<BR><BR><font size=2 face=verdana>No Versions for selected Deliverable on Selected Product.</font>"
	else
%>



<h3>Target Versions</h3>
<h4><%=strDelName & " for " & strProdName%></h4>
<form ID=frmTarget action="TargetAdvancedSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method="post">
<div id="page" class="show">
<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor=tan style="border: solid 1px tan" id="tblTarget">
	<TR id="trApplyTargetSection" class="hide">
        <td colspan="<%=iColSpan%>" style="background-color:#CCCCCC !important;" align="right">
            <table width="720px">
                <tr>
                    <td width="160px" class="right font-small">Apply Target to Selected</td>
                    <td width="100px" class="left">
                        <SELECT id="selMultiTarget" class="small-input">
				            <OPTION value="" selected></OPTION>
                            <OPTION value="Available">Available</OPTION>		
				            <OPTION value="Targeted">Targeted</OPTION>				
			            </SELECT>
                    </td>
                    <%if isPulsar = 0 then%>
                        <td width="80px" class="right font-small">Target Notes</td>
                        <td width="250px" class="left">
                            <input type="text" id="txtMultiNotes" value="" maxlength="255" class="small-input" /> 
                        </td>
                    <% else%>
                        <td width="80px" style="display:none" class="right font-small">Target Notes</td>
                        <td width="250px" style="display:none" class="left">
                            <input type="text" id="txtMultiNotes" value="" maxlength="255" class="small-input" /> 
                        </td>
                    <%end if%>
                    <td width="130px">
                        <input type="button" id="btnApply" value="Apply" class="button"/>&nbsp;
                        <input type="button"  id="btnUndo" value="Undo" class="button"/>&nbsp;
                        <img src="../images/info.png" title="Change Target for all of the selected versions." class="help"/>
                    </td>
                </TR>
            </table>
        </td>
	</TR>
	<TR>
        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold; text-align:center;" class="td-select"><input type="checkbox" id="chkSelectAll" /></td>
		<%if trim(strType) <> "1" then%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Part&nbsp;Number</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">ID</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Version</td>
        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Language</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target</td>
        <%if isPulsar = 1 then%>
            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Release</td>
        <%end if%>
        <%if isPulsar = 0 then%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target&nbsp;Notes</td>
        <%end if%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Distribution</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Images</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Developer<BR>Approved</td>
		<td style="display:none;font-size:xx-small; font-family:Verdana; font-weight:bold;">OEM&nbsp;Ready</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">WHQL</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Workflow</td>
		<%else%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Part&nbsp;Number</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">ID</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Version</td>
        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Language</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target</td>
        <%if isPulsar = 1 then%>
            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Release</td>
        <%end if%>
		<%if isPulsar = 0 then%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target&nbsp;Notes</td>
        <%end if%>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Model</td>
		<td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Developer&nbsp;Approved</td>
		<%end if%>
	</TR>
	<%
		i=0
        iRowCount = 0
        inactiveRowCount = 0
		dim strDist
		dim strWhqlStatus, strWhqlColor
		dim strOemReadyStatus, strOemReadyColor
		
		do while not rs.EOF	
			strDist = ""
			if rs("Preinstall") then
				strDist = strDist & ",Preinstall"
			end if
			if rs("Preload") then
				strDist = strDist & ",Preload"
			end if
			if rs("DropInBox") then
				strDist = strDist & ",DIB"
			end if
			if rs("Web") then
				strDist = strDist & ",Web"
			end if
			if rs("SelectiveRestore") then
				strDist = strDist & ",SelectiveRestore"
			end if
			if rs("ARCD") then
				strDist = strDist & ",DRCD"
			end if
			if rs("DRDVD") then
				strDist = strDist & ",DRDVD"
			end if
			if rs("RACD_Americas") then
				strDist = strDist & ",RACD-Americas"
			end if
			if rs("RACD_EMEA") then
				strDist = strDist & ",RACD-EMEA"
			end if
			if rs("RACD_APD") then
				strDist = strDist & ",RACD-APD"
			end if
			if rs("OSCD") then
				strDist = strDist & ",OSCD"
			end if
			if rs("DocCD") then
				strDist = strDist & ",DocCD"
			end if
			
			if rs("Patch") then
				strDist = strDist & ",Patch" 
			end if

            if rs("RCDOnly") then
				strDist = strDist & ",RCDOnly" 
			end if

			if strDist <> "" then
				strDist = mid(strDist,2)
			else
				strDist = "none"
			end if
			
			strOTS = ""
			if trim(strType) = "1" then			
				strVersion = rs("Vendor") & "&nbsp;" & rs("Version") & ""
			else
				strVersion = rs("Version") & ""
			end if
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if

			strImages = rs("ImageSummary") & ""
			if strImages = "" then
				strImages = "ALL"
			end if
			
			if rs("DevStatus") = 2 then
				strDevStatus = "No"
				strDevBGColor = "mistyrose"'"Red"
			elseif rs("DevStatus") = 1 then
				strDevStatus = "Yes"
				strDevBGColor = "DarkSeaGreen"'"SpringGreen"
			elseif isnull(rs("DevStatus")) then
				strDevStatus = "Not&nbsp;Supported"
				strDevBGColor = "ffff99"'"Yellow"
			else
				strDevStatus = "TBD"
				strDevBGColor = "ffff99"'"Yellow"
			end if	
			
			if rs("OEMReadyRequired") then
			    if rs("OEMReadyStatus") & "" <> "" Then
			        Select Case rs("OEMReadyStatus") & ""
			            Case 0
			                strOEMReadyStatus = "Required"
			                strOEMReadyColor = "mistyrose"'"Red"
			            Case 1
			                strOEMReadyStatus = "Submitted"
			                strOEMReadyColor = "ffff99"'"Yellow"
			            Case 2
			                strOEMReadyStatus = "Approved"
			                strOEMReadyColor = "DarkSeaGreen"'"SpringGreen"
			            Case 3
			                strOEMReadyStatus = "Failed"
			                strOEMReadyColor = "mistyrose"'"Red"
			            Case 4
			                strOEMReadyStatus = "Exempt"
			                strOEMReadyColor = "DarkSeaGreen"'"SpringGreen"
			        End Select
			    else
			        strOEMReadyStatus = "Required"
                    strOEMReadyColor = "mistyrose"'"Red"
			    end if
			else
			    strOEMReadyStatus = "Not Required"
			    strOEMReadyColor = "DarkSeaGreen"'"SpringGreen"    
			end if
			
			if rs("WhqlRequired") then
			    If rs("WhqlStatus") & "" <> "" Then
			        Select Case rs("WhqlStatus") & ""
			            Case 0
			                strWhqlStatus = "Required"
			                strWhqlColor = "mistyrose"
			            Case 1
		    	            strWhqlStatus = "Submitted"
	    		            strWhqlColor = "ffff99"'"Yellow"
    			        Case 2
			                strWhqlStatus = "Approved"
			                strWhqlColor = "DarkSeaGreen"'"SpringGreen"
			            Case 3
		    	            strWhqlStatus = "Failed"
	    		            strWhqlColor = "mistyrose"
    			        Case 4
			                strWhqlStatus = "Waiver"
			                strWhqlColor = "DarkSeaGreen"'"SpringGreen"
			        End Select
			    Else
			        strWhqlStatus = "Required"
                    strWhqlColor = "mistyrose"
                end if
			else
			    strWhqlStatus = "Not Required"
			    strWhqlColor = "DarkSeaGreen"'"SpringGreen"
			end if
			
            strWorkflowStatus = replace(rs("Location") & "","Workflow Complete","Complete")
            if instr(strWorkflowStatus,"(") > 0 then
                strWorkflowStatus = left(strWorkflowStatus,instr(strWorkflowStatus,"(")-1)
            end if

            if (trim(strWorkflowStatus) = "Development" or trim(strWorkflowStatus) = "Functional Test") and request("ExcludeFunComp") then
				strWorkflowColor = "mistyrose"
				enableCboStatus = "disabled"
            elseif strWorkflowStatus <> "Complete" then
                strWorkflowColor = "mistyrose"
				enableCboStatus = ""
            else
                strWorkflowColor = "DarkSeaGreen"'"SpringGreen"
				enableCboStatus = ""
            end if

            If rs("Active") Then 
                If i = 0 Then
                    iRowCount = 0
                Else
                    iRowCount = iRowCount + 1
                End If
                
                iCount = iRowCount
                sVersionStatus = "active"
            Else
                iCount = -1
                inactiveRowCount = inactiveRowCount + 1
                sVersionStatus = "inactive"
            End If
        iCount = iCount + inactiveRowCount
        
	%>
	<%if false then 'trim(strType) = "1" and (not rs("Active")) then%> 
		<TR bgcolor=Gainsboro>
			<TD width=20 valign=top><font size=1 face=verdana><a target=_blank href="../Query/DeliverableVersionDetails.asp?ID=<%=rs("VersionID")%>"><%=rs("VersionID")%></a></font></TD>
			<TD width=100 valign=top><font size=1 face=verdana><%=strVersion%></font></TD>
			<TD valign=top colspan="5">This version is inactive.  Contact the developer for assistance.</TD>
		</TR>
	<%else%>
		<%if not rs("Active") then%>
        <TR bgcolor="gainsboro">
		<%else%>
        	<TR>
        <%end if%>
        <%if rs("Active") then%>
            <TD align="center" valign="top" class="td-select"><input type="checkbox" class="chkbox-option" data-checkbox="<%=iCount%>" data-status="<%=sVersionStatus %>" data-language="<%=rs("Language")%>" data-pn="<%=rs("PartNumber")%>" data-pmalert="<%=rs("PMAlert")%>" /></TD>
        <%else%>
            <TD align="center" valign="top">&nbsp;</TD>        
        <%end if %>
        <TD valign=top><font size=1 face=verdana><%=rs("PartNumber")%>&nbsp;</font></TD>		
        <TD width=20 valign=top><font size=1 face=verdana><a id="aVersionID_<%=iCount%>" target=_blank href="../Query/DeliverableVersionDetails.asp?ID=<%=rs("VersionID")%>"><%=rs("VersionID")%></a></font>
		<TD width=100 valign=top><font size=1 face=verdana><%=strVersion%></font>
		<%if not rs("Active") then%>
			 <font size=1 color=red face=verdana>(inactive)</font>
		<%end if%>
		</TD>
        <TD valign=top width=90>
            <font size=1 face=verdana><%=rs("Language")%></font>
        </TD>
		<TD valign=top width=90>
			<% if isnull(rs("ProdDelID")) then%>
				<INPUT type="hidden" id=txtID name=txtID value="0">
			<% else%>
				<INPUT type="hidden" id=txtID name=txtID value="<%=rs("ProdDelID")%>">
			<%end if%>
			<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=rs("VersionID")%>">
			<SELECT id="cboStatus_<%=iCount%>" name=cboStatus style="Width:85" class="select-option <%=enableCboStatus%>" data-select="<%=iCount%>" data-status="<%=sVersionStatus %>" data-pmalert="<%=rs("PMAlert")%>" data-pn="<%=rs("PartNumber")%>" onchange="selMultiTargetOnchange('cboStatus_<%=iCount%>','<%=StableConsistent%>','<%=rs("Language")%>','<%=rs("PartNumber")%>','<%=rs("PMAlert")%>');" >
				<OPTION selected>Available</OPTION>	
				<%strStatus="Available"%>
				<%if rs("PMAlert") = 1 then
					strStatus = "New"%>
					<OPTION selected>New</OPTION>	
				<%end if%>
				<%if rs("targeted") = true then
					strStatus = "Targeted"%>
					<OPTION value="Targeted" selected>Targeted</OPTION>	
				<%else%>
					<OPTION value="Targeted">Targeted</OPTION>	
				<%end if%>
			</SELECT>
			<INPUT type="hidden" id=cboStatusTag name=cboStatusTag value="<%=strStatus%>" data-dbtarget="<%=iCount%>">
            <input type="hidden" id="select_<%=iCount%>" value="<%=strStatus %>" />
		</TD>
        <%if isPulsar = 1 then%>
        <td>
            <%if strStatus = "Targeted" then%>
                <%if not rs("Active") then%>
                    <span><%=rs("Release")%></span>
                <%else%>
                    <a id="aRelease_<%=iCount%>" onclick="release_onclick(<%=iCount%>,<%=request("ProductID")%>, <%=request("RootID")%>, <%=rs("VersionID")%>);" href="#"><%=rs("Release")%></a><input style="display:none;" type="text" value="<%=rs("TargetedReleases")%>" name="TargetedReleases_<%=iCount%>" id="TargetedReleases_<%=iCount%>"><input style="display:none;" type="text" value="<%=rs("TargetedReleases")%>" name="TargetedReleasesOri_<%=iCount%>" id="TargetedReleasesOri_<%=iCount%>">            
                <%end if%>
            <%else%>
                <%if not rs("Active") then%>
                    <span>Select Release</span>
                <%else%>
                    <%if rs("NoOfReleases") = "1" then
                        strNoOfReleases = 1 
                        if enableCboStatus = "disabled" then%>
                            <span><%=rs("Release")%></span>
                        <%else%>
                            <a id="aRelease_<%=iCount%>" onclick="release_onclick(<%=iCount%>,<%=request("ProductID")%>, <%=request("RootID")%>, <%=rs("VersionID")%>);"  href="#"><%=rs("Release")%></a><input style="display:none;" type="text" value="<%=rs("ReleaseID")%>" name="TargetedReleases_<%=iCount%>" id="TargetedReleases_<%=iCount%>"><input style="display:none;" type="text" value="<%=rs("TargetedReleases")%>" name="TargetedReleasesOri_<%=iCount%>" id="TargetedReleasesOri_<%=iCount%>">
                        <%end if%>
                    <%else
                        if enableCboStatus = "disabled" then%>
                            <span>Select Release</span>
                        <%else%>
                            <a id="aRelease_<%=iCount%> <%=enableCboStatus%>" onclick="release_onclick(<%=iCount%>,<%=request("ProductID")%>, <%=request("RootID")%>, <%=rs("VersionID")%>);"   href="#">Select Release</a><input style="display:none;" type="text" value="<%=rs("TargetedReleases")%>" name="TargetedReleases_<%=iCount%>" id="TargetedReleases_<%=iCount%>"><input style="display:none;" type="text" value="<%=rs("TargetedReleases")%>" name="TargetedReleasesOri_<%=iCount%>" id="TargetedReleasesOri_<%=iCount%>">  
                        <%end if%>
                    <%end if%>
                <%end if%>
            <%end if%>
        </td>
        <%end if%>
        <%if isPulsar = 0 then%>
		<TD valign="top">
            <INPUT type="text" id="txtNotes<%=trim(i)%>" name="txtNotes<%=trim(i)%>" style="WIDTH:100%" value="<%=rs("TargetNotes") & ""%>" maxlength="255" class="text-option" data-text="<%=iCount%>" data-status="<%=sVersionStatus %>">
			<INPUT type="hidden" id="txtNotesTag<%=trim(i)%>" name="txtNotesTag<%=trim(i)%>" value="<%=rs("TargetNotes")& ""%>" data-dbnote="<%=iCount%>">
		</TD>
        <%else%>
         <TD style="display:none" valign="top">
            <INPUT type="text" id="txtNotes<%=trim(i)%>" name="txtNotes<%=trim(i)%>" style="WIDTH:100%" value="<%=rs("TargetNotes") & ""%>" maxlength="255" class="text-option" data-text="<%=iCount%>" data-status="<%=sVersionStatus %>">
			<INPUT type="hidden" id="txtNotesTag<%=trim(i)%>" name="txtNotesTag<%=trim(i)%>" value="<%=rs("TargetNotes")& ""%>" data-dbnote="<%=iCount%>">
		  </TD>
        <%end if%>
        <%if trim(strType) = "1" then%>
			<TD valign=top nowrap><font size=1 face=verdana><%=rs("ModelNumber")%>&nbsp;</font></TD>
		<%else%>
			<TD valign="top"><a href="javascript: ModDist(<%=request("RootID")%>, <%=request("ProductID")%>, <%=rs("VersionID")%>);"><font size=1 face=verdana ID="Dist<%=trim(rs("VersionID"))%>"><%=strDist%></font></a></TD>
			<TD valign="top"><a href="javascript: ModImages(<%=request("RootID")%>, <%=request("ProductID")%>, <%=rs("VersionID")%>,<%=isFusion%>,<%=isPulsar%>);"><font size=1 face=verdana ID="Image<%=trim(rs("VersionID"))%>"><%=strImages%></font></a></TD>
		<%end if%>
		<TD valign=top bgcolor=<%=strDevBGColor%>><font size=1 face=verdana><%=strDevStatus%></font></TD>
		<%if trim(strType) <> "1" then%>
		<td style="display:none;vertical-align:top; white-space:nowrap; font-family:verdana; font-size:xx-small; background-color:<%=strOemReadyColor %>"><%=strOemReadyStatus %></td>
		<td style="vertical-align:top; white-space:nowrap; font-family:verdana; font-size:xx-small; background-color:<%=strWhqlColor %>"><%=strWhqlStatus %></td>
        <td style="vertical-align:top; white-space:nowrap; font-family:verdana; font-size:xx-small;background-color:<%=strWorkflowColor %>"><%=strWorkflowStatus %></td>
		<%end if%>
	</TR>
	<%end if%>
	<%
			i=i+1
			rs.MoveNext
		loop
		rs.Close
	%>
</table>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=clng(request("ProductID"))%>">
<input type="hidden" id="inpGlobal" value="<%=bGlobal %>" />
<input type="hidden" id="inpCount" value="<%=iRowCount %>" />
<input type="hidden" id="inactiveCount" value="<%=inactiveRowCount %>" />
<input type="hidden" id="inpMultiTarget" value=""/>
<input type="hidden" id="inpMultiUsed" value="False"/>
<INPUT type="hidden" id=IsPulsarProduct name=IsPulsarProduct value="<%=isPulsar%>">
<input type="hidden" id="txtTargetReleases" value="" />   
<input type="hidden" id="txtTargetReleaseIDs" value="" /> 
<input type="hidden" id="txtNoOfReleases" name="txtNoOfReleases" value="<%=strNoOfReleases%>" />
<input type="hidden" id="txtStableConsistent" name="txtStableConsistent" value="<%=StableConsistent%>" />
<input type="hidden" id="txtKeepAllTargeted" name="txtKeepAllTargeted" value="1" />
</div>
<div id="loading-dialog">
    <div><p id="msg-page-load" class="font-small"></p>&nbsp;&nbsp;<img src="../images/Loading24.gif" alt="Processing update, please wait." /></div>
</div>
<div id="divReleases" title="Coolbeans" style="display: none;">
        <iframe frameborder="0" name="ifReleases" id="ifReleases" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
</form>


<%
	end if
end if
set rs= nothing
set cn=nothing
%>

</BODY>
</HTML>