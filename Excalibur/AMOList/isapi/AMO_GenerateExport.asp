<%@  Language=VBScript %>
<%
  OPTION EXPLICIT 
'  Response.Buffer = true 
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
'printrequest
call validateSession
call SetPermission
dim sHeader, sHelpfile, sErr, sPALName
dim sUserName
dim sUpdater, bHideOwners
dim sA, sOverride_COM_CD, sBUS_DEF_FIELD4, sCTRY_CD, sCURR_CD, sDIFF_CD, sENTRY_SOURCE_CD
dim sM, sMKT, sMKT_CD, sPA_DISC_FLG, sPRC_DISP_CD, sPRC_TERM_CD
dim sPROD, sPROD_DISP_EXCL_CD, sQBL_SEQ_NBR, sSERIAL_FLG, sSERV_CD, sSRT_CD, sSUI
dim sUOM_CD, sDESC_CD_Common, sDESC_CD_Quote, sOld_Eff_DT
dim sHW_PROD_FAMILY, sSW_PROD_FAMILY, sHW_TAX_CLASS_CD, sSW_TAX_CLASS_CD
dim sParentPin, sRowLevel1, sRowLevel2, sODMCode, sFormatCode, sBrandName, sCountrification, sLCStatus, sLCStatusLocal
dim sMLORW, sSDFFlag, sLevel, sSDFFlagLocal, sLCStatusBaseWithLocal, sSDFFlagBaseWithLocal
dim sBusSegIDs, sOwnerIDs, sBusSegHTML, sOwnersHTML
dim nSelProdLine, nSelSuppCodeGBU, nSelInitialSuppCode, nSelPALUser, nMode, nUserID, nCount
dim oSvr, oSvrPAL, oErr, oRs, oRsBusSeg, oRsGroupOwner
dim lbxProdLineHTML, lbxSuppCodeGBUHTML, lbxInitialSuppCodeHTML

'Hide Owned By in Pulsar, not used
bHideOwners = True

sErr = ""

'different modes:
'2 = save DFT
'3 = view/edit DFT
'5 = save Autoload
'6 = view/edit Autoload
'14 = save DFT Description
'15 = view/edit DFT Description
'17 = save Select Availability (SA) (Blind Date)
'18 = view/edit Select Availability (SA) (Blind Date)
if Request.QueryString("Mode") = "" then
	nMode = 3 'view DFT
else
	nMode = cint(Request.QueryString("Mode"))
end if

if IsODM = 1 and cint(nMode) = cint(3) then
	nMode = 18 'view GPSy DFT Select Availability (SA)
end if

sUpdater = session("FullName")
nUserID = Session("AMOUserID")

if Request.Form("chkBusSeg") = "" then
	'get the cookie from the RAS Review filter.
	sBusSegIDs = GetDBCookie("AMO chkBusSeg")
else
	sBusSegIDs = replace(Request.Form("chkBusSeg"), " ", "")
	'save the cookie for the remaining pages to use
	Call SaveDBCookie("AMO Export_chkBusSeg", sBusSegIDs)
end if

if Request.Form("chkGroupOwner") = "" then
	'get the cookie from the RAS Review filter.
	'sOwnerIDs = GetDBCookie( "AMO chkGroupOwner")
    sOwnerIDs  = 0
else
	'sOwnerIDs = replace(Request.Form("chkGroupOwner"), " ", "")
    sOwnerIDs  = 0
	'save the cookie for the remaining pages to use
	Call SaveDBCookie( "AMO Export_chkGroupOwner", sOwnerIDs )
end if

'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO

if sErr = "" then
	select case nMode
		case 2 
			'Save DFT information
			'get data entered then send to server
			sErr = oSvr.AMO_SaveDFT(Application("REPOSITORY"), _
					sBusSegIDs, _
					sOwnerIDs, _
					nUserID, _
					Request.Form("txtUserName"), _
					Request.Form("lbxProdLine"), _
					Request.Form("lbxSuppCode"), _
					Request.Form("lbxInitialSuppCode"), _
					Request.Form("txtHW_PROD_FAMILY"), _
					Request.Form("txtSW_PROD_FAMILY"), _
					Request.Form("txtHW_TAX_CLASS_CD"), _
					Request.Form("txtSW_TAX_CLASS_CD"), _
					Request.Form("txtOverride_COM_CD"), _
					Request.Form("txtA"), _
					Request.Form("txtBUS_DEF_FIELD4"), _
					Request.Form("txtCTRY_CD"), _
					Request.Form("txtCURR_CD"), _
					Request.Form("txtDIFF_CD"), _
					Request.Form("txtENTRY_SOURCE_CD"), _
					Request.Form("txtM"), _
					Request.Form("txtMKT"), _
					Request.Form("txtMKT_CD"), _
					Request.Form("txtPA_DISC_FLG"), _
					Request.Form("txtPRC_DISP_CD"), _
					Request.Form("txtPRC_TERM_CD"), _
					Request.Form("txtPROD"), _
					Request.Form("txtPROD_DISP_EXCL_CD"), _
					Request.Form("txtQBL_SEQ_NBR"), _
					Request.Form("txtSERIAL_FLG"), _
					Request.Form("txtSERV_CD"), _
					Request.Form("txtSRT_CD"), _
					Request.Form("txtSUI"), _
					Request.Form("txtUOM_CD") )
					
			if sErr <> "True" then
				sErr = "AMO Save DFT Error, AMO_GenerateExport.asp"
		        Response.Write(sErr)
                Response.End()
			else
				Response.Redirect "AMO_DFTFiles.asp"
			end if

		case 5
			'Save Autoload information
			'get data entered then send to server
			sErr = oSvr.AMO_SaveAutoload(Application("REPOSITORY"), _
					sBusSegIDs, _
					sOwnerIDs, _
					nUserID, _
					Request.Form("txtParentPin"), _
					Request.Form("txtRowLevel1"), _
					Request.Form("txtRowLevel2"), _
					Request.Form("txtODMCode"), _
					Request.Form("txtFormatCode"), _
					Request.Form("txtBrandName"), _
					Request.Form("txtCountrification"), _
					Request.Form("txtLCStatus"), _
					Request.Form("txtMLORW"), _
					Request.Form("txtSDFFlag"), _
					Request.Form("txtLevel"), _
					Request.Form("txtLCStatusLocal"), _
					Request.Form("txtSDFFlagLocal"), _
					Request.Form("txtLCStatusBaseWithLocal"), _
					Request.Form("txtSDFFlagBaseWithLocal") )
					
			if sErr <> "True" then
				sErr = "AMO Save DFT Error, AMO_GenerateExport.asp"
		        Response.Write(sErr)
                Response.End()
			else
				Response.Redirect "AMO_AutoloadFiles.asp"
			end if

		case 14
			'Save DFT Description information
			'get data entered then send to server
			sErr = oSvr.AMO_SaveDFTDescription(Application("REPOSITORY"), _
					sBusSegIDs, _
					sOwnerIDs, _
					nUserID, _
					Request.Form("txtUserName"), _
					Request.Form("txtDESC_CD_Common"), _
					Request.Form("txtDESC_CD_Quote"), _
					Request.Form("txtM") )
			if sErr <> "True" then
				sErr = "AMO Save DFT Error, AMO_GenerateExport.asp"
		        Response.Write(sErr)
                Response.End()
			else
				Response.Redirect "AMO_DFTDescriptionFile.asp"
			end if

		case 17
			'Save Select Availability (SA) information
			'get data entered then send to server
			 sErr = oSvr.AMO_SaveBlindDate(Application("REPOSITORY"), _
					sBusSegIDs, _
					sOwnerIDs, _
					nUserID, _
					Request.Form("txtUserName"), _
					Request.Form("txtOld_Eff_DT"), _
					Request.Form("txtM") )
					
			if sErr <> "True" then
				sErr = "AMO Save DFT Error, AMO_GenerateExport.asp"
		        Response.Write(sErr)
                Response.End()
			else
				Response.Redirect "AMO_BlindDateFile.asp"
			end if

	end select
end if

if sErr = "" then
	'Get data for the type
	select case nMode
		case 3
			'DFT
			set oRs = oSvr.AMO_GenerateDFT(Application("REPOSITORY"), nUserID)
		case 6
			'autoload
			set oRs = oSvr.AMO_GenerateAutoload(Application("REPOSITORY"), nUserID)
		case 15
			'DFT Description
			set oRs = oSvr.AMO_GenerateDFTDescription(Application("REPOSITORY"), nUserID)
		case 18
			'Blind Date
			set oRs = oSvr.AMO_GenerateBlindDate(Application("REPOSITORY"), nUserID)
	end select
end if

if sErr = "" then
	select case nMode
		case 3, 1
			'DFT
			if (oRs.EOF) then
				'first time to use, set default values
				sUserName= lcase(Session("LoggedInUser"))
				nSelProdLine = 0
				nSelSuppCodeGBU = 0
				nSelInitialSuppCode = 0
				sHW_PROD_FAMILY = ""
				sSW_PROD_FAMILY = ""
				sHW_TAX_CLASS_CD = ""
				sSW_TAX_CLASS_CD = ""
				sOverride_COM_CD = ""
				sA= ""
				sBUS_DEF_FIELD4 = ""
				sCTRY_CD = ""
				sCURR_CD = ""
				sDIFF_CD = ""
				sENTRY_SOURCE_CD = ""
				sM = ""
				sMKT = ""
				sMKT_CD = ""
				sPA_DISC_FLG = ""
				sPRC_DISP_CD = ""
				sPRC_TERM_CD = ""
				sPROD = ""
				sPROD_DISP_EXCL_CD = ""
				sQBL_SEQ_NBR = ""
				sSERIAL_FLG = ""
				sSERV_CD = ""
				sSRT_CD = ""
				sSUI = ""
				sUOM_CD = ""
			else
				'set the values from recordset
				sUserName = oRs.Fields("UserName").Value
				nSelProdLine = oRs.Fields("ProductLineID").Value
				nSelSuppCodeGBU = oRs.Fields("Supplier_Code_DivisionID").Value
				nSelInitialSuppCode = oRs.Fields("Supplier_Code").Value
				sHW_PROD_FAMILY = oRs.Fields("HW_PROD_FAMILY").Value
				sSW_PROD_FAMILY = oRs.Fields("SW_PROD_FAMILY").Value
				sHW_TAX_CLASS_CD = oRs.Fields("HW_TAX_CLASS_CD").Value
				sSW_TAX_CLASS_CD = oRs.Fields("SW_TAX_CLASS_CD").Value
				sOverride_COM_CD = oRs.Fields("Override_COM_Code").Value
				sA= oRs.Fields("A").Value
				sBUS_DEF_FIELD4 = oRs.Fields("BUS_DEF_FIELD4").Value
				sCTRY_CD = oRs.Fields("CTRY_CD").Value
				sCURR_CD = oRs.Fields("CURR_CD").Value
				sDIFF_CD = oRs.Fields("DIFF_CD").Value
				sENTRY_SOURCE_CD = oRs.Fields("ENTRY_SOURCE_CD").Value
				sM = oRs.Fields("M").Value
				sMKT = oRs.Fields("MKT").Value
				sMKT_CD = oRs.Fields("MKT_CD").Value
				sPA_DISC_FLG = oRs.Fields("PA_DISC_FLG").Value
				sPRC_DISP_CD = oRs.Fields("PRC_DISP_CD").Value
				sPRC_TERM_CD = oRs.Fields("PRC_TERM_CD").Value
				sPROD = oRs.Fields("PROD").Value
				sPROD_DISP_EXCL_CD = oRs.Fields("PROD_DISP_EXCL_CD").Value
				sQBL_SEQ_NBR = oRs.Fields("QBL_SEQ_NBR").Value
				sSERIAL_FLG = oRs.Fields("SERIAL_FLG").Value
				sSERV_CD = oRs.Fields("SERV_CD").Value
				sSRT_CD = oRs.Fields("SRT_CD").Value
				sSUI = oRs.Fields("SUI").Value
				sUOM_CD = oRs.Fields("UOM_CD").Value
			end if

		case 6, 4
			'Autoload
			if (oRs.EOF) then
				'first time to use, set default values
				sParentPin = ""
				sRowLevel1 = ""
				sRowLevel2 = ""
				sODMCode = ""
				sFormatCode = ""
				sBrandName = ""
				sCountrification = ""
				sLCStatus = ""
				sMLORW = ""
				sSDFFlag = ""
				sLevel = ""
				sLCStatusLocal = ""
				sSDFFlagLocal = ""
				sLCStatusBaseWithLocal = ""
				sSDFFlagBaseWithLocal = ""
			else
				'set the values from recordset
				sParentPin = oRs.Fields("ParentPin").Value
				sRowLevel1 = oRs.Fields("RowLevel1").Value
				sRowLevel2 = oRs.Fields("RowLevel2").Value
				sODMCode = oRs.Fields("ODMCode").Value
				sFormatCode = oRs.Fields("FormatCode").Value
				sBrandName = oRs.Fields("BrandName").Value
				sCountrification = oRs.Fields("Countrification").Value
				sLCStatus = oRs.Fields("LCStatus").Value
				sMLORW = oRs.Fields("MLORW").Value
				sSDFFlag = oRs.Fields("SDFFlag").Value
				sLevel = oRs.Fields("Level").Value
				sLCStatusLocal = oRs.Fields("LCStatusLocal").Value
				sSDFFlagLocal = oRs.Fields("SDFFlagLocal").Value
				sLCStatusBaseWithLocal = oRs.Fields("LCStatusBaseWithLocal").Value
				sSDFFlagBaseWithLocal = oRs.Fields("SDFFlagBaseWithLocal").Value
			end if
			
		case 15, 13
			'DFT Description
			if (oRs.EOF) then
				'first time to use, set default values
				sUserName= ""
				sDESC_CD_Common = ""
				sDESC_CD_Quote = ""
				sM = ""
			else
				'set the values from recordset
				sUserName = oRs.Fields("UserName").Value
				sDESC_CD_Common = oRs.Fields("DESC_CD_Common").Value
				sDESC_CD_Quote = oRs.Fields("DESC_CD_Quote").Value
				sM = oRs.Fields("M").Value
			end if

		case 18, 16
			'Select Availability (SA) (Blind Date)
			if (oRs.EOF) then
				'first time to use, set default values
				sUserName= ""
				sOld_Eff_DT = ""
				sM = ""
			else
				'set the values from recordset
				sUserName = oRs.Fields("UserName").Value
				sOld_Eff_DT = oRs.Fields("Old_Eff_DT").Value
				sM = oRs.Fields("M").Value
			end if

	end select
end if

if sErr = "" and (nMode = 1 or nMode = 3) then
	'use function in the \include\listboxRS.inc to populate DD
	'set oSvrPAL = server.CreateObject("JF_S_PAL.IsAVL")
    set oSvrPAL = New ISAVL
	set oRs = oSvrPAL.AVL_ProductLine(Application("REPOSITORY"))
	if oRs is nothing then
		sErr = "Generate Export Error, AMO_GenerateExport.asp"
		Response.Write(strError)
        Response.End()
	end if
end if

if sErr = "" and (nMode = 1 or nMode = 3) then
	lbxProdLineHTML = Lbx_GetHTML5("lbxProdLine", false, 1, 0, _
							oRs, "Value", "CategoryID", nSelProdLine, false, "", false)
    set oSvrPAL = New ISAVL
	set oRs = oSvrPAL.AVL_SupplierCodeGBU(Application("REPOSITORY"))
	if oRs is nothing then
		sErr = "Generate Export Error, AMO_GenerateExport.asp"
		Response.Write(strError)
        Response.End()
	end if
end if

if sErr = "" and (nMode = 1 or nMode = 3) then
	lbxSuppCodeGBUHTML = Lbx_GetHTML5("lbxSuppCode", false, 1, 0, _
						oRs, "Description", "BusinessSegmentId", nSelSuppCodeGBU, false, "", false)
	'don't use this LBox if combine in Supp Code Lbox
    set oSvrPAL = New ISAVL
	set oRs = oSvrPAL.AVL_InitialSupplierCode(Application("REPOSITORY"))
	if oRs is nothing then
		sErr = "Generate Export Error, AMO_GenerateExport.asp"
		Response.Write(strError)
        Response.End()
	end if
end if

if sErr = "" and (nMode = 1 or nMode = 3) then
	lbxInitialSuppCodeHTML = Lbx_GetHTML5("lbxInitialSuppCode", false, 1, 0, _
	    oRs, "Code", "Code", nSelInitialSuppCode, false, "", false)
end if

if sErr = "" then
	'Business segments
	'set oErr = GetMOLCategory(oRsBusSeg, 28)
	set oRsBusSeg = GetMOLCategory(34)	
	if oRsBusSeg is Nothing then
		Response.Write("Recordset error: oRsBusSeg")
		Response.End()
	else
		if not oRsBusSeg is nothing then
			if oRsBusSeg.RecordCount > 0 then
				oRsBusSeg.Sort = "SegmentName ASC"
				if sBusSegIDs = "" then
					'make every one checked
					oRsBusSeg.MoveFirst
					do until oRsBusSeg.EOF
						sBusSegIDs = sBusSegIDs & oRsBusSeg("BusinessSegmentID").Value & ","
						oRsBusSeg.MoveNext
					loop
					if right(sBusSegIDs, 1) = "," then
						sBusSegIDs = left(sBusSegIDs, len(sBusSegIDs)-1)
					end if
				end if
				sBusSegIDs = replace(sBusSegIDs, " ", "")
				oRsBusSeg.MoveFirst
				Do until oRsBusSeg.EOF
					if sBusSegHTML <> "" then 
						sBusSegHTML = sBusSegHTML & "&nbsp;"
					end if
					if instr("," & sBusSegIDs & ",", "," & cstr(oRsBusSeg("BusinessSegmentID").Value) & ",") > 0 then 
						sBusSegHTML = sBusSegHTML & "<INPUT type='checkbox' id=chkBusSeg NAME=chkBusSeg value=" & oRsBusSeg("BusinessSegmentID").Value & " checked>" & oRsBusSeg("SegmentName").Value
					else
						sBusSegHTML = sBusSegHTML & "<INPUT type='checkbox' id=chkBusSeg NAME=chkBusSeg value=" & oRsBusSeg("BusinessSegmentID").Value & ">" & oRsBusSeg("SegmentName").Value
					end if
					oRsBusSeg.MoveNext
				Loop
			end if
		end if
	end if
end if

'Hide Owned By in Pulsar, not used
if sErr = "" then
    If bHideOwners = False Then
	    'Group Owners
	    'set oErr = GetCategory(oRsGroupOwner, 206, 0)
	    set oRsGroupOwner = GetCategory(206, 0)	
	    if oRsGroupOwner is Nothing then
		    Response.Write("Recordset error: oRsGroupOwner")
		    Response.End()
 	    else
		    if not oRsGroupOwner is nothing then
			    if oRsGroupOwner.RecordCount > 0 then
				    oRsGroupOwner.Sort = "Name ASC"
				    if sOwnerIDs = "" then
					    'make every one checked
					    oRsGroupOwner.MoveFirst
					    do until oRsGroupOwner.EOF
						    sOwnerIDs = sOwnerIDs & oRsGroupOwner("GroupID").Value & ","
						    oRsGroupOwner.MoveNext
					    loop
					    if right(sOwnerIDs, 1) = "," then
						    sOwnerIDs = left(sOwnerIDs, len(sOwnerIDs)-1)
					    end if
				    end if
				    sOwnerIDs = replace(sOwnerIDs, " ", "")
				    oRsGroupOwner.MoveFirst
				    Do until oRsGroupOwner.EOF
					    if sOwnersHTML <> "" then 
						    sOwnersHTML = sOwnersHTML & "&nbsp;"
					    end if
					    if instr("," & sOwnerIDs & ",", "," & cstr(oRsGroupOwner("GroupID").Value) & ",") > 0 then 
						    sOwnersHTML = sOwnersHTML & "<INPUT type='checkbox' id=chkGroupOwner NAME=chkGroupOwner value=" & oRsGroupOwner("GroupID").Value & " checked>" & oRsGroupOwner("Name").Value
					    else
						    sOwnersHTML = sOwnersHTML & "<INPUT type='checkbox' id=chkGroupOwner NAME=chkGroupOwner value=" & oRsGroupOwner("GroupID").Value & ">" & oRsGroupOwner("Name").Value
					    end if
					    oRsGroupOwner.MoveNext
				    Loop
			    end if
		    end if
	    end if
    End If
end if

set oSvr = nothing
set oSvrPAL = nothing

sHeader = "After Market Option List - Create Export Files"
%>
<!DOCTYPE html>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sHeader%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sHeader%></title>
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
<script language="JavaScript" src="../library/scripts/General.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<style>
INPUT.text200
{
	HEIGHT: 22px; 
	WIDTH: 200px
}
INPUT.text300
{
	HEIGHT: 22px; 
	WIDTH: 300px
}
INPUT.text1
{
	HEIGHT: 22px; 
	WIDTH: 20px
}
</style>
<SCRIPT type="text/javascript">
function btnSave_onclick() {
	//Save data, refresh the page then generate files
	var nextMode;
	var bNoBusSeg = 0, bNoOwnedBy = 0;
	var nMode = $("#inpMode").val();
	var iIsODM = $("#inpIsODM").val();
	
	// First make sure that at least one Business Segment and one Owned By is checked
	var oObject = document.getElementsByName("chkBusSeg")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoBusSeg = 1;
				break;
			}
		}
	}
	// if no business segment check boxes checked, then check all
	/*oObject = document.getElementsByName("chkGroupOwner")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoOwnedBy = 1;
				break;
			}
		}
	}*/
	if (bNoBusSeg == 0) {
		alert("Please select at least one Business Segment and one Owned By before proceeding");
		return false;
	}
	
	switch(nMode) {
	    case "3":
	        if(iIsODM == "0"){
	            nextMode = "2";
	        }
			break;
	    case "6":
	        if(iIsODM == "0"){
	            nextMode = "5";
	        }
			break;
		case "15":
			nextMode = "14";
			break;
		case "18":
			// Verify date
		    if (!checkDate (thisform.txtOld_Eff_DT, "Old_Eff_DT", true)){
		        return false;
		    }
			nextMode = "17";
			break;
	}
        thisform.action = "AMO_GenerateExport.asp?Mode=" + nextMode;
	    thisform.submit();
    }

    function btnCancel_onclick(){
	    thisform.action = "AMO_ModuleList_RAS.asp";
	    thisform.submit();
    }

    function lbFileType_onchange() {
        var nMode = $("#inpMode").val();
	    var listbox = thisform.lbFileType;
	    var nType = listbox.value;

	    if (nType != nMode) {
		    if (confirm("Any changes you have made will be lost when changing the File Type.\n\nAre you sure you want to change the File Type?")){
			    thisform.action = "AMO_GenerateExport.asp?Mode=" + nType;
			    thisform.submit ();
		    } else {
			    // put the drop down back to where it was
			    for (var i=0; i < listbox.length; i++) {
				    if (listbox.options[i].defaultSelected == true) {
					    listbox.options[i].selected = true;
					    break;
				    }
			    }
		    }
	    }
	    return false;
    }
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor="#FFFFFF">
<% 'insert the header, global navigation and overview links
'InsertGlobalNavigationBar_HomeParm(true)
'Response.write BuildHelp(sHeader, sHelpfile)
%>
<h2 class="page-title"><%=sHeader%></h2>
<hr size=1 width=100%>
<% 
if sErr <> "" then
	Response.Write sErr
else
	%>
	<FORM name=thisform method=post>
	File Type &nbsp; &nbsp; <select name="lbFileType" size="1" onchange='return lbFileType_onchange()' >
 														<% if IsODM = 0 then %><option value="3" <% if nMode = 3 then response.write "SELECTED" %>>GPSy NPI DFT</option>
														<% end if %><option value="18" <% if nMode = 18 then response.write "SELECTED" %>>GPSy DFT Select Availability (SA)</option>
														<option value="15" <% if nMode = 15 then response.write "SELECTED" %>>GPSy DFT Description</option><% if IsODM = 0 then %>
														<option value="6" <% if nMode = 6 then response.write "SELECTED" %>>PHweb NPI Autoload</option><% end if %>
													</select>
					&nbsp; &nbsp; <font size=1><i>Selecting a File Type occurs immediately</i></font>

	<TABLE width="100%" border=0 cellPadding="1" cellSpacing="5">
		<tr>
			<td width="150">Business Segment</td>
			<td><%= sBusSegHTML %>
			</td>
		</tr>
        <% if bHideOwners = False then %>					
		<tr>
			<td>Owned By</td>
			<td><%= sOwnersHTML %>
			</td>
		</tr>
        <%end if %>					
	</table>					
					
	<TABLE border=0 cellPadding=1 cellSpacing=1 width="100%">
		<TR>
			<TD><INPUT id=btnSave name=btnSave type=button value="Generate Files" LANGUAGE=javascript onclick="return btnSave_onclick()"></td>
			<TD colspan=2><INPUT id=btnCancel name=btnCancel type=button value="Cancel" LANGUAGE=javascript onclick="return btnCancel_onclick()"></td>
		</TR>
		<%
		select case nMode
			case 3
				DFTFields
			case 6
				AutoloadFields
			case 15
				DFTDescriptionFields
			case 18
				BlindDateFields
		end select
		%>
		<TR>
			<TD align=left colspan=3><hr size=1 width=100%></TD>
		</TR>
		<TR>
			<TD><INPUT id=btnSave name=btnSave type=button value="Generate Files" LANGUAGE=javascript onclick="return btnSave_onclick()"></td>
			<TD colspan=2><INPUT id=btnCancel name=btnCancel type=button value="Cancel" LANGUAGE=javascript onclick="return btnCancel_onclick()"></td>
		</TR>
	</TABLE>
	<input type=hidden name="PALName" value="<%= sPALName %>">
    <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
    <input type="hidden" id="inpMode" value="<%=nMode %>" />
    <input type="hidden" id="inpIsODM" value="<%=IsODM %>" />
	</FORM> 
	<%
end if 'no error


%>
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

<%
function DFTFields()
	%>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Mapped Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Base PN DFT File</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Localized DFT File</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">PROD_NBR</TD>
		<TD>HP Part Number</TD>
		<TD>HP Part Number + spaces + DASH Code</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>CD_DESC</TD>
		<TD>Option (Short Description)</TD>
		<TD>GPSy Description</TD>
	</TR>
	<TR>
		<TD>QU_DESC</TD>
		<TD>Option (Short Description)</TD>
		<TD>GPSy Description</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>START_EFF_DT</TD>
		<TD>Select Availability (SA)</TD>
		<TD>Select Availability (SA)</TD>
	</TR>
	<TR>
		<TD>LCLP</TD>
		<TD>AMO Price</TD>
		<TD>0 for all localizations</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>SUPP_CD</TD>
		<TD>Supplier Code</TD>
		<TD>Not used</TD>
	</TR>
	<TR>
		<TD>MFG_CD</TD>
		<TD>Manufacturing Code</TD>
		<TD>Not used</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>COM_CD</TD>
		<TD>COM Code</TD>
		<TD>Not used</TD>
	</TR>
	<TR>
		<TD>WTY_CD</TD>
		<TD>Warranty Code</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>PROD_CLASS_CD</TD>
		<TD>If Option is Hardware, then "HW". If Option is Software, then "SW".</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>AIR_PKG_VOL_QTY</TD>
		<TD>Air Packed Cubic</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>AIR_PKG_VOL_CD</TD>
		<TD>If AIR_PKG_VOL_QTY is 0, then "9". If AIR_PKG_VOL_QTY greater than 0, then "1".</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>AIR_PKG_WT_QTY</TD>
		<TD>Air Packed Weight</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>AIR_PKG_WT_CD</TD>
		<TD>If AIR_PKG_WT_QTY is 0, then "9". If AIR_PKG_WT_QTY greater than 0, then "0".</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>EXPORT_VOL_QTY</TD>
		<TD>Export Cubic</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>EXPORT_VOL_CD</TD>
		<TD>If EXPORT_VOL_QTY is 0, then "9". If EXPORT_VOL_QTY greater than 0, then "1".</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>EXPORT_WT_QTY</TD>
		<TD>Export Weight</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>EXPORT_WT_CD</TD>
		<TD>If EXPORT_WT_QTY is 0, then "9". If EXPORT_WT_QTY greater than 0, then "0".</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>NET_WT_QTY</TD>
		<TD>Net Weight</TD>
		<TD>Same</TD>
	</TR>
	<tr style="background-color: lightgrey;">
		<TD>NET_WT_CD</TD>
		<TD>If NET_WT_QTY is 0, then "9". If NET_WT_QTY greater than 0, then "0".</TD>
		<TD>Same</TD>
	</TR>

	<tr>
		<td colspan=3>&nbsp;</td>
	</tr>
	<TR>
		<TD colspan=3 style="background-color: lightsteelblue;"><font size=2><b>Pulsar Supplied Values</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">User Name</TD>
		<td class="FieldComment"><INPUT id=txtUserName name=txtUserName class="text300" value="<%=server.htmlencode(sUserName)%>" maxlength=64> 64 maximum characters</td>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>PROD_LINE</TD>
		<TD><%=lbxProdLineHTML%></TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<td>SUPP_CD_GBU</TD>
		<TD><%=lbxSuppCodeGBUHTML%></TD>
		<td>Not used</td>
	</TR>
	<TR>
		<td>INITIAL SUPP_CD</TD>
		<TD><%=lbxInitialSuppCodeHTML%></TD>
		<td>Not used</td>
	</TR>
	<TR>
		<TD>HW PROD_FAMILY (PROD_FAMILY)</TD>
		<TD class="FieldComment"><INPUT id=txtHW_PROD_FAMILY name=txtHW_PROD_FAMILY style="HEIGHT: 22px; WIDTH: 75px" value="<%=server.htmlencode(sHW_PROD_FAMILY)%>" maxlength=4 > 4 maximum characters. Used if Option is Hardware.</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>SW PROD_FAMILY (PROD_FAMILY)</TD>
		<TD class="FieldComment"><INPUT id=txtSW_PROD_FAMILY name=txtSW_PROD_FAMILY style="HEIGHT: 22px; WIDTH: 75px" value="<%=server.htmlencode(sSW_PROD_FAMILY)%>" maxlength=4 > 4 maximum characters. Used if Option is Software.</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>HW TAX_CLASS_CD (TAX_CLASS_CD)</TD>
		<TD class="FieldComment"><INPUT id=txtHW_TAX_CLASS_CD name=txtHW_TAX_CLASS_CD style="HEIGHT: 22px; WIDTH: 75px" value="<%=server.htmlencode(sHW_TAX_CLASS_CD)%>" maxlength=4 > 4 maximum characters. Used if Option is Hardware.</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>SW TAX_CLASS_CD (TAX_CLASS_CD)</TD>
		<TD class="FieldComment"><INPUT id=txtSW_TAX_CLASS_CD name=txtSW_TAX_CLASS_CD style="HEIGHT: 22px; WIDTH: 75px" value="<%=server.htmlencode(sSW_TAX_CLASS_CD)%>" maxlength=4 > 4 maximum characters. Used if Option is Software.</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Override_COM_CD</TD>
		<TD class="FieldComment"><INPUT id=txtOverride_COM_CD name=txtOverride_COM_CD style="HEIGHT: 22px; WIDTH: 50px" value="<%=server.htmlencode(sOverride_COM_CD)%>" maxlength=2> 2 maximum characters</TD>
		<td>Not used</td>
	</TR>
	<TR>
		<TD>SUPP_SEQ</TD>
		<TD>Automatically incremented starting from the number 2</TD>
		<td>Not used</td>
	</TR>

	<tr>
		<TD colspan=3>&nbsp;</td>
	</tr>
	<TR>
		<TD colspan=3 style="background-color: lightsteelblue;"><font size=2><b>User Defined Values</b></font></TD>
	</TR>
	<TR>
		<TD>A</TD>
		<TD class="FieldComment"><INPUT id=txtA name=txtA class="text200" value="<%=server.htmlencode(sA)%>" maxlength=16> 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>BUS_DEF_FIELD4</TD>
		<TD class="FieldComment"><INPUT id=txtBUS_DEF_FIELD4 name=txtBUS_DEF_FIELD4 class="text200" value="<%=server.htmlencode(sBUS_DEF_FIELD4)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>CTRY_CD</TD>
		<TD class="FieldComment"><INPUT id=txtCTRY_CD name=txtCTRY_CD class="text200" value="<%=server.htmlencode(sCTRY_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>CURR_CD</TD>
		<TD class="FieldComment"><INPUT id=txtCURR_CD name=txtCURR_CD class="text200" value="<%=server.htmlencode(sCURR_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>DIFF_CD</TD>
		<TD class="FieldComment"><INPUT id=txtDIFF_CD name=txtDIFF_CD class="text200" value="<%=server.htmlencode(sDIFF_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>ENTRY_SOURCE_CD</TD>
		<TD class="FieldComment"><INPUT id=txtENTRY_SOURCE_CD name=txtENTRY_SOURCE_CD class="text200" value="<%=server.htmlencode(sENTRY_SOURCE_CD)%>" maxlength=16> 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>M</TD>
		<TD class="FieldComment"><INPUT id=txtM name=txtM class="text200" value="<%=server.htmlencode(sM)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>MKT</TD>
		<TD class="FieldComment"><INPUT id=txtMKT name=txtMKT class="text200" value="<%=server.htmlencode(sMKT)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>MKT_CD</TD>
		<TD class="FieldComment"><INPUT id=txtMKT_CD name=txtMKT_CD class="text200" value="<%=server.htmlencode(sMKT_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Not used</TD>
	</TR>
	<TR>
		<TD>PA_DISC_FLG</TD>
		<TD class="FieldComment"><INPUT id=txtPA_DISC_FLG name=txtPA_DISC_FLG class="text200" value="<%=server.htmlencode(sPA_DISC_FLG)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>PRC_DISP_CD</TD>
		<TD class="FieldComment"><INPUT id=txtPRC_DISP_CD name=txtPRC_DISP_CD class="text200" value="<%=server.htmlencode(sPRC_DISP_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>PRC_TERM_CD</TD>
		<TD class="FieldComment"><INPUT id=txtPRC_TERM_CD name=txtPRC_TERM_CD class="text200" value="<%=server.htmlencode(sPRC_TERM_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>PROD</TD>
		<TD class="FieldComment"><INPUT id=txtPROD name=txtPROD class="text200" value="<%=server.htmlencode(sPROD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>PROD_DISP_EXCL_CD</TD>
		<TD class="FieldComment"><INPUT id=txtPROD_DISP_EXCL_CD name=txtPROD_DISP_EXCL_CD class="text200" value="<%=server.htmlencode(sPROD_DISP_EXCL_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>QBL_SEQ_NBR</TD>
		<TD class="FieldComment"><INPUT id=txtQBL_SEQ_NBR name=txtQBL_SEQ_NBR class="text200" value="<%=server.htmlencode(sQBL_SEQ_NBR)%>" maxlength=16> 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>SERIAL_FLG</TD>
		<TD class="FieldComment"><INPUT id=txtSERIAL_FLG name=txtSERIAL_FLG class="text200" value="<%=server.htmlencode(sSERIAL_FLG)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>SERV_CD</TD>
		<TD class="FieldComment"><INPUT id=txtSERV_CD name=txtSERV_CD class="text200" value="<%=server.htmlencode(sSERV_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Not used</TD>
	</TR>
	<TR>
		<TD>SRT_CD</TD>
		<TD class="FieldComment"><INPUT id=txtSRT_CD name=txtSRT_CD class="text200" value="<%=server.htmlencode(sSRT_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Not used</TD>
	</TR>
	<TR>
		<TD>SUI</TD>
		<TD class="FieldComment"><INPUT id=txtSUI name=txtSUI class="text200" value="<%=server.htmlencode(sSUI)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>UOM_CD</TD>
		<TD class="FieldComment"><INPUT id=txtUOM_CD name=txtUOM_CD class="text200" value="<%=server.htmlencode(sUOM_CD)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
<%
end function

function AutoloadFields()
	%>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Mapped Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Base Autoload Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Localized Autoload Data</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">Pin</TD>
		<TD>HP Part Number</TD>
		<TD>HP Part Number + # + DASH Code</TD>
	</TR>
	<TR>
		<TD>Category Group</TD>
		<TD>AMO PHweb Category from System Admin matching Module Category</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Short Description</TD>
		<TD>Option (Short Description)</TD>
		<TD>Option (Short Description) + PHweb Countrification Description</TD>
	</TR>
	<TR>
		<TD>Long Description</TD>
		<TD>Option (Short Description)</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Avail. Date</TD>
		<TD>Earliest PHweb (General) Availability (GA) in Pulsar</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Disc Build</TD>
		<TD>Latest End of Manufacturing (EM) in Pulsar</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Action</TD>
		<TD>"ADD"</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>Countrification</TD>
		<TD>User Defined - see below</TD>
		<TD>PHweb Countrification Description</TD>
	</TR>
 	<TR>
		<TD>Initial MLORW Target Cost</TD>
		<TD>AMO Price</TD>
		<td>Not used</td>
	</TR>

	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Supplied Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Base Autoload Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Localized Autoload Data</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">OBS Date</TD>
		<td>Last day of the month, 3 months after DISC Build Date</td>
		<td>Same</td>
	</TR>
	<TR>
		<TD width="20%">EOL Date</TD>
		<td>3 months prior to DISC Build Date, first Friday of that month</td>
		<td>Same</td>
	</TR>

	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>User Defined Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Base Autoload Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Localized Autoload Data</b></font></TD>
	</TR>
	<TR>
		<TD>Parent Pin</TD>
		<TD class="FieldComment"><INPUT id=txtParentPin name=txtParentPin class="text200" value="<%=server.htmlencode(sParentPin)%>" maxlength=16> 16 maximum characters</TD>
		<td>Same</td>
	</TR>
	<TR>
		<TD>Row Level 1</TD>
		<TD class="FieldComment"><INPUT id=txtRowLevel1 name=txtRowLevel1 class="text1" value="<%=server.htmlencode(sRowLevel1)%>" maxlength=1> 1 maximum character</TD>
		<td>Not used</td>
	</TR>
	<TR>
		<TD>Row Level 2</TD>
		<TD class="FieldComment"><INPUT id=txtRowLevel2 name=txtRowLevel2 class="text1" value="<%=server.htmlencode(sRowLevel2)%>" maxlength=1> 1 maximum character</TD>
		<td>Not used</td>
	</TR>
 	<TR>
		<TD>ODM Code</TD>
		<TD class="FieldComment"><INPUT id=txtODMCode name=txtODMCode class="text200" value="<%=server.htmlencode(sODMCode)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Same</td>
	</TR>
 	<TR>
		<TD>Format Code</TD>
		<TD class="FieldComment"><INPUT id=txtFormatCode name=txtFormatCode class="text200" value="<%=server.htmlencode(sFormatCode)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Same</td>
	</TR>
 	<TR>
		<TD>Brand Name</TD>
		<TD class="FieldComment"><INPUT id=txtBrandName name=txtBrandName class="text200" value="<%=server.htmlencode(sBrandName)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Same</td>
	</TR>
	<TR>
		<TD>Countrification</TD>
		<TD class="FieldComment"><INPUT id=txtCountrification name=txtCountrification class="text200" value="<%=server.htmlencode(sCountrification)%>" maxlength=16> 16 maximum characters</TD>
		<td>Pulsar Mapped - see above</td>
	</TR>
 	<TR>
		<TD>LC Status</TD>
		<TD class="FieldComment"><INPUT id=txtLCStatus name=txtLCStatus class="text200" value="<%=server.htmlencode(sLCStatus)%>" maxlength=16 > 16 maximum characters</TD>
		<TD class="FieldComment"><INPUT id=txtLCStatusLocal name=txtLCStatusLocal class="text200" value="<%=server.htmlencode(sLCStatusLocal)%>" maxlength=16 > 16 maximum characters</TD>
	</TR>
 	<TR>
		<TD>LC Status for Base with Localized</TD>
		<TD class="FieldComment"><INPUT id=txtLCStatusBaseWithLocal name=txtLCStatusBaseWithLocal class="text200" value="<%=server.htmlencode(sLCStatusBaseWithLocal)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Not used</td>
	</TR>
	<input type=hidden name="txtMLORW" value="">
<!--  	<TR>
		<TD>Initial MLORW Target Cost</TD>
		<TD class="FieldComment"><INPUT id=txtMLORW name=txtMLORW class="text200" value="<% 'server.htmlencode(sMLORW)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Not used</td>
	</TR> -->
 	<TR>
		<TD>SDF Flag</TD>
		<TD class="FieldComment"><INPUT id=txtSDFFlag name=txtSDFFlag class="text200" value="<%=server.htmlencode(sSDFFlag)%>" maxlength=16 > 16 maximum characters</TD>
		<TD class="FieldComment"><INPUT id=txtSDFFlagLocal name=txtSDFFlagLocal class="text200" value="<%=server.htmlencode(sSDFFlagLocal)%>" maxlength=16 > 16 maximum characters</TD>
	</TR>
 	<TR>
		<TD>SDF Flag for Base with Localized</TD>
		<TD class="FieldComment"><INPUT id=txtSDFFlagBaseWithLocal name=txtSDFFlagBaseWithLocal class="text200" value="<%=server.htmlencode(sSDFFlagBaseWithLocal)%>" maxlength=16 > 16 maximum characters</TD>
		<td>Not used</td>
	</TR>
 	<TR>
		<TD>Level</TD>
		<td>Not used</td>
		<TD class="FieldComment"><INPUT id=txtLevel name=txtLevel class="text200" value="<%=server.htmlencode(sLevel)%>" maxlength=16 > 16 maximum characters</TD>
	</TR>
<%
end function

function DFTDescriptionFields()
	%>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Mapped Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Common Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Quote Data</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">PROD_NBR</TD>
		<TD>HP Part Number</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>CD_DESC (used for DESC_TXT)</TD>
		<TD>Short Description</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>QU_DESC (used for DESC_TXT)</TD>
		<TD>Short Description</TD>
		<TD>Same</TD>
	</TR>

	<tr>
		<td colspan=3>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Supplied Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Common Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Quote Data</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">User Name</TD>
		<td class="FieldComment" colspan=2><INPUT id=txtUserName name=txtUserName class="text300" value="<%=server.htmlencode(sUserName)%>" maxlength=64> 64 maximum characters</td>
	</TR>
	<TR>
		<TD>DESC</TD>
		<TD>Always use "DESC"</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>LANG_CD</TD>
		<TD>Always use "99"</TD>
		<TD>Same</TD>
	</TR>
	<TR>
		<TD>DESC_TXT</TD>
		<TD>Use CD_DESC</TD>
		<TD>Use QU_DESC</TD>
	</TR>
	<TR>
		<TD>START_EFF_DT</TD>
		<TD>Tomorrow's date (<%= tomorrow() %>)</TD>
		<TD>Same</TD>
	</TR>

	<tr>
		<TD colspan=3>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>User Defined Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Common Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Quote Data</b></font></TD>
	</TR>
	<TR>
		<TD>DESC_CD</TD>
		<TD class="FieldComment"><INPUT id=txtDESC_CD_Common name=txtDESC_CD_Common class="text200" value="<%=server.htmlencode(sDESC_CD_Common)%>" maxlength=16 > 16 maximum characters</TD>
		<TD class="FieldComment"><INPUT id=txtDESC_CD_Quote name=txtDESC_CD_Quote class="text200" value="<%=server.htmlencode(sDESC_CD_Quote)%>" maxlength=16 > 16 maximum characters</TD>
	</TR>
	<TR>
		<TD>M</TD>
		<TD class="FieldComment"><INPUT id=txtM name=txtM class="text200" value="<%=server.htmlencode(sM)%>" maxlength=16 > 16 maximum characters</TD>
		<TD>Same</TD>
	</TR>
	<%
end function

function BlindDateFields()
	%>
	<tr style="background-color: lightsteelblue;">
		<TD><font size=2><b>Pulsar Mapped Values</b></font></td>
		<TD><font size=2><b>Pulsar Fields for Non-Localized Data</b></font></TD>
		<TD><font size=2><b>Pulsar Fields for Localized Data</b></font></TD>
	</TR>
	<TR>
		<TD width="20%">PROD_NBR</TD>
		<TD>HP Part Number</TD>
		<TD>HP Part Number + spaces + DASH Code</TD>
	</TR>
	<TR>
		<TD>New_Eff_DT</TD>
		<TD>Select Availability (SA)</TD>
		<TD>Select Availability (SA)</TD>
	</TR>

	<tr>
		<td colspan=3>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD colspan=3><font size=2><b>Pulsar Supplied Values</b></font></td>
	</TR>
	<TR>
		<TD width="20%">User Name</TD>
		<td class="FieldComment" colspan=2><INPUT id=txtUserName name=txtUserName class="text300" value="<%=server.htmlencode(sUserName)%>" maxlength=64> 64 maximum characters</td>
	</TR>
	<TR>
		<TD>DTE</TD>
		<TD colspan=2>Always use "DTE"</TD>
	</TR>
	<TR>
		<TD>Review_Type_CD</TD>
		<TD colspan=2>Always use "NEW"</TD>
	</TR>

	<tr>
		<TD colspan=3>&nbsp;</td>
	</tr>
	<tr style="background-color: lightsteelblue;">
		<TD colspan=3><font size=2><b>User Defined Values</b></font></td>
	</TR>
	<TR>
		<TD>Old_Eff_DT</TD>
		<TD colspan=2><input id="txtOld_Eff_DT" name="txtOld_Eff_DT" value="<%= sOld_Eff_DT %>" class="filter-dateselection" style="HEIGHT: 22px;" size=10 maxLength="10"> (MM/DD/YYYY)</TD>
	</TR>
	<TR>
		<TD>M</TD>
		<TD class="FieldComment" colspan=2><INPUT id=txtM name=txtM class="text200" value="<%=server.htmlencode(sM)%>" maxlength=16 > 16 maximum characters</TD>
	</TR>
	<%
end function
%>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->
