<%@  Language=VBScript %>
<% OPTION EXPLICIT %>
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

dim sHeader, sErr, sBusSegIDs, sOwnerIDs
dim strSUI, strM, strProd_Nbr, strServ_CD, strMKT_CD, strSRT_CD, strStart_Eff_Dt
dim intSeq
dim nUserID, nDFTfile
dim oSvr, oErr, oRsOptions, oRsSuppliers
dim bGiveBlank
dim hasSuppliers
dim FeatureID
sHeader = "After Market Option List - Create Files"
sErr = ""

nUserID = Session("AMOUserID")

'get the cookies from the Generate Export filter.
sBusSegIDs = GetDBCookie("AMO Export_chkBusSeg")
sOwnerIDs = GetDBCookie("AMO Export_chkGroupOwner")


if sBusSegIDs = "" And sOwnerIDs = "" then
	sErr = "Missing Business Segment and/or Owned By Parameters, AMO_CreateDFTFile.asp"
	Response.Write(sErr)
    Response.End()
end if

if sErr = "" then
	if Request.Form("DFTfile") = "" then
	    sErr = "Missing DFT File type parameter, AMO_CreateDFTFile.asp"
	    Response.Write(sErr)
        Response.End()	
	else
		nDFTfile = clng(Request.Form("DFTfile"))
	end if
end if

if sErr = "" then
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO

     
	if nDFTfile = 0 then
		'Base Part Number file
		set oRsOptions = oSvr.AMO_GetDFTFileData_Base(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID)		
		if oRsOptions is nothing then
			sErr = "Empty Recordset - GetDFTFileData_Base, AMO_CreateDFTFile.asp"
	        Response.Write(sErr)
            Response.End()
		end if
      
        set oRsSuppliers = oSvr.AMO_GetDFTFileSuppliers_Base(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID)		
		if oRsSuppliers is nothing then
			sErr = "Empty Recordset - GetDFTFileSuppliers_Base, AMO_CreateDFTFile.asp"
	        Response.Write(sErr)
            Response.End()
		end if
      if not oRsSuppliers.EOF then
        hasSuppliers=true
       end if 
             
	else
		'Localized file
		set oRsOptions = oSvr.AMO_GetDFTFileData_Localized(Application("REPOSITORY"), sBusSegIDs, sOwnerIDs, nUserID)		
		if oRsOptions is nothing then
			sErr = "Empty Recordset - GetDFTFileData_Localized, AMO_CreateDFTFile.asp"
	        Response.Write(sErr)
            Response.End()
		end if
	end if
end if

if sErr = "" then
	Response.clear()
	Response.ContentType = "text/plain"
	if nDFTfile = 0 then
		Response.AddHeader "content-disposition", "attachment; filename=Option.txt"
		SectionOneTwo
		
		if hasSuppliers   then
			response.write vbCrLf
           
			response.write "MSK" & vbTab
				response.write "SUI" & vbTab & "M" & vbTab & "PROD_NBR" & vbTab
				response.write "SUPP_SEQ" & vbTab & "SUPP_CD" & vbTab & "SERV_CD" & vbTab
				response.write "MFG_CD" & vbTab & "MKT_CD" & vbTab & "SRT_CD" & vbTab
				response.write "COM_CD" & vbTab & "START_EFF_DT" & vbCrLf
	        
			'oRsOptions.MoveFirst
			bGiveBlank = False
            FeatureID = 0
			do while not oRsSuppliers.EOF
				if bGiveBlank then
					response.write vbCrLf
				end if
				bGiveBlank = True
				strSUI = oRsSuppliers("SUI")
				strM = oRsSuppliers("M")
				strProd_Nbr = oRsSuppliers("PROD_NBR")
				strServ_CD = oRsSuppliers("SERV_CD")
				strMKT_CD = oRsSuppliers("MKT_CD")
				strSRT_CD = oRsSuppliers("SRT_CD")
				strStart_Eff_Dt = makeDateFormat(oRsSuppliers("START_EFF_DT"))
				if cint( oRsSuppliers("FeatureID")) <> FeatureID then			
				    intSeq = 2
                else
    				intSeq = intSeq + 1	
                end if
                FeatureID =  cint( oRsSuppliers("FeatureID"))
				response.write strSUI & vbTab & strM & vbTab & strProd_Nbr & vbTab
				response.write cstr(intSeq) & vbTab & oRsSuppliers("SUPP_CD") & vbTab & strServ_CD & vbTab
				response.write oRsSuppliers("MFG_CD") & vbTab & strMKT_CD & vbTab & strSRT_CD & vbTab
				response.write oRsSuppliers("COM_CD") & vbTab & strStart_Eff_Dt & vbCrLf					
				
	
				oRsSuppliers.MoveNext
			loop
		end if
			
	else
		Response.AddHeader "content-disposition", "attachment; filename=AMO_GBULocalizations.txt"
		SectionOneTwo
	end if
	
	response.write "END" & vbCrLf

	response.end
end if
%>
<HTML>
<head>
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>
</HEAD>
<BODY>
<FORM name=thisform method=post>
<%

'Response.Write BuildHelp(sHeader, "")

Response.Write sErr


%>
</FORM>
</BODY>
</HTML>

<%
function SectionOneTwo()
	dim i
	
	response.write oRsOptions("UserName") & vbCrLf
	response.write "MSK" & vbTab
		response.write "PROD" & vbTab & "A" & vbTab & "PROD_LINE" & vbTab 
		response.write "DIFF_CD" & vbTab & "PROD_NBR" & vbTab & "PROD_FAMILY" & vbTab
		response.write "ENTRY_SOURCE_CD" & vbTab & "PROD_CLASS_CD" & vbTab & "PRC_DISP_CD" & vbTab 
		response.write "TAX_CLASS_CD" & vbTab & "CD_DESC" & vbTab & "QU_DESC" & vbTab 
		if nDFTfile = 0 then
			response.write "SUPP_CD" & vbTab & "SERV_CD" & vbTab & "MFG_CD" & vbTab
			response.write "MKT_CD" & vbTab & "SRT_CD" & vbTab
		end if
		response.write "SERIAL_FLG" & vbTab 
		if nDFTfile = 0 then
			response.write "COM_CD" & vbTab
		end if
		response.write "NET_WT_QTY" & vbTab & "NET_WT_CD" & vbTab 
		response.write "AIR_PKG_WT_QTY" & vbTab & "AIR_PKG_WT_CD" & vbTab & "AIR_PKG_VOL_QTY" & vbTab 
		response.write "AIR_PKG_VOL_CD" & vbTab & "EXPORT_WT_QTY" & vbTab & "EXPORT_WT_CD" & vbTab 
		response.write "EXPORT_VOL_QTY" & vbTab & "EXPORT_VOL_CD" & vbTab & "START_EFF_DT" & vbTab 
		response.write "PROD_DISP_EXCL_CD" & vbTab & "WTY_CD" & vbTab & "PA_DISC_FLG" & vbTab 
		response.write "BUS_DEF_FIELD4" & vbTab & "UOM_CD" & vbCrLf
	
	do while not oRsOptions.EOF
		response.write oRsOptions("PROD") & vbTab & oRsOptions("A") & vbTab & oRsOptions("PROD_LINE") & vbTab 
		response.write oRsOptions("DIFF_CD") & vbTab & oRsOptions("PROD_NBR") & vbTab & oRsOptions("PROD_FAMILY") & vbTab 
		response.write oRsOptions("ENTRY_SOURCE_CD") & vbTab & oRsOptions("PROD_CLASS_CD") & vbTab & oRsOptions("PRC_DISP_CD") & vbTab
		response.write oRsOptions("TAX_CLASS_CD") & vbTab
		response.write stripFancyChars(oRsOptions("CD_DESC")) & vbTab
		response.write stripFancyChars(oRsOptions("QU_DESC")) & vbTab
		if nDFTfile = 0 then
			response.write oRsOptions("SUPP_CD") & vbTab & oRsOptions("SERV_CD") & vbTab & oRsOptions("MFG_CD") & vbTab
			response.write oRsOptions("MKT_CD") & vbTab & oRsOptions("SRT_CD") & vbTab
		end if
		response.write oRsOptions("SERIAL_FLG") & vbTab
		if nDFTfile = 0 then
			response.write oRsOptions("COM_CD") & vbTab
		end if
		response.write oRsOptions("NET_WT_QTY") & vbTab & oRsOptions("NET_WT_CD") & vbTab
		response.write oRsOptions("AIR_PKG_WT_QTY") & vbTab & oRsOptions("AIR_PKG_WT_CD") & vbTab & oRsOptions("AIR_PKG_VOL_QTY") & vbTab
		response.write oRsOptions("AIR_PKG_VOL_CD") & vbTab & oRsOptions("EXPORT_WT_QTY") & vbTab & oRsOptions("EXPORT_WT_CD") & vbTab
		response.write oRsOptions("EXPORT_VOL_QTY") & vbTab & oRsOptions("EXPORT_VOL_CD") & vbTab & makeDateFormat(oRsOptions("START_EFF_DT")) & vbTab
		response.write oRsOptions("PROD_DISP_EXCL_CD") & vbTab & oRsOptions("WTY_CD") & vbTab & oRsOptions("PA_DISC_FLG") & vbTab
		response.write oRsOptions("BUS_DEF_FIELD4") & vbTab & oRsOptions("UOM_CD") & vbCrLf

		oRsOptions.MoveNext
	loop
	
	response.write vbCrLf
	'oRsOptions.MoveFirst

	response.write "MSK" & vbTab
		response.write "MKT" & vbTab & "M" & vbTab & "PROD_NBR" & vbTab
		response.write "START_EFF_DT" & vbTab & "CTRY_CD" & vbTab & "CURR_CD" & vbTab
		response.write "PRC_TERM_CD" & vbTab & "QBL_SEQ_NBR" & vbTab & "LCLP" & vbCrLf

	do while not oRsOptions.EOF
		response.write oRsOptions("MKT") & vbTab & oRsOptions("M") & vbTab & oRsOptions("PROD_NBR") & vbTab
		response.write makeDateFormat(oRsOptions("START_EFF_DT")) & vbTab & oRsOptions("CTRY_CD") & vbTab & oRsOptions("CURR_CD") & vbTab
		response.write oRsOptions("PRC_TERM_CD") & vbTab & oRsOptions("QBL_SEQ_NBR") & vbTab
		
		if nDFTfile = 0 then
			Response.Write oRsOptions("LCLP") & vbCrLf
		else
			'For localized AVs, all localizations are 0 except UUW which is -13: suggestion 5983
			'if oRsOptions("DASHCode").Value = "UUW" then
			'	response.write "-13" & vbCrLf
			'else
				response.write "0" & vbCrLf
			'end if
		end if

		oRsOptions.MoveNext
	loop
end function
%>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->