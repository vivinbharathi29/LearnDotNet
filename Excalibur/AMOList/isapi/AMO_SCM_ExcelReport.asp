<%@ language=vbscript %>
<%Server.ScriptTimeout = 6000 %>
<% 
Response.Expires = 0

'Response.Buffer = false
'Response.CacheControl = "No-cache"
'Response.Clear 
Response.contenttype="application/vnd.ms-excel" 
'Response.AddHeader "content-disposition", "inline; filename=MOL.xls"
%> 

<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../includes/AMO.inc" -->
<!-- #include file="../includes/AMO_SCM_ExcelReport.inc" -->
<%
	
dim sDescHTML, sErr, sStatusIDs, strEOLDate
dim sModuleCategoryHTML, sStatusHTML, sFilter, sName
dim nNumHWRequest, nNumSWRequest, nNumTotalRequest, nNumHWRequestAll, nNumSWRequestAll, nStatusID
dim nCommitment, nHistory, nCategoryID, nNumTotalModules
dim nMOLID, nHWTypeID, nSWTypeID, nViewType
dim bPMUpdate, bRAS
dim oSvr, oErr, oRs, oRsDL, oRsChassis, oRsStatus, oRsCategory, oRsMOL, oRsModuleData
dim oRsAMOModules, oRsAllCategory
dim lbxHWTypeHTML, lbxSWTypeHTML
dim bsnapshot, sModuleIDS
dim sBusSegIDs , sBusSegHTML
dim SCMID, PublishDate

SCMID = 1
sErr = ""
PublishDate = "N/A"

'sBusSegIDs = Request.Form("chkBusSeg1")
sBusSegIDs = Request.querystring("chkBusSeg1")
Call SaveDBCookie( "AMO chkBusSeg1", sBusSegIDs )

'nCategoryID = GetDBCookie( "AMO lbxCategory")
if trim(nCategoryID) = "" then
	nCategoryID = -1
else
	nCategoryID = clng(nCategoryID)
end if

sStatusIDs = cstr(Application("AMO_COMPLETE"))

'strEOLDate = Request.Form("txtEOLDate")
strEOLDate = Request.querystring("txtEOLDate")

Call SaveDBCookie( "AMO txtEOLDate", strEOLDate )

'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO
PublishDate = Request.QueryString("PublishDate")
if len(Request.QueryString("SCMID")) > 0 and Request.QueryString("SCMID")<> "1" then
	SCMID = clng(Request.QueryString("SCMID"))
	'PublishDate = Request.QueryString("PublishDate")
	sFilter = " and SCMID = " & SCMID		
	set oRsAMOModules = oSvr.AMOModule_Search(Application("REPOSITORY"), sFilter, "", null, null)
else
	sFilter = "  and O.SCMID = 1  and S.StatusID in (" & sStatusIDs & ")"
	if sBusSegIDs <> "" then 
		sFilter = sFilter & " and MD.DivisionID in (" & sBusSegIDs & ")" 
	end if
	
	if strEOLDate <> "" then
		sFilter = sFilter & " and (O.RASDiscontinueDate >= '" & strEOLDate & "' or O.RASDiscontinueDate is null) "
	end if
	set oRsAMOModules = oSvr.AMOModule_Search(Application("REPOSITORY"), sFilter, "", null, null)
end if 

if oRsAMOModules is nothing then
    sErr = "Missing required parameters.  Unable to complete your request."
	Response.Write(sErr)
	Response.End()
else
	oRsAMOModules.Sort = "Category ASC, Description ASC"
end if

'Snapshot part
bsnapshot = Request.QueryString("bsnapshot") 

if len(bsnapshot) > 0 then
	 'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO

	if oRsAMOModules.recordcount> 0 then
		do while not oRsAMOModules.eof
			sModuleIDS = sModuleIDS & oRsAMOModules("ModuleID") & ","
			oRsAMOModules.movenext 
		loop
			
		if len(sModuleIDS) > 0 then
			sModuleIDS = left(sModuleIDS, len(sModuleIDS)-1 )
		end if
	end if 
	'Response.Write sModuleIDS & "fdsufds"
	'Response.End
	
	if len(sModuleIDS) > 0 then
		sErr = oSvr.AMO_SCM_Publish(Application("REPOSITORY"), session("FullName"), sModuleIDS,SCMID )
		if sErr <> True then
			sErr = "Missing required parameters.  Unable to complete your request."
		    Response.Write(sErr)
		    Response.End()
		else
			'	Response.Write sModuleIDS & "fdsufds" &	SCMID
			'	Response.end
			Response.Redirect "AMO_Reports.asp?SCMID=" & SCMID
		end if	
	else	
		Response.Redirect "AMO_Reports.asp?SCMID=" & SCMID & "&Norecords=1"
	end if 
end if 
'end of snapshot
%>

<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
<TITLE>AMO_SCM Report</TITLE>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{mso-header-data:"&C&BAMO_SCM Report&B";
	mso-footer-data:"&L&BHP Confidential&B&C&D&RPage &P";
	margin:.75in .5in .75in .5in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-page-orientation:landscape;}
.font8
	{color:windowtext;
	font-size:8.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;}
.font10
	{color:windowtext;
	font-size:8.0pt;
	font-weight:700;
	font-style:italic;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;}
col
	{mso-width-source:auto;}
br
	{mso-data-placement:same-cell;}
.style0
	{mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;
	border:none;
	mso-protection:locked visible;
	mso-style-name:Normal;
	mso-style-id:0;}
tr
	{mso-height-source:auto;}
td
	{mso-style-parent:style0;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xlTitle
	{mso-style-parent:style0;
	font-weight:800;
	text-align:left;
	vertical-align:middle;
	white-space:nowrap;}
.xlHeading
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingBottomBorder	
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingyellow
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}	
	
	
.xlHeadingyellowgreen
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:yellowgreen;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingaquamarine
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:aquamarine;
	mso-pattern:auto none;
	white-space:normal;}
	
.xlHeadingWhite
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:left;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
	
.xlHeadingWhiteLarge
	{mso-style-parent:style0;
	font-size:15.0pt;
	font-weight:900;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:left;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingWhitemedium
	{mso-style-parent:style0;
	font-size:11.0pt;
	font-weight:900;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:left;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
	
.xlHeadingRotate
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: bottom;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:none;
	mso-rotate:90;	
	mso-pattern:auto none;
	white-space:normal;}
	
	
.xlHeadingRotatethin
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: bottom;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:none;
	mso-rotate:90;	
	mso-pattern:auto none;
	white-space:none;}
.xlHeadingRotateBottomborder
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: bottom;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:none;
	mso-rotate:90;	
	mso-pattern:auto none;
	white-space:normal;}
	
	
.xlHeadingLeft
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingLeftMedium
	{mso-style-parent:style0;
	font-size:11.0pt;
	font-weight:900;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:left;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlHeadingRight
	{mso-style-parent:style0;
	font-weight:700;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align: middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlCategory
	{mso-style-parent:style0;
	font-weight:700;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	text-align:left;
	background:none;
	mso-pattern:auto none;
	white-space:normal;}
.xlCategoryGroup
	{mso-style-parent:style0;
	font-weight:700;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:#CCCCCC;
	text-align:center;
	mso-pattern:auto none;
	white-space:nowrap;
	mso-ignore:colspan}
.xlContentLeft
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:left;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	white-space:normal;}
.xlContent
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	text-align:center;
	align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	white-space:normal;}
	
.xlContentGreen
	{mso-style-parent:style0;
	font-size:9.0pt;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	text-align:center;
	background:#00FF00;
	mso-pattern:auto none;
	white-space:normal;}
	
	
.xlContentRight
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-family:"Arial", sans-serif;
	text-align:center;
	mso-font-charset:0;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	white-space:normal;}


.xlLC
	{mso-style-parent:style0;
	font-size:8.0pt;
	font-family:"Arial", sans-serif;
	mso-font-charset:0;
	white-space:nowrap;}
.xlEmptyRow
	{mso-style-parent:style0;
	mso-ignore:rowspan;
	height:12.75pt;
	font-weight:700;}
.xlEmptyRowBorder
	{mso-style-parent:style0;
	mso-ignore:colspan;
	mso-ignore:rowspan;
	height:12.75pt;
	font-weight:700;
	border-top:1.0pt solid windowtext;}
}
.xl24
	{mso-style-parent:style0;
	font-size:9.0pt;
	font-weight:700;
	font-family:Wingdings;
	mso-generic-font-family:auto;
	mso-font-charset:2;
	mso-number-format:"\@";
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;}

-->
</style>
<xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>AMO_SCM Report</x:Name>
    <x:WorksheetOptions>
    <x:FitToPage/>
     <x:Print>
      <x:FitWidth>2</x:FitWidth>
	  <x:FitHeight>100</x:FitHeight>
      <x:ValidPrinterInfo/>
      <x:HorizontalResolution>600</x:HorizontalResolution>
      <x:VerticalResolution>0</x:VerticalResolution>
     </x:Print>
     <x:Selected/>
     <x:FreezePanes/>
     <x:SplitHorizontal>1</x:SplitHorizontal>
     <x:TopRowBottomPane>1</x:TopRowBottomPane>
     <x:ActivePane>2</x:ActivePane>
     <x:ProtectContents>False</x:ProtectContents>
     <x:ProtectObjects>False</x:ProtectObjects>
     <x:ProtectScenarios>False</x:ProtectScenarios>
    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
    <x:WindowHeight>15000</x:WindowHeight>
  <x:WindowWidth>18000</x:WindowWidth>
  <x:WindowTopX>300</x:WindowTopX>
  <x:WindowTopY>300</x:WindowTopY>
  <x:ProtectStructure>False</x:ProtectStructure>
  <x:ProtectWindows>False</x:ProtectWindows>
  
 </x:ExcelWorkbook>

  
 <x:ExcelName>
  <x:Name>Print_Titles</x:Name>
  <x:SheetIndex>1</x:SheetIndex>
  <x:Formula>='AMO_SCM Report'!$A:$B</x:Formula>
 </x:ExcelName>
 
</xml>

</head>
<body link=blue vlink=purple>
<%
if oRsAMOModules.recordcount> 0 then
	WriteAMO_SCM_ReportGridHTML oRsAMOModules, SCMID, PublishDate
else
	Response.Write "No After Market Option found that matches filter"
end if 
%>
</body>

</html>
<%
'free the report object
oRsAMOModules.Close
set oRsAMOModules = nothing

Server.ScriptTimeout = 60
%>
