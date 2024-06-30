<%@ language=vbscript %>
<% Server.ScriptTimeout = 6000
Response.Expires = 0

'Response.Buffer = false
'Response.CacheControl = "No-cache"
'Response.Clear 
Response.contenttype="application/vnd.ms-excel" 
'Response.AddHeader "content-disposition", "inline; filename=MOL.xls"
%> 
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../includes/AMO_SCM_ExcelReport.inc" -->
<%
	dim sErr, oSvr, oErr, rsAVAMO, RsTwoAMOPublishs, rsModules
	dim sSCMIDS, sFilter, sBusSegIDs , strproductfamily
	
	sFilter = ""
	strproductfamily = Request.Form("cboPdfamily")
	if strproductfamily <> "" then
		sFilter = strproductfamily
	end if
	
	'Response.Write sFilter & "abc"
	'Response.End
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO	
    			
	set rsModules = oSvr.AVAMO_Modules(Application("REPOSITORY"), sFilter)
	if rsModules is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		set oSvr=nothing
		Response.End()		
	end if
	
	
	set rsAVAMO = oSvr.AVAMO_Matrix(Application("REPOSITORY"), sFilter, rsAVAMO)
	if rsAVAMO is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		set oSvr=nothing
		Response.End()
	end if
%>

<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
<TITLE>AVAMO Matrix Report</TITLE>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{mso-header-data:"&C&BAVAMO Matrix Report&B";
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
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
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
	background:skyblue;
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
	text-align:center;
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
	vertical-align: middle;
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
	vertical-align: middle;
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
	vertical-align: middle;
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
.xlChangeType
	{mso-style-parent:style0;
	font-weight:700;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	text-align:left;
	background:silver;
	mso-pattern:auto none;
	white-space:normal;}
.xlCategoryGroup
	{mso-style-parent:style0;
	font-weight:700;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:#FFFF99;
	text-align:left;
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
	border-left:.5pt solid windowtext;
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
    <x:Name>AVAMO Matrix Report</x:Name>
    <x:WorksheetOptions>
    <x:FitToPage/>
     <x:Print>
      <x:FitWidth>1</x:FitWidth>
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
  <x:Formula>='AVAMO_Matrix Report'!$A:$B</x:Formula>
 </x:ExcelName>
 
</xml>

</head>
<body link=blue vlink=purple>
<%
if rsAVAMO.recordcount> 0 and rsModules.recordcount> 0 then
	Write_AVAMO_MAtrix rsAVAMO, rsModules
else
	Response.Write "No data found that matches the filter"
end if 
%>
</body>
</html>
<%	
'free the report object
'rsAVAMO.Close
'set rsAVAMO = nothing
Server.ScriptTimeout = 60
%>

