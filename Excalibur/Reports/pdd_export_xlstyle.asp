<%
Dim excelStyle
excelStyle = "<Styles>" & _
"<Style ss:ID=""Default"">" & _
	"</Style>" & _
"<Style ss:ID=""TitleBorder"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Top"" ss:WrapText=""1""/>" & _
	"<Borders/>" & _
	"<Font x:Family=""Swiss""/>" & _
	"<Interior ss:Color=""#003366"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>"	& _
	"</Style>" & _
"<Style ss:ID=""TitleText"">" & _
	"<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>" & _
	"<Font x:Family=""Swiss"" ss:Size=""14"" ss:Bold=""1"" ss:Italic=""1""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""SubTitle"">" & _
	"<Alignment ss:Horizontal=""Left"" />" & _
	"<Font ss:Bold=""1"" />" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrLeft"">" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"<Interior ss:Color=""#0066FF"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Hdr"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"<Interior ss:Color=""#0066FF"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrSingle"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Size=""11"" ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"<Interior ss:Color=""#0066FF"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrRight"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"<Interior ss:Color=""#0066FF"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrRightCenter"" ss:Parent=""HdrRight"">" & _
	"<Alignment ss:Vertical=""Center"" ss:Horozontal=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrLeftCenter"" ss:Parent=""HdrLeft"">" & _
	"<Alignment ss:Horozontal=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrSingleCenter"">" & _
	"<Alignment ss:Horozontal=""Centered"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Size=""11""  ss:Color=""#FFFFFF"" ss:Bold=""1""/>" & _
	"<Interior ss:Color=""#0066FF"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""HdrCenter"" ss:Parent=""Hdr"">" & _
	"<Alignment ss:Horozontal=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellLeft"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Right"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""None""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellLeft_Date"" ss:Parent=""CellLeft"">" & _
	"<NumberFormat ss:Format=""mm/dd/yyyy;@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Cell_Date"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horozontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<NumberFormat ss:Format=""mm/dd/yyyy;@""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellCentered"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellRight"" ss:Parent=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellRightCentered"" ss:Parent=""CellRight"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellLeftCentered"" ss:Parent=""CellLeft"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellHighlight"" ss:Parent=""Cell"">" & _
	"<Interior ss:Color=""#FF3333"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"</Style>" & _
"<Style ss:ID=""CellLeftHighlight"" ss:Parent=""CellLeft"">" & _
	"<Interior ss:Color=""#FF3333"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"</Style>" & _
"<Style ss:ID=""CellRightHighlight"" ss:Parent=""CellRight"">" & _
	"<Interior ss:Color=""#FF3333"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"</Style>" & _
"<Style ss:ID=""CellRegion"" ss:Parent=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Interior ss:Color=""#008080"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"</Style>" & _
"<Style ss:ID=""CellSingle"" ss:Parent=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"</Style>" & _
"<Style ss:ID=""CellSingleCentered"" ss:Parent=""CellSingle"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"</Style>" & _
"<Style ss:ID=""BotLeft"">" & _
	"<Alignment ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Bot"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""BotRight"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""None""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Illegal"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#FFFFFF"" ss:Bold=""1""/>" & _
	"<Interior ss:Color=""#FF0000"" ss:Pattern=""Solid""/>" & _
	"</Style>" & _
"<Style ss:ID=""Embargo"" ss:Parent=""Illegal"">" & _
	"</Style>" & _
"<Style ss:ID=""Cert_Late"" ss:Parent=""Illegal"">" & _
	"</Style>" & _
"<Style ss:ID=""NotSupported"" ss:Parent=""Illegal"">" & _
	"</Style>" & _
"<Style ss:ID=""Unknown"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000""/>" & _
	"<Interior ss:Color=""#FFCC00"" ss:Pattern=""Solid""/>" & _
	"</Style>" & _
"<Style ss:ID=""NotNeeded"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8""/>" & _
	"</Style>" & _
"<Style ss:ID=""AgencyCell"" ss:Parent=""CellLeft"">" & _
	"<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8""/>" & _
	"</Style>" & _
"<Style ss:ID=""Cert_Open"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8""/>" & _
	"<Interior ss:Color=""#00FF00"" ss:Pattern=""Solid""/>" & _
	"</Style>" & _
"<Style ss:ID=""Cert_Complete"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#008000"" ss:Bold=""1""/>" & _
	"<Interior/>" & _
	"</Style>" & _
"<Style ss:ID=""Changed"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>" & _
	"<Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/>" & _
	"</Style>" & _
"<Style ss:ID=""Region"" ss:Parent=""CellRegion"">" & _
	"</Style>" & _
"<Style ss:ID=""NS_AgencyCell"" ss:Parent=""Cell"">" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8""/>" & _
	"<Interior ss:Color=""#99CCFF"" ss:Pattern=""Solid""/>" & _
	"</Style>" & _
"<Style ss:ID=""NS_Cert_Late"" ss:Parent=""NS_AgencyCell"">" & _
	"</Style>" & _
"<Style ss:ID=""NS_Cert_Complete"" ss:Parent=""NS_AgencyCell"">" & _
	"</Style>" & _
"<Style ss:ID=""NS_Cert_Open"" ss:Parent=""NS_AgencyCell"">" & _
	"</Style>" & _
"<Style ss:ID=""NS_NotNeeded"" ss:Parent=""NS_AgencyCell"">" & _
	"</Style>" & _
"<Style ss:ID=""NS_NotSupported"" ss:Parent=""NS_AgencyCell"">" & _
	"</Style>" & _
"<Style ss:ID=""ILMCell"">" & _
	"<Alignment ss:Horozontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""0""/>" & _
	"<Interior />" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""ILMCellLeft"" ss:Parent=""ILMCell"">" & _
	"</Style>" & _
"<Style ss:ID=""ILMCell_Date"" ss:Parent=""ILMCell"">" & _
	"<NumberFormat ss:Format=""mm/dd/yyyy;@""/>" & _
	"</Style>" & _
"<Style ss:ID=""ILMCol"">" & _
	"<Alignment ss:Horozontal=""Center""/>" & _
	"</Style>" & _
"</Styles>"
'  "<Style ss:ID=""Region"" ss:Parent=""CellRegion"">" & _
'   "<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & _
'   "<Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#FFFFFF"" ss:Bold=""1""/>" & _
'   "<Interior ss:Color=""#008080"" ss:Pattern=""Solid""/>" & _
'   "<Borders>" & _
'   "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
'   "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
'   "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
'   "</Borders>" & _
'  "</Style>" & _

%>
