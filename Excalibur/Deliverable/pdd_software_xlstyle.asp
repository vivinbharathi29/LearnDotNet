<%
Dim excelStyle
excelStyle = "<Styles>" & _
"<Style ss:ID=""HdrLeft"">" & _
	"<Alignment ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"</Borders>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"<Interior ss:Color=""#808080"" ss:Pattern=""Solid""/>" & _
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
	"<Interior ss:Color=""#808080"" ss:Pattern=""Solid""/>" & _
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
	"<Interior ss:Color=""#808080"" ss:Pattern=""Solid""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellLeft"">" & _
	"<Alignment ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""CellRight"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & _
	"<NumberFormat ss:Format=""@""/>" & _
	"</Style>" & _
"<Style ss:ID=""BotLeft"">" & _
	"<Alignment ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & _
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
"<Style ss:ID=""CellCategory"" ss:Parent=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>" & _
	"<Interior ss:Color=""#9A9A9A"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#FFFFFF"" />" & _
	"</Style>" & _
"<Style ss:ID=""CellRoot"" ss:Parent=""Cell"">" & _
	"<Borders>" & _
	"<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""2""/>" & _
	"<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/>" & _
	"</Borders>" & _
	"<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>" & _
	"<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>" & _
	"<Font ss:Bold=""1"" ss:Color=""#000000"" />" & _
	"</Style>" & _
	"</Styles>"
%>
