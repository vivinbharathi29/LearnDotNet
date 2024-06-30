<%@ Language=VBScript %>
<!--#INCLUDE FILE="ExcelExportForms.asp"-->
<%
Response.ContentType = "application/vnd.ms-excel"
'Response.ContentType = "application/msword"
%>
<HTML>
<HEAD>
<STYLE>
.HeaderRow
{
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
    COLOR: White;
    BACKGROUND-COLOR: Black;
}
.HeaderRow A
{
    COLOR: white;
    TEXT-DECORATION: none
}
.HeaderRow A:active
{
    COLOR: white;
    TEXT-DECORATION: none
}
.HeaderRow A:link
{
    COLOR: white;
    TEXT-DECORATION: none
}
.HeaderRow A:visited
{
    COLOR: white;
    TEXT-DECORATION: none
}
.Illegal
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Embargo
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Unknown
{
    FONT-FAMILY: Verdana;
    COLOR: black;
    FONT-SIZE: xx-small;
    BACKGROUND-COLOR: gold;
}
.Cert_Late
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Cert_Open
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: springgreen;
}
.Cert_Complete
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    FONT-WEIGHT: bold;
    COLOR: DarkGreen;
}
.NotNeeded
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
}
.NotSupported
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: white;
    BACKGROUND-COLOR: red;
}
.Changed
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: yellow;
}
.AgencyCell
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
}
.Region
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: silver;
}
.NS_AgencyCell
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Late
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Open
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_Cert_Complete
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: DarkGreen;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_NotNeeded
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
.NS_NotSupported
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: LightSteelBlue;
}
</STYLE>
</HEAD>
<BODY>

<% 
	if Request.TotalBytes > 60000 then
		'Response.Write "Unable to export that much data.  Try using the Advanced Query instead."
		Dim FormFields
		Set FormFields = GetForm

		Dim Field
		For Each Field In FormFields
			'Response.Write "<br>" & Field & ":" & Len(FormFields(Field))
			Response.Write FormFields(Field)
		Next
	else
	    Dim sOutput
	    sOutput = request.Form("txtData")
	    if instr(lcase(sOutput), "<script>") > 0 Then
            Response.Write "Invalid Input"
        else	        
		    Response.Write sOutput
		End If
	end if
	
%>
</BODY>
</HTML>
