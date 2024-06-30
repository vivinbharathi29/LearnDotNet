<%@ Language=VBScript %>
<%
	'Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=pdd_export.xls"
%>
<!-- #include file="AgencyPivot.asp" -->
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
<table border=1>
<%
m_from_where = "Certification"
Call DrawPMViewMatrix(Request("ID"), false)
%>
</table>
</BODY>
</HTML>
