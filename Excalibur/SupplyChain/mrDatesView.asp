<%@ Language=VBScript %>
<%
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include file = "../Includes/Common.asp" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->

/*
function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}
*/

function window_onload() {
/*
	if (window.frameElement == null)
	{
	window.DIV1.style.width = window.screen.availWidth - window.DIV1.style.left - 50;	
	window.DIV1.style.height = window.screen.availHeight - window.DIV1.style.top - 320;
	}
	else
	{
	window.DIV1.style.width = window.frameElement.width - window.DIV1.style.left - 50;	
	window.DIV1.style.height = window.frameElement.height - window.DIV1.style.top - 140;
	}
*/
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
TH
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

H3
{
    FONT-SIZE: small;
    FONT-FAMILY: Verdana;
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
<!--
.pinnedRow
{
	position: relative;
	top: expression(document.getElementById('DIV1').scrollTop);
	
}

.pinnedCol
{
	position: relative;
	left: expression(document.getElementById('DIV1').scrollLeft);
}

.pinnedColRow
{
	position: relative;
	top: expression(document.getElementById('DIV1').scrollTop);
	left: expression(document.getElementById('DIV1').scrollLeft);
}
-->
</STYLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<BODY LANGUAGE=javascript onload="return window_onload()">
<h3>Manufacturing Readiness Dates</h3>
<%

'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim m_PartnerID
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	
	m_PartnerID = Security.CurrentPartnerID

	Set Security = Nothing
'##############################################################################	

iShowAll = Trim(Request("ShowAll"))
iDevCenter = Trim(Request("DevCenter"))
iODM = Trim(Request("ODM"))
on error resume next
If iShowAll = "" Then iShowAll = Request.Cookies("ProdSummary")("iShowAll")
If iShowAll = "" Then iShowAll = 5
If iDevCenter = "" Then iDevCenter = Request.Cookies("ProdSummary")("iDevCenter")
If iDevCenter = "" Then iDevCenter = 0
If m_PartnerID = 1 Then
	If iODM = "" Then iODM = Request.Cookies("ProdSummary")("iODM")
	If iODM = "" Then iODM = 0
Else
	iODM = m_PartnerID
End If

If iShowAll = 0 AND Not Request.Cookies("ProdSummary")("Clear") Then
	iShowAll = 5
End If

Response.Cookies("ProdSummary")("iShowAll") = iShowAll 
Response.Cookies("ProdSummary")("iDevCenter") = iDevCenter 
Response.Cookies("ProdSummary")("iODM") = iODM 
Response.Cookies("ProdSummary")("Clear") = True
Response.Cookies("ProdSummary").Expires = Now() + 365 
on error goto 0
%>
<BR>
<table ID=ProductTable cellpadding=2 cellspacing=1 bgcolor=tan style="behavior:url(../includes/client/tablehl.htc);" slcolor='#FFFFCC' hlcolor='#BEC5DE' >
<%
	dim cn
	dim rs
	dim strSeries
	dim SeriesArray
	dim RASName
	dim i
	
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.CursorLocation = adUseClient
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")

	rs.Open "rpt_SelectAvMrDates_Pulsar " & Request.QueryString("BID"),cn, adOpenStatic
%>
<THEAD>
	<TR bgcolor=CornSilk>
		<TH><u>Av Number</u></TH>
		<TH><u>GPG Description</u></TH>
<%
	For i = 6 to rs.Fields.Count - 1 
		Response.Write "<TH><u>" & rs.Fields(i).Name & "</u></TH>"
	Next
%>		
	</TR>
</THEAD>
<TBODY>
<%

	do while not rs.EOF
	    If rs.Fields(0).Value & "" <> "" Then
		    Response.Write "<TR bgcolor=""ivory"">"
		    Response.Write "<TD align=left bgcolor=""CornSilk"">" & rs.Fields(0).Value & "</TD>"
		    Response.Write "<TD nowrap align=left bgcolor=""CornSilk"">" & rs.Fields(1).Value & "</TD>"
    			
		    For i = 6 to rs.Fields.Count - 1
			    Response.Write "<TD nowrap align=left>" & iif(Trim(rs.Fields(i).Value & "") = "","&nbsp;", Trim(rs.Fields(i).Value )) & "</TD>"
		    Next
		    Response.Write "</TR>" & vbcrlf
        End If
		rs.MoveNext
	loop
	rs.Close
	

	set rs = nothing
	cn.Close
	set cn = nothing
  
%>
</TBODY>
</table>
<BR>
<font size=1 face=verdana>Generated: <%=Now%></font>
</body>