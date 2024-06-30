<%@ Language=VBScript %>
<%
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->

function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}



function window_onload() {
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
}

function ToggleInactiveCycles(){
    if (InactiveCycles.style.display != "none")
        {
        InactiveCycles.style.display = "none";
        CycleToggle.innerHTML = "more >>";
        }
    else
        {
        InactiveCycles.style.display = "";
        CycleToggle.innerHTML = "<< less";
        }

    
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
</STYLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<BODY LANGUAGE=javascript onload="return window_onload()">
<h3>Product Information Summary</h3>
<%
	dim cn
	dim rs
	dim strSeries
	dim SeriesArray
	dim RASName
	dim RowCount
	
	RowCount = 0
	
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.Open

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
iCycle = Trim(Request("Cycle"))
on error resume next
If iShowAll = "" Then iShowAll = Request.Cookies("ProdSummary")("iShowAll")
If iShowAll = "" Then iShowAll = 6
If iDevCenter = "" Then iDevCenter = Request.Cookies("ProdSummary")("iDevCenter")
If iDevCenter = "" Then iDevCenter = 0
If iCycle = "" Then iCycle = Request.Cookies("ProdSummary")("iCycle")
If iCycle = "" Then iCycle = 0
If m_PartnerID = 1 Then
	If iODM = "" Then iODM = Request.Cookies("ProdSummary")("iODM")
	If iODM = "" Then iODM = 0
Else
	iODM = m_PartnerID
End If

If iShowAll = 0 AND Not Request.Cookies("ProdSummary")("Clear") Then
	iShowAll = 6
End If

Response.Cookies("ProdSummary")("iShowAll") = iShowAll 
Response.Cookies("ProdSummary")("iDevCenter") = iDevCenter 
Response.Cookies("ProdSummary")("iODM") = iODM 
Response.Cookies("ProdSummary")("iCycle") = iCycle
Response.Cookies("ProdSummary")("Clear") = True
Response.Cookies("ProdSummary").Expires = Now() + 365 

on error goto 0
%>
<table class=DisplayBar Width=100% CellSpacing=0 CellPadding=2 >
	<tr>
		<td valign=top>
			<table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width=100%><table>
<%

Response.Write "<tr><td nowrap><b>Show:</b></td><td width='100%'>"
Select Case iShowAll
'	Case 0	'Show All
'		Response.Write _
'			"All&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5'>Inactive</a>"
	Case 1	'Definition
		Response.Write _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6&Cycle=" & iCycle & "'>All Active</a>&nbsp;|&nbsp;" & _
			"Definition&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
	Case 2	'Development
		Response.Write _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6&Cycle=" & iCycle & "'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"Development&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
	Case 3	'Production
		Response.Write _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6&Cycle=" & iCycle & "'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"Production&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
	Case 4	'Post-production
		Response.Write _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6&Cycle=" & iCycle & "'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"Post-Production&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
	Case 5	'Inactive
		Response.Write _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6&Cycle=" & iCycle & "'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"Inactive"
	Case 6	'All Active
		Response.Write _
			"All Active&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
	Case Else
		Response.Write _
			"All Active&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1&Cycle=" & iCycle & "'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2&Cycle=" & iCycle & "'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3&Cycle=" & iCycle & "'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4&Cycle=" & iCycle & "'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5&Cycle=" & iCycle & "'>Inactive</a>"
End Select
Response.Write "</td></tr>"

Response.Write "<tr><td nowrap><b>Dev Center:</b></td><td width='100%'>"
Select Case iDevCenter
Case 1
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"Houston&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
Case 2
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Taiwan&nbsp;(CNB)|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
Case 3
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"Taiwan&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
Case 4
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Singapore&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
Case 5
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Brazil&nbsp;(BNB)"
Case 0
	Response.Write _
		"All&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
Case Else
	Response.Write _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=0&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=1&Cycle=" & iCycle & "'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=2&Cycle=" & iCycle & "'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=3&Cycle=" & iCycle & "'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=4&Cycle=" & iCycle & "'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=5&Cycle=" & iCycle & "'>Brazil&nbsp;(BNB)</a>"
End Select
Response.Write "</td></tr>"

If m_PartnerID = 1 Then
	Response.Write "<tr><td nowrap><b>ODM:</b></td><td width='100%'>"
	Select Case iODM
	Case 3
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"Compal&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case 2
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"Inventec&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case 4
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"Quanta&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case 7
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"Wistron<br />"
	Case 10
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"Foxconn&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case 16
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"Flextronics&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case 0
		Response.Write _
			"All&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	Case Else
		Response.Write _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=0&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=3&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=16&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=10&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=2&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=4&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=7&DevCenter=" & iDevCenter & "&Cycle=" & iCycle & "'>Wistron</a><br />"
	End Select
	Response.Write "</td></tr>"
End If

'GetProgram List
strPrograms = ""
strProgramIDs = ""
strLastStatus = 1
if iCycle = 0 then
    blnFound = true
else
    blnFound = false
end if
set rs2 = server.CreateObject("ADODB.recordset")
rs2.open "spListProgramsAll",cn,adOpenStatic
do while not rs2.EOF
    'if trim(rs2("OTSCycleName") & "") = "" then
    if trim(rs2("programGroupID")) <> "3" then 
        strPrograms = strPrograms & "," & rs2("Program")
		strProgramIDs = strProgramIDs & "," & trim(rs2("ProgramID"))
		if instr(strProgramList& "," , "," & rs2("Program") & ",")=0 then
            if strLastStatus <> rs2("Active") then
                if not blnFound then
                    strProgramLinks = strProgramLinks & "<span id=InactiveCycles style=""display:"">"
                else
                    strProgramLinks = strProgramLinks & "<span id=InactiveCycles style=""display:none"">"
                end if
            end if
		    strprogramList = strProgramlist & "," & rs2("Program")
            strLastStatus = rs2("Active") 
            if iCycle = trim(rs2("ProgramID")) then
			    strProgramLinks = strProgramLinks & " | " & replace(rs2("Program")," ","&nbsp;")
		        if strLastStatus = 1 then
                    blnFound = true
                end if
        	else
			    strProgramLinks = strProgramLinks & " | <a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Cycle=" & rs2("ProgramID") & "'>"  & replace(rs2("Program")," ","&nbsp;") & "</a>"
            end if
        end if
    end if
    'else
     '   strPrograms = strPrograms & ",BNB " & rs2("Name")  
     '   strProgramIDs = strProgramIDs & "," & trim(rs2("ID"))
     '   if instr(strProgramList& "," , ",BNB " & rs2("Name") & ",")=0 then
     '       strprogramList = strProgramlist & ",BNB " & rs2("Name")
     '       if iCycle = trim(rs2("ID")) then
     '           strProgramLinks = strProgramLinks & " | " & "BNB " & rs2("Name")
     '       else
     '           strProgramLinks = strProgramLinks & " | <a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Cycle=" & rs2("ID") & "'>" & "BNB " & rs2("Name") & "</a>"
     '       end if
     '   end if
    'end if
    rs2.MoveNext
loop
rs2.Close	
set rs2 = nothing
if blnFound then
    strProgramLinks = strProgramLinks & "</span> | <a style=""color:purple"" id=""CycleToggle"" href=""javascript:ToggleInactiveCycles();"">more >></a>"
else
    strProgramLinks = strProgramLinks & "</span> | <a style=""color:purple"" id=""CycleToggle"" href=""javascript:ToggleInactiveCycles();""><< less</a>"
end if
if strPrograms="" then
    strPrograms = "&nbsp;"
else
    strPrograms = ucase(mid(strPrograms,2))
    strProgramIDs =strProgramIDs & ","
end if

	if strProgramLinks <>""then
		if iCycle = 0 then
			strProgramLinks = "<tr><td valign=top><font size=1 face=verdana><b>Cycle: </b></font></td><td><font size=1 face=verdana>All | " & mid(strProgramLinks,3) & "</font></td></tr>"
		else
			strProgramLinks = "<tr><td valign=top><font size=1 face=verdana><b>Cycle: </b></font></td><td><font size=1 face=verdana><a href='ProductSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Cycle=0'>All</a> | " & mid(strProgramLinks,3) & "</font></td></tr>"
		end if
	end if

response.Write strProgramLinks



%>
</table></td></tr></table>
<br /><div style="OVERFLOW-Y: scroll;OVERFLOW-X: scroll;width:98%;height:600px" id=DIV1>
<table ID=ProductTable cellpadding=2 cellspacing=1 bgcolor=LightGrey style="behavior:url(../includes/client/tablehl.htc);" slcolor='#FFFFCC' hlcolor='#BEC5DE' >
<%

	set rs = server.CreateObject("ADODB.recordset")

	rs.Open "SELECT * FROM vProductInfoSummary ORDER BY Product",cn,adOpenForwardOnly
%>

<thead>
	<tr bgcolor=Beige>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 0 ,1,2);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>ID</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 1 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Product</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 2 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Series</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 3 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Generation</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 4 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Division</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 5 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>ODM</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 6 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();" nowrap><u>Regulatory Model</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 7 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SystemID</u></th>
        <th class=pinnedRow onclick="SortTable( 'ProductTable', 8 ,2,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>End of Manufacturing</u></th>
        <th class=pinnedRow onclick="SortTable( 'ProductTable', 9 ,2,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>End of Service</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 10  ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();" style='width: 20em;' nowrap><u>Operating Systems (Preinstall)</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 11  ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();" style='width: 20em;' nowrap><u>Operating Systems (Web)</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 12 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>STL</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 13 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>CM/PM</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 14 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Platform Dev Mgr</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 15 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SE PM</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 16 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SE PE</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 17 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>PIN PM</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 18 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SE Test Lead</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 19 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Service</u></th>
	</tr>
</thead>
<tbody>
    
<%    	
	' gMaria: Now we show all the product types 
	' and (rs("TypeID") = 1) 
    do while not rs.EOF
        If (rs("Name")&"" <> "Test Product") _
			and (((clng(iShowAll) = clng(rs("ProductStatusID"))) or (iShowAll = 0)) or ((iShowAll = 6) and (clng(rs("ProductStatusID")) <> 5))) _
			and ((clng(iDevCenter) = clng(rs("DevCenter"))) or (iDevCenter = 0)) _
			and ((clng(iODM) = clng(rs("PartnerID"))) or (iODM = 0)) _
			and ((clng(iCycle) = clng(rs("ProgramID")&"")) or (iCycle = 0)) then
			if clng(rs("ProductStatusID"))< 5 then
				Response.Write "<tr bgcolor=""ivory"">"
				Response.Write "<td align=""left"" ><a target=""_blank"" href=""../pmView.asp?ID=" & rs("ID") & "&amp;List=General"">" & rs("ID") & "</a></td>"
				Response.Write "<td nowrap align=""left"" >" & rs("Name") & " " & rs("Version") & "</td>"
			else
				Response.Write "<tr bgcolor=""gainsboro"">"
				Response.Write "<td align=""left"" ><a target=""_blank"" href=""../pmView.asp?ID=" & rs("ID") & "&amp;List=General"">" & rs("ID") & "</a></td>"
				Response.Write "<td nowrap align=""left"">" & rs("Name") & " " & rs("Version") & "</td>"
			end if
			Response.Write "<td nowrap align=left>" & rs("SeriesList") & "&nbsp;</td>"
            Response.Write "<td align=left nowrap>" & rs("Generation") & "&nbsp;</td>"
            Response.Write "<td align=left nowrap>" & rs("DivisionType") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("Partner") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("RegulatoryModel") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("SystemBoardID") & "&nbsp;</td>"
			if rs("EndOfProduction") & "" = "" Then
			    Response.Write "<td>&nbsp;</td>"
			else
			    Response.Write "<td nowrap align=left>" & rs("EndOfProduction") & "</td>"
			end if
			if rs("ServiceLifeDate") & "" = "" Then
			    Response.Write "<td>&nbsp;</td>"
			Else
			    Response.Write "<td nowrap align=left>" & rs("ServiceLifeDate") & "</td>"
			End If
			Response.Write "<td align=left style='width: 20em;'>" & replace(rs("OsPin") & "" ,",",",<br>" ) & "&nbsp;</td>"
			Response.Write "<td align=left style='width: 20em;'>" & replace(rs("OsWeb") & "" ,",",",<br>" ) & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("STL") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("CM") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("PDM") & "&nbsp;</td>"
			Response.Write "<td nowrap align=left>" & rs("SEPM") & "&nbsp;</td>"
			Response.Write "<td align=left nowrap>" & rs("SEPE") & "&nbsp;</td>"
			Response.Write "<td align=left nowrap>" & rs("PINPM") & "&nbsp;</td>"
			Response.Write "<td align=left nowrap>" & rs("SETL") & "&nbsp;</td>"
            Response.Write "<td align=left nowrap>" & rs("ServiceManager") & "&nbsp;</td>"
			Response.Write "</tr>" & vbcrlf
			RowCount = RowCount + 1
		end if
		rs.MoveNext
	loop
	rs.Close
	

	set rs = nothing
	cn.Close
	set cn = nothing
	if RowCount = 0 then
		Response.write "<tr  bgcolor=""ivory""><td colspan=12><b>none</b></td></tr>"
	end if
%>
</tbody>
</table>
</div>
<br />
<span style="font-size:xx-small; font-family:Verdana;">Generated: <%=Now%></span>
</body>