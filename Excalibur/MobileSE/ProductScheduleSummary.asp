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
<h3>Product Schedule Summary</h3>
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
iBusiness = Trim(Request("Business"))
iODM = Trim(Request("ODM"))
on error resume next
If iShowAll = "" Then iShowAll = Request.Cookies("ProdSummary")("iShowAll")
If iShowAll = "" Then iShowAll = 6
If iDevCenter = "" Then iDevCenter = Request.Cookies("ProdSummary")("iDevCenter")
If iDevCenter = "" Then iDevCenter = 0
If iBusiness = "" Then iBusiness = Request.Cookies("ProdSummary")("iBusiness")
If iBusiness = "" Then iBusiness = 0
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
Response.Cookies("ProdSummary")("iBusiness") = iBusiness 
Response.Cookies("ProdSummary")("iODM") = iODM 
Response.Cookies("ProdSummary")("Clear") = True
Response.Cookies("ProdSummary").Expires = Now() + 365 

on error goto 0

%>
<table class=DisplayBar Width=100% CellSpacing=0 CellPadding=2 >
	<TR>
		<TD valign=top>
			<table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<TD width=100%><table>
<%
Response.Write "<tr><td nowrap><b>Show:</b></td><td width='100%'>"
Select Case iShowAll
'	Case 0	'Show All
'		Response.Write _
'			"All&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
'			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&ShowAll=5'>Inactive</a>"
	Case 1	'Definition
		Response.Write _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
			"Definition&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
	Case 2	'Development
		Response.Write _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"Development&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
	Case 3	'Production
		Response.Write _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"Production&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
	Case 4	'Post-Production
		Response.Write _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Production</a>&nbsp;|&nbsp;" & _
			"Post-Production&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
	Case 5	'Inactive
		Response.Write _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=6'>All Active</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Post-Production</a>&nbsp;|&nbsp;" & _
			"Inactive"
	Case 6	'All Active
		Response.Write _
			"All Active&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
	Case Else
		Response.Write _
			"All Active&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=1'>Definition</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=2'>Development</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=3'>Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=4'>Post-Production</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?DevCenter=" & iDevCenter & "&ODM=" & iODM & "&Business=" & iBusiness & "&ShowAll=5'>Inactive</a>"
End Select
Response.Write "</td></tr>"

Response.Write "<tr><td nowrap><b>Dev Center:</b></td><td width='100%'>"
Select Case iDevCenter
Case 1
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"Houston&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
Case 2
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Taiwan&nbsp;(CNB)&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
Case 3
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"Taiwan&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
Case 4
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Singapore&nbsp;(BNB)&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
Case 5
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"Brazil&nbsp;(BNB)</font>"
Case 0
	Response.Write _
		"All&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
Case Else
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=1'>Houston&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=2'>Taiwan&nbsp;(CNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=3'>Taiwan&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=4'>Singapore&nbsp;(BNB)</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&Business=" & iBusiness & "&DevCenter=5'>Brazil&nbsp;(BNB)</a></font>"
End Select
Response.Write "</td></tr>"

Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
Select Case iBusiness
Case 1
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
		"Commercial&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=2'>Consumer</a></font>"
Case 2
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
		"Consumer</font>" 
Case 0
	Response.Write _
		"All&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=2'>Consumer</a></font>"
Case Else
	Response.Write _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
		"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=" & iODM & "&DevCenter=" & iDevCenter & "&Business=2'>Consumer</a></font>" 
End Select
Response.Write "</td></tr>"

If m_PartnerID = 1 Then
	Response.Write "<tr><td nowrap><b>ODM:</b></td><td width='100%'>"
	Select Case iODM
	Case 3
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"Compal&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case 2
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"Inventec&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case 4
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"Quanta&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case 7
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"Wistron</font><BR>"
	Case 10
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"Foxconn&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case 16
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"Flextronics&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case 0
		Response.Write _
			"All&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	Case Else
		Response.Write _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=0" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>All</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=3" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Compal</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=16" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Flextronics</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=10" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Foxconn</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=2" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Inventec</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=4" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Quanta</a>&nbsp;|&nbsp;" & _
			"<a href='ProductScheduleSummary.asp?ShowAll=" & iShowAll & "&ODM=7" & "&Business=" & iBusiness & "&DevCenter=" & iDevCenter & "'>Wistron</a></font><BR>"
	End Select
	Response.Write "</td></tr>"
End If

%>
</table></td></tr></table><BR>
<div style="OVERFLOW-Y: scroll;OVERFLOW-X: scroll;width:800px;height:600px" id=DIV1>
<table ID=ProductTable cellpadding=2 cellspacing=1 bgcolor=LightGrey style="behavior:url(../includes/client/tablehl.htc);" slcolor='#FFFFCC' hlcolor='#BEC5DE' >
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

	rs.Open "rpt_ProgScheduleSummary",cn, adOpenStatic
%>
<THEAD>
	<TR bgcolor=Beige>
		<TH class=pinnedRow style="z-index:1;" onclick="SortTable( 'ProductTable', 0 ,1,2);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>ID</u></TH>
		<TH class=pinnedRow style="z-index:2;" onclick="SortTable( 'ProductTable', 1 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Product Version</u></TH>
		<TH class=pinnedRow onclick="SortTable( 'ProductTable', 2 ,2,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>PreInstall Cut-Off</u></TH>
<%
	For i = 1 to rs.Fields.Count - 1 
		Response.Write "<TH class=pinnedRow onclick=""SortTable( 'ProductTable', " & i + 2 & " ,2,1);"" onmouseover=""HeaderMouseOver();"" onmouseout=""HeaderMouseOut();""><u>" & rs.Fields(i).Value & "</u></TH>"
		'i = i + 1
	Next
%>		
	</TR>
</THEAD>
</TBODY>
<%
	Set rs = rs.NextRecordset()
	rs.Sort = "familyname, version, schedulename"

	do while not rs.EOF
'		if rs("familyname") <> "Test Product" and rs("TypeID") = 1 _
		if rs("TypeID") = 1 _
			and rs("ID") <> 100 _
			and (((clng(iShowAll) = clng(rs("ProductStatusID"))) or (iShowAll = 0)) or ((iShowAll = 6) and (clng(rs("ProductStatusID")) <> 5))) _
			and (clng(iDevCenter) = clng(rs("DevCenter")) or iDevCenter = 0) _
			and (( (clng(iBusiness) = 2 and clng(rs("DevCenter"))=2) or (clng(iBusiness) =1 and clng(rs("DevCenter"))<>2)  ) or iBusiness = 0) _
			and (clng(iODM) = clng(rs("partnerid")) or iODM = 0) then

			if rs("ProductStatusID") <> 5 then
				Response.Write "<TR bgcolor=""ivory"" style=""height:30px;"">"
				Response.Write "<TD align=""left""><a target=""_blank"" href=""../pmView.asp?ID=" & rs("ID") & "&amp;List=General"">" & rs("ID") & "</a></TD>"
				Response.Write "<TD nowrap align=""left"">" & rs("familyname") & " " & rs("schedulename") & "</TD>"
			else
				Response.Write "<TR bgcolor=""gainsboro"" style=""height:30px;"">"
				Response.Write "<TD align=""left""><a target=""_blank"" href=""../pmView.asp?ID=" & rs("ID") & "&amp;List=General"">" & rs("ID") & "</a></TD>"
				Response.Write "<TD nowrap align=""left"">" & rs("familyname") & " " & rs("schedulename") & "</TD>"
			end if
			
			Response.Write "<TD nowrap align=left>" & iif(Trim(rs("preinstallcutoff")&"") = "","&nbsp;", Trim(rs("preinstallcutoff"))) & "</TD>"
			For i = 21 to rs.Fields.Count - 3
				Response.Write "<TD nowrap align=left>" & iif(Trim(rs.Fields(i).Value & "") = "","&nbsp;", Trim(rs.Fields(i).Value )) & "</TD>"
			Next
			Response.Write "</TR>" & vbcrlf

		end if
		rs.MoveNext
	loop
	rs.Close
	

	set rs = nothing
	cn.Close
	set cn = nothing
  
%>
</TBODY>
</table>
</div>
<BR>
<font size=1 face=verdana>Generated: <%=Now%></font>
</body>