<%@ Language=VBScript %>
<%
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim AppRoot : AppRoot = Session("ApplicationRoot")
Dim SAType
%>
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!-- #include file = "../_ScriptLibrary/sort.js" -->

var oPopup = window.createPopup();

function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}

function window_onload() {
}

function AddSA(SAType) {
    var strID
    strID = window.parent.showModalDialog("<%=AppRoot %>/MobileSE/SubAssemblyFrame.asp?SAType=" + SAType + "&Step=1" + "&pulsarplusDivId=SupplyChain", "", "dialogWidth:600px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
    window.location.reload();
}

function AddRemoveSA(SAType, FeatureCategoryID, SubassemblyID, BusinessID) {
	var strID
	strID = window.parent.showModalDialog("<%=AppRoot %>/MobileSE/SubAssemblyFrame.asp?SAType=" + SAType + "&Step=3&FeatureCategoryID=" + FeatureCategoryID + "&SubassemblyID=" + SubassemblyID + "&Existing=1" + "&BusinessID=" + BusinessID + "&pulsarplusDivId=SupplyChain", "", "dialogWidth:600px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
    window.location.reload();
}

function AvSa_onclick(FeatureCategoryID, ProductNoID, SAType, FeatureCategoryText, ProductNoText, BusinessID)
{
	if (window.event.srcElement.className != "text" && window.event.srcElement.className != "cell")
		return;

    var lefter = event.clientX;
    var topper = event.clientY;
    var popupBody;
	if (window.event.srcElement.className == "text")
		{
	    if (typeof(SelectedRow) != "undefined")
			if (SelectedRow != null)
				if (SelectedRow != window.event.srcElement.parentElement.parentElement)
					SelectedRow.style.color="black";
				
		SelectedRow = window.event.srcElement.parentElement.parentElement;
		SelectedRow.style.color="red";
			
		}
	else if (window.event.srcElement.className == "cell")
    	{
	    if (typeof(SelectedRow) != "undefined")
			if (SelectedRow != null)
				if (SelectedRow != window.event.srcElement.parentElement)
					SelectedRow.style.color="black";
					
		SelectedRow = window.event.srcElement.parentElement;
		SelectedRow.style.color="red";
    	}
	
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\"><SPAN><HR width=95%></SPAN>";


	//popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	//popupBody = popupBody + "<FONT face=Arial size=2>";
	//popupBody = popupBody + "<SPAN onclick=\"parent.location.href='mailto: " + strEmail + "?Subject=Excalibur Error&Body=" + strDesc + "'\" >&nbsp;&nbsp;&nbsp;Email&nbsp;User...</SPAN></FONT></DIV>";

	//popupBody = popupBody + "<DIV>";
	//popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
    popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddSA(" + SAType + ");'\" >&nbsp;&nbsp;&nbsp;Add&nbsp;New...</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
    popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddRemoveSA(" + SAType + "," + FeatureCategoryID + "," + ProductNoID + "," + BusinessID + ");'\" >&nbsp;&nbsp;&nbsp;Add&nbsp;/&nbsp;Remove&nbsp;Regions...</SPAN></FONT></DIV>";

	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

    oPopup.document.body.innerHTML = popupBody; 
	oPopup.show(lefter, topper, 170, 86, document.body);

	//Adjust window size
	var NewHeight;
	var NewWidth;

	NewHeight = oPopup.document.body.scrollHeight;
	NewWidth = oPopup.document.body.scrollWidth;
	oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
	
}

function AvSa_onmouseover() {
	if (window.event.srcElement.className=="text")
		{
		window.event.srcElement.parentElement.parentElement.style.color="red";
		window.event.srcElement.parentElement.parentElement.style.cursor="hand";
		}
	else
		{
		window.event.srcElement.parentElement.style.color="red";
		window.event.srcElement.parentElement.style.cursor="hand";
		}
}

function AvSa_onmouseout() {
	if (window.event.srcElement.className=="text")
		window.event.srcElement.parentElement.parentElement.style.color="black";
	else
		window.event.srcElement.parentElement.style.color="black";
}

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

</STYLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<BODY LANGUAGE=javascript onload="return window_onload()">
<br /><h3>Master Subassembly Assignment List</h3>
<%
	dim cn
	dim rs
	dim sLastSaDesc
    dim iRowCount
    dim sLocalizations
	
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

on error goto 0
%>
<table class=DisplayBar Width=100% CellSpacing=0 CellPadding=2 >
	<tr>
		<td valign=top>
			<table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width=100%><table>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 100 ALL
If Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 100 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>Linux</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 1 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "XP&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 2 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "Vista&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 10 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "Win7&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 11 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "WinE&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 100 And Request("Family") = 9 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "Linux&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 1 COMMERCIAL
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 100 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>Consumer&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Linux</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 1 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;" 
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "XP&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 2 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;" 
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "Vista&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 10 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Consumer&nbsp;</a>&nbsp;|&nbsp;" 
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "Win7&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 11 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>Consumer&nbsp;</a>&nbsp;|&nbsp;" 
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "WinE&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 1 And Request("Family") = 9 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Consumer&nbsp;</a>&nbsp;|&nbsp;" 
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "Linux&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 2 CONSUMER
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 100 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=100'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=100'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Linux</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 1 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=1'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "XP&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 2 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=2'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "Vista&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 10 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=10'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=10'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "Win7&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 11 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=11'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=11'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "WinE&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=9'>Linux</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("SAType") = 223 And Request("Business") = 2 And Request("Family") = 9 Then
    Response.Write "<tr><td nowrap><b>SA Type:</b></td><td width='100%'>"
    Response.Write "COA&nbsp;"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=1&Family=9'>Commercial&nbsp;</a>&nbsp;|&nbsp;"
    Response.Write "Consumer&nbsp;|&nbsp;"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=100&Family=9'>All&nbsp;</a>"
    Response.Write "</td></tr>"

    Response.Write "<tr><td nowrap><b>OS Family:</b></td><td width='100%'>"
    Response.Write "<a href='SubAssembly.asp?SAType=223&Business=2&Family=1'>XP&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=2'>Vista&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=10'>Win7&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=11'>WinE&nbsp;</a>&nbsp;|&nbsp;" & _
    "Linux&nbsp;|&nbsp;" & _
    "<a href='SubAssembly.asp?SAType=223&Business=2&Family=100'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If
%>
</table></td></tr></table>
<br />
<% 
SAType = request("SAType")
Response.Write "<table><span style=font-size: xx-small; font-family: Verdana><td><a href=""javascript:AddSA(" & SAType & ")"">Add New</a></span>"

If request("Filter") <> "" Then%>
   | <span style=font-size: xx-small; font-family: Verdana><a href="SubAssembly.asp?SAType=<%=request("SAType")%>&Business=<%=request("Business")%>&Family=<%=request("Family")%>">Show All Regions</a></span> 
     <span style=font-size: xx-small; font-family: Verdana>(Only Displaying Subassembies With <%=request("OptionConfig")%> Regions)</span>
<%End If

Response.Write("</td></table>")
%>
<br />
<table width=100% ID=ProductTable cellpadding=2 cellspacing=1 bgcolor=LightGrey slcolor='#FFFFCC' hlcolor='#BEC5DE' style="behavior:url(../includes/client/tablehl.htc);" >
<%
	set rs = server.CreateObject("ADODB.recordset")
	
    If request("Filter") <> "" Then
	    rs.Open "usp_SelectCOAMasterList " & clng(request("SAType")) & ","  & clng(request("Business")) & ","  & clng(request("Family")) & ","  & request("Filter"),cn,adOpenForwardOnly
	Else
	    rs.Open "usp_SelectCOAMasterList " & clng(request("SAType")) & ","  & clng(request("Business")) & ","  & clng(request("Family")),cn,adOpenForwardOnly
	End If
%>

<thead>
	<tr bgcolor=Beige>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 0 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>AV Feature Category</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 1 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SA Description</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 2 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>SA Part No</u></th>
		<th class=pinnedRow onclick="SortTable( 'ProductTable', 3 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();"><u>Regions</u></th>
	</tr>
</thead>
<tbody>
<%

sLastSaDesc = ""
iRowCount = 0
sLocalizations = ""

do while not rs.EOF
    If iRowCount = 0 Then
      %>
      <tr bgcolor="ivory" LANGUAGE="javascript" onmouseover="return AvSa_onmouseover()" onmouseout="return AvSa_onmouseout()" onclick="return AvSa_onclick(<%=rs("AvFeatureCategoryID")%>,<%=rs("DelRootID")%>,<%=request("SAType")%>,'<%=rs("AvFeatureCategory")%>','<%=rs("DelRootName")%>',<%=rs("BusinessID")%>)">
      <%
      Response.Write "<td nowrap align=left class=""cell"">" & rs("AvFeatureCategory") & "&nbsp;</td>"
	    Response.Write "<td nowrap align=left class=""cell"">" & rs("DelRootName") & "&nbsp;</td>"
	    Response.Write "<td nowrap align=left class=""cell"">" & rs("SubAssembly") & "&nbsp;</td>"
	    sLastSaDesc = rs("DelRootID") & rs("AvFeatureCategoryID")
	    sLocalizations = "<a href=SubAssembly.asp?SAType=" & request("SAType") & "&Business=" & request("Business") & "&Family=" & request("Family") & "&Filter=" & rs("RegionID") & "&OptionConfig=" & rs("OptionConfig") & ">" & rs("OptionConfig") & "</a>" 
	    iRowCount = iRowCount + 1
	ElseIf sLastSaDesc = rs("DelRootID") & rs("AvFeatureCategoryID") Then
	    sLocalizations = sLocalizations + ", " + "<a href=SubAssembly.asp?SAType=" & request("SAType") & "&Business=" & request("Business") & "&Family=" & request("Family") & "&Filter=" & rs("RegionID") & "&OptionConfig=" & rs("OptionConfig") & ">" & rs("OptionConfig") & "</a>" 
	Else
	    Response.Write "<td align=left>" & sLocalizations & "&nbsp;</td></tr>"
	    sLocalizations = ""
	    %>
        <tr bgcolor="ivory" LANGUAGE="javascript" onmouseover="return AvSa_onmouseover()" onmouseout="return AvSa_onmouseout()" onclick="return AvSa_onclick(<%=rs("AvFeatureCategoryID")%>,<%=rs("DelRootID")%>,<%=request("SAType")%>,'<%=rs("AvFeatureCategory")%>','<%=rs("DelRootName")%>',<%=rs("BusinessID")%>)">
        <%
        Response.Write "<td nowrap align=left class=""cell"">" & rs("AvFeatureCategory") & "&nbsp;</td>"
	    Response.Write "<td nowrap align=left class=""cell"">" & rs("DelRootName") & "&nbsp;</td>"
	    Response.Write "<td nowrap align=left class=""cell"">" & rs("SubAssembly") & "&nbsp;</td>"
	    sLastSaDesc = rs("DelRootID") & rs("AvFeatureCategoryID")
	    sLocalizations = "<a href=SubAssembly.asp?SAType=" & request("SAType") & "&Business=" & request("Business") & "&Family=" & request("Family") & "&Filter=" & rs("RegionID") & "&OptionConfig=" & rs("OptionConfig") & ">" & rs("OptionConfig") & "</a>" 
	End If
    rs.MoveNext
loop
Response.Write "<td align=left>" & sLocalizations & "&nbsp;</td></tr>"
rs.Close
	
set rs = nothing
cn.Close
set cn = nothing
if iRowCount = 0 then
    Response.write "<tr  bgcolor=""ivory""><td colspan=12><b>none</b></td></tr>"
end if

%>
</tbody>
</table>
<br />
<span style="font-size:xx-small; font-family:Verdana;">Generated: <%=Now%></span>
</body>