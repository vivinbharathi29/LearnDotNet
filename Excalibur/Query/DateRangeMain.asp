<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<TITLE>Date Range</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../includes/Date.asp" -->

function cmdClear_onclick() {
	txtDateStart.value="";
	txtDateEnd.value="";
	var OutArray = new Array();
	OutArray[0]= "";
	OutArray[1]= "";
	window.returnValue = OutArray;
	window.parent.opener = self;
	window.parent.close();
	
}

function cmdDate_onclick(strField) {
	var strID;
	var i;
	
	strID = window.showModalDialog("../Mobilese/today/calDraw1.asp",window.document.all(strField).value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			window.document.all(strField).value = strID;
		}
}

function cmdOK_onclick() {
	var blnOK=true;
	
	if (txtDateStart.value != "")
		if (! isDate(txtDateStart.value))
			{
			alert("You must enter a valid Start Date.")
			blnOK=false;
			txtDateStart.focus();
			}
	if (txtDateEnd.value != "" && blnOK)
		if (! isDate(txtDateEnd.value))
			{
			alert("You must enter a valid End Date.")
			blnOK=false;
			txtDateEnd.focus();
			}
	if (blnOK)
		{
		var OutArray = new Array();
		OutArray[0]= txtDateStart.value;
		OutArray[1]= txtDateEnd.value;
		window.returnValue = OutArray;
		window.parent.opener = self;
		window.parent.close();
		}
}

function cmdCancel_onclick() {
		window.parent.opener = self;
		window.parent.close();
}

function window_onload() {
	txtDateStart.focus();
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	font-family: Verdana;
	font-size: x-small;
}
h1{
	font-family: Verdana;
	font-size: small;
	font-weight: bold;
}
</STYLE>
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">

<h1>Select Date Range</h1>
	<table width=100% ID="tabGeneral" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR><TD nowrap>Start Date:</TD>
		<TD>
		<INPUT type="text" id=txtDateStart name=txtDateStart value="<%=request("StartDate")%>">
			<a href="javascript: cmdDate_onclick('txtDateStart')"><img ID="picTarget" SRC="../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</TD>
	</TR>
	<TR><TD nowrap>End Date:</TD>
		<TD>
		<INPUT type="text" id=txtDateEnd name=txtDateEnd value="<%=request("EndDate")%>">
			<a href="javascript: cmdDate_onclick('txtDateEnd')"><img ID="picTarget" SRC="../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</TD>
	</TR>
	</Table>
	<Table border=0 width=100%>
	<TR>
		<TD colspan=2 align=right>
			<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
			<INPUT type="button" value="Clear" id=cmdClear name=cmdClear LANGUAGE=javascript onclick="return cmdClear_onclick()">
			<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
		</TD>
	</tr>
	</TABLE>
</BODY>
</HTML>
