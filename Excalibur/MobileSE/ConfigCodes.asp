<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>Master Localization List</TITLE>
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


//-->
</SCRIPT>
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<STYLE>
<!--
TD
{
	border-right: LightGrey 1px solid;
	border-left: LightGrey 1px solid;
}
TH
{
	border-right: LightGrey 1px solid;
	border-left: LightGrey 1px solid;
}
//-->
</STYLE>
</HEAD>

<BODY style="background-color:white">
<h4>Master Localization List</h4>
<br />
<table class=DisplayBar Width=100% CellSpacing=0 CellPadding=2 >
	<tr>
		<td valign=top>
			<table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table>
		<td width=100%><table>
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 0 ALL
If Request("Business") = 0 And Request("Status") = 0 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 1 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 2 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 3 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 0 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 0 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 0 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 1 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 1 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 1 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 2 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 2 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 2 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 3 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 3 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 0 And Request("Status") = 3 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 1 COMMERCIAL
ElseIf Request("Business") = 1 And Request("Status") = 0 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 1 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 2 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 3 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 0 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 0 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 0 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 1 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 1 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 1 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 2 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 2 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 2 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 3 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 3 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 1 And Request("Status") = 3 And Request("Region") = 3 Then
   Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "Commercial&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 2 CONSUMER
ElseIf Request("Business") = 2 And Request("Status") = 0 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 1 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 2 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 3 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 0 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 0 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 0 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 1 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 1 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 1 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 2 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 2 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 2 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 3 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 3 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 2 And Request("Status") = 3 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Android Mobililty Consumer</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Business = 3 Android Mobililty Consumer
ElseIf Request("Business") = 3 And Request("Status") = 0 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 1 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 2 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 3 And Request("Region") = 0 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=0&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 0 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 0 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 0 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=0'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=0'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=0'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Transitioning</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;All"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 1 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 1 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 1 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=1'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=1'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=1'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "Current&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Transitioning&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 2 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 2 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 2 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=2'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=2'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=2'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=2'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;Transitioning|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>Inactive</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 3 And Request("Region") = 1 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=1&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=1&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "Americas&nbsp;|&nbsp;<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>APJ&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=2&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=2&Region=1&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=1&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 3 And Request("Region") = 2 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=2&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=2&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=2&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=3'>APJ</a>&nbsp;|&nbsp;EMEA&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=2&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=0'>All</a>"
    Response.Write "</td></tr>"
ElseIf Request("Business") = 3 And Request("Status") = 3 And Request("Region") = 3 Then
    Response.Write "<tr><td nowrap><b>Business:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=1&Region=3&Status=3'>Commercial&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=2&Region=3&Status=3'>Consumer&nbsp;</a>&nbsp;|&nbsp;Android Mobililty Consumer&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=0&Region=3&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Region:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=1&Status=3'>Americas&nbsp;</a>&nbsp;|&nbsp;APJ&nbsp;|" & _
    "<a href='ConfigCodes.asp?Business=3&Region=2&Status=3'>EMEA</a>&nbsp;|&nbsp;" &_
    "<a href='ConfigCodes.asp?Business=3&Region=0&Status=3'>All</a>"
    Response.Write "</td></tr>"
    
    Response.Write "<tr><td nowrap><b>Status:</b></td><td width='100%'>"
    Response.Write "<a href='ConfigCodes.asp?Business=3&Region=3&Status=1'>Current&nbsp;</a>&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=2'>Transitioning&nbsp;</a>|&nbsp;" & _
    "Inactive&nbsp;|&nbsp;" & _
    "<a href='ConfigCodes.asp?Business=3&Region=3&Status=0'>All</a>"
    Response.Write "</td></tr>"
End If
%>
</table></td></tr></table>
<br />
<font size=1 color=green>Click on headers to sort.</font><BR><BR>
<table ID="ConfigTable" cellpadding=2 cellspacing=0 STYLE="border-collapse:collapse">
<thead class="th">
<Th onclick="SortTable( 'ConfigTable', 0 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">ID</Th>
<Th onclick="SortTable( 'ConfigTable', 1 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Cons</Th>
<Th onclick="SortTable( 'ConfigTable', 2 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Comm</Th>
<Th onclick="SortTable( 'ConfigTable', 3 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Android Mobility Cons</Th>
<Th onclick="SortTable( 'ConfigTable', 4 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Country Region</Th>
<Th onclick="SortTable( 'ConfigTable', 5 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">HP Code</Th>
<Th onclick="SortTable( 'ConfigTable', 6 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">DASH Code</Th>
<Th onclick="SortTable( 'ConfigTable', 7 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Country Code</Th>
<Th onclick="SortTable( 'ConfigTable', 8 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">GM Code</Th>
<Th onclick="SortTable( 'ConfigTable', 9 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Image Languages</Th>
<Th onclick="SortTable( 'ConfigTable', 10 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">MUI</Th>
<Th onclick="SortTable( 'ConfigTable', 11 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Keyboard</Th>
<Th onclick="SortTable( 'ConfigTable', 12 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">PowerCord</Th>
<Th onclick="SortTable( 'ConfigTable', 13 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Image Docs</Th>
<Th onclick="SortTable( 'ConfigTable', 14 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Printed Docs</Th>
<Th onclick="SortTable( 'ConfigTable', 15 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">Comments</Th>
<!--<Th onclick="SortTable( 'ConfigTable', 14 ,0,1);" onmouseover="HeaderMouseOver();" onmouseout="HeaderMouseOut();">OS Restore Solution</Th>-->

</thead>
<%
	dim strMob
	dim strCons
	dim strEnt
	
	strMob=""
	strCons=""
	strEnt=""

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	'rs.Open "Select r.ID, r.Tablet, r.Consumer, r.Commercial, r.OSLanguage, r.restoremedia, r.DocKits, r.MUI, r.keyboard, r.powercord, r.kwl, r.OtherLanguage, r.CountryCode, r.OptionConfig, r.Dash, r.Name, r.Transitioning, r.PrintedDocs, r.Comments from regions r where r.active=1 order by r.dash;",cn,adOpenForwardOnly
    rs.Open "usp_SelectMasterLocalizationList " & clng(request("Business")) & ","  & clng(request("Region")) & ","  & clng(request("Status")),cn,adOpenForwardOnly

	do while not rs.EOF
		strLanguages =  "<u>" & rs("OSLanguage") & "</u>"
		if trim(rs("OtherLanguage") & "") <> "" then
			strLanguages = strLanguages & "," & rs("OtherLanguage")
		end if
		if trim(rs("KWL") & "") ="" then
			strKWL = "&nbsp;"
		else
			strKWL = rs("KWL")
		end if		

		if rs("AndroidMobilityConsumer") then
			strMob = "X"
		else	
			strMob = "&nbsp;"
		end if

		if rs("Commercial") then
			strEnt = "X"
		else	
			strEnt = "&nbsp;"
		end if

		if rs("Consumer") then
			strCons = "X"
		else	
			strCons = "&nbsp;"
		end if

		Response.Write "<TR class=""td""><TD>" & rs("ID") & "</TD><TD>" & strCons & "&nbsp;</TD><TD>" & strEnt & "&nbsp;</TD><TD>" & strMob & "&nbsp;</TD><TD>" & rs("Name") & "</td><td>" & rs("OptionConfig") & "</td><td>" & rs("Dash") & "</td><td>" & rs("CountryCode") & "</td><td>" & rs("GMCode") & "</td><td>" & strLanguages & "</td><td>" & rs("MUI") & "&nbsp;</td><td>"  & rs("Keyboard") &  "</td><td>"  & rs("PowerCord") & "</td><td>"  & rs("DocKits") & "</td><td>"  & rs("PrintedDocs") & "&nbsp;</td><td>"  & rs("Comments") & "&nbsp;</td></TR>" '<td>"  & rs("RestoreMedia") & "&nbsp;</td>

		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
	set cn = nothing
%>
</table>
</BODY>
</HTML>
