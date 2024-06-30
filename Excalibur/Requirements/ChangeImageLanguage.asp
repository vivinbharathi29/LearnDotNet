<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->
	
<HTML>
<HEAD>
<TITLE>Select Languages</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {
	var strValue="";
	var i;

	for (i = 0;i<chkLang.length;i++)
		if (chkLang(i).checked)
			strValue = strValue + "," + chkLang(i).value;
	if (strValue.length > 0)
		strValue = strValue.substr(1);
		
	if (strValue == "" )
		{
		alert("You must select at least one language.");
		document.focus();
		}
	else	
		{
		window.returnValue = strValue;
		window.close();
		}
}

function window_onload() {
//	cmdOK.focus();
}



function cmdCancel_onclick() {
		window.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">



<table width=100%><tr><td width=20>&nbsp;</td><TD>
<font size=3 face =verdana><b><BR>Select Languages</b></font><BR><BR>


<Table bordercolor=tan cellpadding=1 cellspacing=0 border=1 width=100%><TR bgcolor=wheat><TD><font size=2 face=verdana><b>Languages In Image</b></font></td></tr>

<TR bgcolor=cornsilk><TD><font size=2 face=verdana>
<%

	dim LangArray
	dim i
	
	LangArray = split(request("AllLangs"),",")
	
	for i = lbound(LangArray) to ubound(LangArray)
		if instr("," & request("SelectedLangs") & "," , "," & LangArray(i) & ",") > 0 then
			Response.Write "<INPUT  hidefocus checked type=""checkbox"" id=chkLang name=chkLang value=""" & LangArray(i) & """>" & LangArray(i) & "<BR>"	
		else
			Response.Write "<INPUT  hidefocus type=""checkbox"" id=chkLang name=chkLang value=""" & LangArray(i) & """>" & LangArray(i) & "<BR>"	
		end if
	next


%>

</font></TD></TR></TABLE>


</td></tr>
<TR><TD colspan=2 align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr>
</table>
</BODY>
</HTML>
