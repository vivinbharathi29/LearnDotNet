<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<DOCTYPE html>
<HTML>
<HEAD>
<TITLE>Select Brands</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {
	var strValue="";
	var i;

	for (i = 0;i<chkBrand.length;i++)
		if (chkBrand(i).checked)
			strValue = strValue + "," + chkBrand(i).value;
	if (strValue.length > 0)
		strValue = strValue.substr(1);
		
	if (strValue == "" )
		{
		alert("You must select at least one brand.");
		document.focus();
		}
	else	
	    {
	        if (window.location != window.parent.location) {
	            parent.PickBrandResult(strValue);
	            parent.modalDialog.cancel();
	        } else {
	            window.returnValue = strValue;
	            window.close();
	        }
		}
}



function cmdCancel_onclick() {
    if (window.location != window.parent.location) {
        parent.modalDialog.cancel();
    } else {
        window.close();
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory>



<table width=100%><tr><td width=20>&nbsp;</td><TD>
<font size=3 face =verdana><b><BR>Select Brands</b></font><BR><BR>

<font size=1 face=verdana color=red>The selected brands do no apply to CTO-only images.</font><BR><BR>
<Table bordercolor=tan cellpadding=1 cellspacing=0 border=1 width=100%><TR bgcolor=wheat><TD><font size=2 face=verdana><b>Brands</b></font></td></tr>

<TR bgcolor=cornsilk><TD><font size=2 face=verdana>

<%

	dim BrandArray
	dim strBrand
	
	BrandArray = split(request("AllBrands"),",")
	for each strBrand in BrandArray
		if trim(request("SelectedBrands")) = "" or instr("," & request("SelectedBrands") & ",", "," & trim(strBrand) & ",") > 0 then
			Response.Write "<INPUT hidefocus checked type=""checkbox"" id=chkBrand name=chkBrand value=""" & strBrand & """>"
		else
			Response.Write "<INPUT hidefocus type=""checkbox"" id=chkBrand name=chkBrand value=""" & strBrand & """>"
		end if
		response.write strBrand
		response.write "<BR>"
	next


%>
</font></TD></TR></TABLE>


</td></tr>
<TR><TD colspan=2 align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr>
</table>
</BODY>
</HTML>
