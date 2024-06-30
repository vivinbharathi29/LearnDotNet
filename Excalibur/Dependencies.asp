<%@ Language=VBScript %>

<%Response.Expires = 0%>

<HTML>
<HEAD><TITLE>Add Device</TITLE>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function window_onload() {
}

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
	var Result;
	var Buffer;

	Result= cboRoot.options[cboRoot.selectedIndex].value + ":" + cboRoot.options[cboRoot.selectedIndex].text;
	window.returnValue = Result;
	window.close();
}

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=event.srcElement.length-1;i>=0;i--)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}

function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<%
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
%>



<font face=verdana size=2><BR>&nbsp;&nbsp;Dependent&nbsp;deliverable:<BR>
&nbsp;&nbsp;<SELECT style="WIDTH:95%" id=cboRoot name=cboRoot LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
<OPTION selected></OPTION>
<%
	rs.Open "spGetDelRoot",cn,adOpenForwardOnly
	do while not rs.EOF
		response.write "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
		rs.Movenext
	loop
	rs.Close
%>
</SELECT>

</font>
	<table width="100%" style="BORDER-TOP: white thin solid">
	<TR>
		<TD colspan=2 align=right bgColor=Ivory><HR>
      <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()" style="FONT-FAMILY: "> <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>
<%
	cn.Close
	set rs = nothing
	set cn = nothing
%>
</BODY>
</HTML>
