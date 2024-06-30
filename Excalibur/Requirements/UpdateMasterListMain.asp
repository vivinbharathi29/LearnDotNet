<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>

  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.close();
}

function VerifySave(){
	var i;
	var blnSuccess;
	var strNTName;
	blnSuccess = true;

	if (AddRequirement.txtNamevalue == "")
		{
			window.alert("Requirement Name is required.");
			AddRequirement.txtName.focus();
			blnSuccess = false;
		}

	return blnSuccess;
}

function cmdOK_onclick() {
	if (VerifySave())
		{
		AddRequirement.cmdCancel.disabled = true;
		AddRequirement.cmdOK.disabled = true;
		window.returnValue=1;
		window.AddRequirement.submit();
		}
}

function window_onload() {
	AddRequirement.txtName.focus();
}

function txtName_onkeypress() {
	if (window.event.keyCode == 13)
		cmdOK_onclick();
}

//-->
</SCRIPT>
</head>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">


<%
'	dim cn
'	dim rs
'	dim strName
  
'	set cn = server.CreateObject("ADODB.Connection")
'	cn.ConnectionString = Session("PDPIMS_ConnectionString")
'	cn.Open

'	set rs = server.CreateObject("ADODB.recordset")
  
	strName = ""
	
'	if request("ID") <> "" then
'		rs.Open "spGetEmployeeByID " &  request("ID"),cn,adOpenForwardOnly
'		if not (rs.EOF and rs.BOF) then
'			strName = rs("Name") & ""
'			strPhone = rs("Phone") & ""
'			strEmail = rs("Email") & ""
'			strNTName = rs("NTName") & ""
'			strEmail = rs("Email") & ""
'			strWorkgroupID = rs("WorkgroupID") & ""
'			strDivision = rs("Division") & ""
'			strDomain = rs("Domain") & ""
'		end if
'		rs.Close
'	end if
  
  
	
  


'	set rs = nothing
'	set cn = nothing
	

%>


<font face=verdana>
<FORM ACTION="UpdateMasterSave.asp" METHOD="post" NAME="AddRequirement">

<font size=4 face=verdana>Add New Requirement</font><BR><BR>
<table border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan>
  <tr>
    <td valign=top width="160"><strong><font size=2>Requirement&nbsp;Name:</font><font color=red size=1> *</font></strong></td>
    <td valign=top><INPUT id=txtName name=txtName style="WIDTH: 100%" maxlength=255 value="<%=strName%>" LANGUAGE=javascript onkeypress="return txtName_onkeypress()"></td>
  </tr>
</table>
<table width="100%" border=0>
  <tr><TD align=right>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
  </TD></tr>
</table>
</FORM>
</font>
</body>
</html>