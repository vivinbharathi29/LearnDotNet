<%@ Language=VBScript %>


<html>
<head>

  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.close();
    }
}

function VerifySave(){
	var i;
	var blnSuccess;
	var strFamilyName;
	blnSuccess = true;

	if (AddFamily.txtName.value == "")
		{
			window.alert("Family Name is required.");
			AddFamily.txtName.focus();
			blnSuccess = false;
		}
	else if (trim(AddFamily.txtName.value).indexOf(" ") > -1 )
		{
			window.alert("Spaces are not allowed in Product Family names.");
			AddFamily.txtName.focus();
			blnSuccess = false;
		}
		
	//Look for dup Family Names
	else
		{	
			strFamilyName = AddFamily.txtName.value;
			for (i=0;i<AddFamily.lstFamilies.length;i++)
				{
				if (AddFamily.lstFamilies.options[i].text == strFamilyName.toLowerCase())
					{
					window.alert("That Family Name already exists.");
					AddFamily.txtName.focus();
					blnSuccess = false;
					}	
				}
		}
	return blnSuccess;
}
		
function trim( varText )
{
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
}


function cmdOK_onclick() {
	if (VerifySave())
		{
		AddFamily.cmdOK.disabled = true;
		AddFamily.cmdCancel.disabled = true;

		window.returnValue=1;
		window.AddFamily.submit();
		}
}

function window_onload() {
	AddFamily.txtName.focus();
}

//-->
</SCRIPT>
</head>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">


<%
	dim cn
	dim rs
	dim cnString
	dim strFamilyNames
  
	cnString = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
  
	strFamilyNames = ""
	rs.Open "spGetProductFamiliesAll",cn,adOpenForwardOnly
	do while not rs.EOF
		strFamilyNames = strFamilyNames & "<OPTION>" & lcase(rs("Name") & "" ) & "</OPTION>"
		rs.MoveNext
	loop

	rs.Close

	set rs = nothing
	set cn = nothing
	


%>


<font face=verdana>
<FORM ACTION="familysave.asp" METHOD="post" NAME="AddFamily">


<H3>Add New Product Family</H3>
<table border="1" cellPadding="2" cellSpacing="0" width="500" bgcolor=cornsilk bordercolor=tan>
  <tr>
    <td valign=top width="160"><strong><font size=2>Family 
      Name:</font><font color=red size=1> *</font></strong></td>
    <td valign="top"><input id="txtName" name="txtName" style="width: 340px; height: 22px" size="6" maxlength="64"></td>
  </tr>
</table>
<table width="400" border=0>
  <tr>
     <td align=right>
        <SELECT style="Display:none" size=2 id=lstFamilies name=lstFamilies><%=strFamilyNames%></SELECT>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
  </td></tr>
</table>
</FORM>
</font>
</body>
</html>
