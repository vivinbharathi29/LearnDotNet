<%@ Language=VBScript %>

<%Response.Expires = 0%>

<HTML>
<HEAD><TITLE>Add Device</TITLE>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

  function isHex(value) 
  {
    if (value == "")  { return false }
    if (value.charAt(0) == "-")
      start = 1;
    else
      start = 0;
    
    for (i=start; i<value.length; i++)
    {
      if (!((value.charAt(i) >= "0" && value.charAt(i) <= "9" ) || (value.charAt(i) >= "a" && value.charAt(i) <= "f") || (value.charAt(i) >= "A" && value.charAt(i) <= "F"))) 
		{
			return false;		
		}
    }
    return true;
  }


function window_onload() {
	var s = window.dialogArguments;
	if (s.substr(0,7) == "USB\\VID" )
		{
		cboType.selectedIndex = 1;
		TypeText.innerText =  "VID:";
		DevText.innerText =  "PID:";
		VenRow.style.display="";
		SubSysRow.style.display="";
		DevRow.style.display="";
		OtherRow.style.display="none";
		}
	else if (s.substr(0,7) == "PCI\\VEN" || s=="")
		{
		cboType.selectedIndex = 0;
		TypeText.innerText =  "VEN:";
		DevText.innerText =  "DEV:";
		VenRow.style.display="";
		SubSysRow.style.display="";
		DevRow.style.display="";
		OtherRow.style.display="none";
		}
	else 
		{
		cboType.selectedIndex = 2;
		VenRow.style.display="none";
		SubSysRow.style.display="none";
		DevRow.style.display="none";
		OtherRow.style.display="";
		
		}

	if (cboType.selectedIndex == 2)
		{
		if (s.indexOf("=",0) > -1)
			{
			txtOther.value = s.substring(0,s.indexOf("=",0));
			txtDescription.value = s.substr(s.indexOf("=",0)+2)
			if (txtDescription.value.substring(0,1)=='"')
				txtDescription.value = txtDescription.value.substring(1);
			if (txtDescription.value.substr(txtDescription.value.length-1,1)=='"')
				txtDescription.value = txtDescription.value.substring(txtDescription.value.length-1,0);
			}
		else
			{
			txtOther.value = s;
			}
		
		}
	else
		{		
		txtPCI.value = s.substr(8,4); 
		txtDEV.value = s.substr(17,4); 
		if (s.substr(21,1)=="&")
			{
			txtSUBSYS.value = s.substr(29,8); 
			txtDescription.value = s.substr(38); 
			}
		else
			{
			txtSUBSYS.value = "";
			txtDescription.value = s.substr(23); 
			}
		if (txtDescription.value.substring(0,1)=='"')
			txtDescription.value = txtDescription.value.substring(1);
		if (txtDescription.value.substr(txtDescription.value.length-1,1)=='"')
			txtDescription.value = txtDescription.value.substring(txtDescription.value.length-1,0);

		}
}

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
	var Result;
	var Buffer;
	var ValidationFailed=false;
	
	
	//Validation
	if (cboType.selectedIndex == 2)
		{
		if (txtOther.value == "")
			{
			window.alert("ID field is required.");
			ValidationFailed = true;
			txtOther.focus();
			}
		if (txtDescription.value == "")
			{
			window.alert("Description field is required.");
			ValidationFailed = true;
			txtDescription.focus();
			}
		}
	else
		{
		Buffer = txtPCI.value;
		if (Buffer.length < 4)
			{
				window.alert(TypeText.innerText.replace(":","") + " must be 4 characters long.");
				ValidationFailed = true;
				txtPCI.focus();
			}
			
		if (!ValidationFailed)
			{
				if (!isHex(Buffer))
					{
					window.alert(TypeText.innerText.replace(":","") + " field must contain only Hex characters (0-9 and A-F).");
					ValidationFailed = true;
					txtPCI.focus();
					}
			}	


		Buffer = txtDEV.value;
		if (!ValidationFailed && Buffer.length < 4)
			{
				window.alert(DevText.innerText.replace(":","") + " field must be 4 characters long.");
				ValidationFailed = true;
				txtDEV.focus();
			}


		if (!ValidationFailed)
			{
				if (!isHex(Buffer))
					{
					window.alert(DevText.innerText.replace(":","") + " field must contain only Hex characters (0-9 and A-F).");
					ValidationFailed = true;
					txtDEV.focus();
					}
			}	


		Buffer = txtSUBSYS.value;
		if (!ValidationFailed && Buffer.length < 8 && Buffer.length != 0)
			{
				window.alert("SUBSYS field must be 8 characters long if it is supplied.");
				ValidationFailed = true;
				txtSUBSYS.focus();
			}

		if (!ValidationFailed && Buffer.length != 0)
			{
				if (!isHex(Buffer))
					{
					window.alert("SUBSYS field must contain only Hex characters (0-9 and A-F).");
					ValidationFailed = true;
					txtSUBSYS.focus();
					}
			}	
	
		Buffer = txtDescription.value;
		if (!ValidationFailed && Buffer.length == 0)
			{
				window.alert("Description field is required.");
				ValidationFailed = true;
				txtDescription.focus();
			}	
		}
	
	//Save
	if (!ValidationFailed)
		{
		if (cboType.selectedIndex == 0 )
			Result= 'PCI\\VEN_' + txtPCI.value + '&DEV_' + txtDEV.value;
		else if (cboType.selectedIndex == 1 )
			Result= 'USB\\VID_' + txtPCI.value + '&PID_' + txtDEV.value;
		else 
			Result= txtOther.value;
		
		if (txtSUBSYS.value != '' && cboType.selectedIndex < 2)
			Result = Result + '&SUBSYS_' + txtSUBSYS.value;
		Result = Result + '="' + txtDescription.value + '"';
			
		window.returnValue = Result;
		window.close();
		}
}

function txtPCI_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
		
}

function txtDEV_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
}

function txtSUBSYS_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
}

function txtDescription_onkeypress() {
	if (window.event.keyCode==13)
		{
			cmdOK_onclick();
		}
}

function cboType_onchange() {
	if (cboType.selectedIndex ==0 )
		{
		TypeText.innerText =  "VEN:";
		DevText.innerText =  "DEV:";
		VenRow.style.display="";
		SubSysRow.style.display="";
		DevRow.style.display="";
		OtherRow.style.display="none";
		}
	else if (cboType.selectedIndex ==1 )
		{
		TypeText.innerText =  "VID:";
		DevText.innerText =  "PID:";
		VenRow.style.display="";
		SubSysRow.style.display="";
		DevRow.style.display="";
		OtherRow.style.display="none";
		}
	else
		{
		VenRow.style.display="none";
		SubSysRow.style.display="none";
		DevRow.style.display="none";
		OtherRow.style.display="";
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<TABLE WIDTH="100%" bgcolor=Ivory BORDER=0 CELLSPACING=0 CELLPADDING=2 bordercolor=tan style="LEFT: 9px; BORDER-BOTTOM: lightgrey thin solid; TOP: 1px" align=center>
	<TR>
		<TD nowrap width="10%" vAlign=top><FONT size=1 face=Verdana>
		&nbsp;Type:</FONT> <FONT color=#ff0000 size=1>*</FONT> </TD>
		<TD vAlign=top>
		<SELECT id=cboType name=cboType style="WIDTH: 62px;" LANGUAGE=javascript onchange="return cboType_onchange()">			<OPTION>PCI</OPTION>
			<OPTION>USB</OPTION>
			<OPTION>Other</OPTION>
		</SELECT>&nbsp;</TD>
    </TR>
	<TR ID=VenRow>
		<TD nowrap width="10%" vAlign=top><FONT size=1 face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT><label ID=TypeText>PCI\VEN:</label></FONT> <FONT color=#ff0000 size=1>*</FONT> </TD>
		<TD  vAlign=top><INPUT id=txtPCI name=txtPCI style="WIDTH: 62px; HEIGHT: 22px" size=8 LANGUAGE=javascript onkeypress="return txtPCI_onkeypress()" maxLength=4>&nbsp;<FONT color=green size=1 
      face="MS Sans Serif">(4-digit hex number)</FONT></TD>
    </TR>
	<TR ID=DevRow>
		<TD nowrap width="10%" vAlign=top><FONT size=1 face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT><label ID=DevText>DEV:</label></FONT> <FONT color=#ff0000 size=1>*</FONT> </TD>
		<TD  vAlign=top><INPUT id=txtDEV name=txtDEV style="WIDTH: 63px; HEIGHT: 22px" size=8 LANGUAGE=javascript onkeypress="return txtDEV_onkeypress()" maxLength=4>&nbsp;<FONT color=green size=1 
      face="MS Sans Serif">(4-digit hex number)</FONT></TD>
    </TR>
	<TR  ID=SubSysRow>
		<TD nowrap width="10%" vAlign=middle><FONT size=1 face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT>SUBSYS:</FONT></TD>
		<TD  vAlign=top><INPUT id=txtSUBSYS name=txtSUBSYS style="WIDTH: 100px; HEIGHT: 22px" LANGUAGE=javascript onkeypress="return txtSUBSYS_onkeypress()" maxLength=8>&nbsp;<FONT color=green size=1 
      face="MS Sans Serif">(8-digit hex number)<br /><table cellspacing=0 cellpadding=0 style="margin-top: 3px;"><tr><td style="border-top:solid 1px darkgray"><font size=1 color=darkgray>SubDevID</font></td><td width="3px">&nbsp;</td><td style="border-top:solid 1px darkgray"><font size=1 color=darkgray>SubVenID</font></td></tr></table></font></FONT></TD>
    </TR>
	<TR ID=OtherRow>
		<TD nowrap width="10%" vAlign=middle><FONT size=1 face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT>ID:</FONT> <FONT color=#ff0000 size=1>*</FONT></TD>
		<TD  vAlign=top><FONT color=green size=1 face="MS Sans Serif">Note: Excalibur does not check the format or value of this field.</FONT>
		<INPUT id=txtOther name=txtOther style="WIDTH: 100%; HEIGHT: 22px" size=17 LANGUAGE=javascript onkeypress="return txtSUBSYS_onkeypress()" maxLength=50></TD>
    </TR>
	<TR>
		<TD width="10%" nowrap  vAlign=top><FONT size=1 
     face=Verdana><FONT face="Times New Roman" 
      size=3>&nbsp;</FONT>HW Marketing 
      Desc:</FONT> <FONT color=#ff0000 size=1>*</FONT></TD>
		<TD  vAlign=top><INPUT id=txtDescription 
      style="WIDTH: 100%; HEIGHT: 22px" size=72 name=txtDescription LANGUAGE=javascript onkeypress="return txtDescription_onkeypress()" maxLength=255></TD>
	</TR>
	</TABLE>
	<table width="100%" style="BORDER-TOP: white thin solid">
	<TR>
		<TD colspan=2 align=right bgColor=Ivory>
      <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()" style="FONT-FAMILY: "> <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>

</BODY>
</HTML>
