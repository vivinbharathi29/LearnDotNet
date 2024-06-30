<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD><TITLE>Enter Text</TITLE>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function rtrimCRLF( varText )
{
    var i = 0;
    var j = varText.length - 1;
    
//	for( i = 0; i < varText.length; i++ )
//		{
//		if( varText.substr( i, 1 ) != "\r" && varText.substr( i, 1 ) != "\n")
//		break;
//		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != "\r" && varText.substr( j, 1 ) != "\n")
		break;
		}

    if( i <= j )
		return( varText.substr( 0, (j+1)-i ) );
	else
		return("");
}

function window_onload() {
	if (window.dialogArguments.value.length>1)
		{
//		if(window.dialogArguments.value.substr(window.dialogArguments.value.length-1,2)=="\r")
//			txtInput.value= "test";window.dialogArguments.value;
//		else
			txtInput.value= rtrimCRLF(window.dialogArguments.value);		
		}
	else
		txtInput.value= window.dialogArguments.value;
	txtInput.focus();
	//lblCharCount.innerText = txtInput.value.length;
}

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
//	if (txtInput.value.length > 7998)
//		{
//		alert("This field can not contain more than 8000 characters.");
//		}
//	else
//		{
		window.returnValue = txtInput.value;
		window.close();
//		}
}


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<TABLE width=100%><TR><TD width=30>&nbsp;</TD><TD width=100%>
<font size=2 face=verdana><b>Enter PNP Devices</b></font><BR><BR>
<font size=1 face=verdana>- Separate each device ID with a single carriage return.<BR>- The information you enter here is not validated by Excalibur. Please ensure that the information is accurate.</font><BR>
<TABLE WIDTH="100%" bgcolor=Ivory BORDER=0 CELLSPACING=0 CELLPADDING=2 bordercolor=tan style="LEFT: 9px; BORDER-BOTTOM: lightgrey thin solid; TOP: 1px" align=center>
	<TR>
		<TD  width=100% vAlign=top>
		<TEXTAREA style="width: 100%" rows=20 id=txtInput  name=txtInput></TEXTAREA>
		</TD>
	</TR>
	</TABLE>
	<table width="100%" style="BORDER-TOP: white thin solid">
	<TR>
		<TD colspan=2 align=right bgColor=Ivory>
      <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()" style="FONT-FAMILY: "> <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</table>
</TD></TR></TABLE>
</TEXTAREA>
</BODY>
</HTML>
