<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {

	var expireDate = new Date();
	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "Favorites=2900,2909,;expires=" + expireDate.toGMTString() + ";";
	document.cookie = "FavCount=2;expires=" + expireDate.toGMTString() + ";";
	
}

function button2_onclick() {
	var expireDate = new Date();
	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "Favorites=2900,2909,2910,;expires=" + expireDate.toGMTString() + ";";
	document.cookie = "FavCount=3;expires=" + expireDate.toGMTString() + ";";
}

function button3_onclick() {
	var expireDate = new Date();
	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "Favorites=;expires=" + expireDate.toGMTString() + ";";
	document.cookie = "FavCount=0;expires=" + expireDate.toGMTString() + ";";

}

//-->
</SCRIPT>
</HEAD>
<BODY>

<%if request("Type") = "" then %>
<font face=verdana><b>This function is not implemented yet.</b></font>
<%else%>
<FONT face=verdana>
<H3>Deliverable Confirmation Screen<BR>
<HR>
</H3>
Not implemented yet
</font>
<%end if%>
<INPUT style="Display:none" type="button" value="Show 2" id=button1 name=button1 LANGUAGE=javascript onclick="return button1_onclick()">
<INPUT style="Display:none" type="button" value="Show 3" id=button2 name=button2 LANGUAGE=javascript onclick="return button2_onclick()">
<INPUT style="Display:none" type="button" value="Clear" id=button3 name=button3 LANGUAGE=javascript onclick="return button3_onclick()">

</BODY>
</HTML>
