<%
	'if Request.Cookies("TitleColor") = "" then
	'	Response.Cookies("TitleColor") = "#0000cd"
	'end if 
%>

<html xmlns:t ="urn:schemas-microsoft-com:time" >
<?IMPORT namespace="t" implementation="#default#time2">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Test Management Tool</title>
<style>
 .digital {position:absolute; top:110; left:00; color:white; font-family:courier new; font-weight:bold; font-size:13pt; filter:progid:DXImageTransform.Microsoft.Alpha(opacity='10');}
</style>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
//var oPopup = window.createPopup();

function TitleColor_onclick() {
	var expireDate = new Date();
	expireDate.setMonth(expireDate.getMonth()+12);
	if (TitleColor.bgColor == "#0000cd")
		{
		TitleColor.bgColor = "#b22222";
		document.cookie = "TitleColor=#b22222" + ";expires=" + expireDate.toGMTString() + ";";
		}
	else if (TitleColor.bgColor == "#b22222")
		{
		TitleColor.bgColor = "#2e8b57";
		document.cookie = "TitleColor=#2e8b57" + ";expires=" + expireDate.toGMTString() + ";";
		}
	else if (TitleColor.bgColor == "#2e8b57")
		{
		TitleColor.bgColor = "#000000";
		document.cookie = "TitleColor=#000000" + ";expires=" + expireDate.toGMTString() + ";";
		}
	else
		{
		TitleColor.bgColor = "#0000cd";
		document.cookie = "TitleColor=#0000cd" + ";expires=" + expireDate.toGMTString() + ";";
		}
		
		
}


function Dale_onclick() {
    var lefter = event.clientX;
    var topper = event.clientY;

}

//-->
</script>
</head>
<%
	dim TitleColor
	on error resume next
	TitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		TitleColor = Request.Cookies("TitleColor")
	else
		TitleColor = "#0000cd"
	end if
	on error goto 0
	'Response.Write "<font color=black>->" & TitleColor & "</font>"
	'if TitleColor & "" = "" then
		'TitleColor = "#0000cd"'"#2e8b57"
		
	'end if
%>
<body topmargin="0" leftMargin="0" rightMargin="0" bottomMargin="0" LANGUAGE="javascript">
<% 
Dim sUser
sUser = LCase(Session("LoggedInUser"))
If false then 
%>
<div id="oBox" style="position:absolute; overflow:hidden; top:0; left:0; width:100%; height:50; filter: progid:DXImageTransform.Microsoft.Gradient(startColorstr='skyBlue',endColorstr='DarkBlue');">
<!-- Three lines of binary -->
<div id="glyphs1" class="digital" style="top:5">
0101110011010010100111001010011011100101010110101100011101010010010111001101001010011100101001101110010101011010110001110101001001011100110100101001110010100110111001010101101011000111010100100101110011010010100111001010011011100101010110101100011101010010</div>   
<div id="glyphs2" class="digital" style="top:13; font-size:28pt">
0101110011010010100111001010111001101001010011100101011100110100101001110010101110011010010100111001</div>   
<div id="glyphs3" class="digital" style="top:30">
0101110011010010100111001010011011100101010110101100011101010010010111001101001010011100101001101110010101011010110001110101001001011100110100101001110010100110111001010101101011000111010100100101110011010010100111001010011011100101010110101100011101010010</div>

<!-- Foreground text -->
</div>
<div style="position:absolute; top:5; left:25; width:400; height:50; font-family:Verdana; font-size:x-large; color:white; font-style:italic; font-weight:bold">Test Management Tool <img ID=Dale src="images/3.gif" WIDTH="37" HEIGHT="26" alt=3 LANGUAGE=javascript onclick="return Dale_onclick()">
<%if Application("On_Live_Server") = 0 then%>
<font size=2 face=verdana color=white>[ Sandbox ]</font>
<%end if%>

</div>

<!-- Animation of binary strings -->
<t:animate 
	targetElement="glyphs2" 
	attributeName="left" 
	begin="0"
	from="00"
	to="-500"
	dur="30"
	repeatcount="indefinite"
/>
<t:animate 
	targetElement="glyphs3" 
	attributeName="left" 
	begin="0"
	from="0"
	to="-500"
	dur="15"
	repeatcount="indefinite"
/>
<% Else %>
<table width="100%" height="100%" bgcolor="<%=TitleColor%>" id="TitleColor" LANGUAGE="javascript" onclick="return TitleColor_onclick()">
  <tr>
    <td nowrap><font color="white"><b>
      <font> <font face="Verdana" size="5"><font size="6">Test Management Tool</font>
      <img ID=Dale src="images/3.gif" WIDTH="37" HEIGHT="26" alt=3 LANGUAGE=javascript onclick="return Dale_onclick()">
      <%if Application("On_Live_Server") = 0 then%>
        <font size=2 face=verdana color=white>[ Sandbox ]</font>
      <%end if%>

      </font>
      </font>
      </b></font>
    </td>
  </tr>
</table>
<%end if %>
<!--<TABLE id=Warning width="100%" bgcolor=red  LANGUAGE=javascript onclick="return TitleColor_onclick()"><TR><TD nowrap valign=top><font face=verdana color=white size=1><b>WARNING: This is       my development copy. If you get a script error, try again later.</FONT>                </b></font> </TD></TR></TABLE>-->
</body>
</html>
