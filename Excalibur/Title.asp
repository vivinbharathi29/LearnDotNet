<html xmlns:t ="urn:schemas-microsoft-com:time" >
<?IMPORT namespace="t" implementation="#default#time2">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Pulsar</title>
 <link rel="shortcut icon" href="favicon.ico" >
<style>
 .digital {position:absolute; top:110; left:00; color:white; font-family:courier new; font-weight:bold; font-size:13pt; filter:progid:DXImageTransform.Microsoft.Alpha(opacity='10');}
    a {
        color:white;
        text-decoration:none;
    }

    a.IRS {
        color:white;
        text-decoration: underline;
    }
</style>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
    function UpdateUserAccess() {
        window.location.href = "UpdateUserAccess.asp";
    }
    function TitleColor_onclick() {
	var expireDate = new Date();
	expireDate.setMonth(expireDate.getMonth()+12);
	document.cookie = "TitleColor=#0096d6" + ";expires=" + expireDate.toGMTString() + ";";
	//document.cookie = "TitleColor=#0000cd" + ";expires=" + expireDate.toGMTString() + ";";
    /*
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
		*/
		
}



//-->
</script>
</head>
<%
	TitleColor = "#0096d6"
	'TitleColor = "#0000cd"
%>
<body style="margin: 0 0 0 0" LANGUAGE="javascript" onload="TitleColor_onclick()">
<table style="margin:  0 0 0 0" cellspacing=0 cellpadding=0 width="100%" height="100%" bgcolor="<%=TitleColor%>" id="TitleColor"> <!-- LANGUAGE="javascript" onclick="return TitleColor_onclick()">-->
<tr>
<td nowrap><font color="white"><b>
<font> <font face="Trebuchet MS,Verdana" size="3"><font size="6">&nbsp;<a target="_parent" href="Excalibur.asp">Pulsar</a></font>
<!-- <img src="images/3.gif" WIDTH="37" HEIGHT="26" alt=3>-->
<%if Application("Excalibur_ServerName") = "PulsarDev.usa.hp.com" then%>
<font size=2 face=verdana color=white>[ Development ]</font>
<%end if%>
<%if Application("Excalibur_ServerName") = "PulsarTest.usa.hp.com" then%>
<font size=2 face=verdana color=white>[ Test ]</font>
<%end if%>
<%if Application("Excalibur_ServerName") = "PulsarSandbox.usa.hp.com" then%>
<font size=2 face=verdana color=white>[ Sandbox ]</font>
<%end if%>
<%if Application("Excalibur_ServerName") = "PulsarWeb.usa.hp.com" then%>
<font size=2 face=verdana color=white>[ HP Restricted ]</font>
<%end if%>
</font>
</font>
</b></font> </td>
    <td align="right"><div style="display:none"><font face=verdana color=white size=3><a class="IRS" target="_blank" href="http://irsweb.usa.hp.com">Open IRS</a>&nbsp;&nbsp;</div></font><img
            align="middle" src="images/hpblue.gif" />&nbsp;&nbsp;</td>
</tr>
</table>
<div style="display:none"><a href="UpdateUserAccess.asp"></a></div>

</body>
</html>
