<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" Content="VBScript"><link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<title>Error 401.2</title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
H1
{
    COLOR: #006697;
    FONT-FAMILY: Tahoma
}
H2
{
    COLOR: #006697;
    FONT-FAMILY: Tahoma
}
BODY
{
    FONT-SIZE: x-small;
    COLOR: black;
    FONT-FAMILY: Verdana, Tahoma, Arial
}
A:hover
{
    FONT-WEIGHT: bolder;
    COLOR: red
}
TABLE
{
    FONT-SIZE: x-small
}
</STYLE>
</head>
<body>
<h2>HTTP Error 401</h2>

<p><strong>401.2 Unauthorized: Logon Failed due to server configuration</strong></p>

<p>This error indicates that the credentials passed to the server do not match the credentials required to log on to the server. This is usually caused by not sending the proper WWW-Authenticate header field.</p>

<p>To fix this problem, make sure the &quot;Bypass proxy server for local addresses&quot; option is checked in your Internet Options.</p>

<p><STRONG>Windows 2000 Users</STRONG></p>

<p>If you are connecting through a RAS connection you need to make this setting for each of your RAS connections.</p>

<p align="center"><img SRC="http://<%=Application("Excalibur_ServerName")%>/errors/autoconfig.gif"></p>

<p align="center"><img SRC="http://<%Application("Excalibur_ServerName")%>/errors/401_2.gif" WIDTH="378" HEIGHT="116"></p>

<p>Please contact the Web server's administrator to verify that you have permission to access to requested resource.</p>
</body>
</html>
