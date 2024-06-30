<%@  language="VBScript" %>
<% option explicit  %>
<html>
<head>
	<title>Observation Reports - Offline</title>
	<script type="text/javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
	<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
	<script type="text/javascript" src="../../Includes/Client/Common.js"></script>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--



//-->
	</script>
	<style type="text/css">
		TEXTAREA
		{
			font-weight: normal;
			font-size: x-small;
			font-family: Verdana;
		}
		A:link
		{
			color: blue;
		}
		A:visited
		{
			color: blue;
		}
		A:hover
		{
			color: red;
		}
		TD.HeaderButton
		{
			font-size: 8pt;
			font-family: Verdana;
			font-weight: bold;
			color: White;
			padding: 3px;
		}
		body
		{
			background-color: lightsteelblue;
			font-size: 10pt;
			font-family: Verdana;
		}
		td
		{
			font-size: 10pt;
			font-family: Verdana;
		}
	</style>
</head>
<body>
	<%
    dim cnExcalibur, rs
	set cnExcalibur = server.CreateObject("ADODB.Connection")
	cnExcalibur.ConnectionString = Session("PDPIMS_ConnectionString")
	cnExcalibur.Open
	set rs = server.CreateObject("ADODB.recordset")

	dim strNoticeTable,BulletinHeaderColors,BulletinBodyColors

	strNoticeTable = ""
	BulletinHeaderColors = split("SeaGreen,Firebrick,Gold,SeaGreen",",")
	BulletinBodyColors = split("Honeydew,MistyRose,LightYellow,Honeydew",",")
	rs.Open "Select * FROM Bulletins with (NOLOCK) where active=1 and OTS=1 Order By id;",cnExcalibur,adOpenForwardOnly
	do while not rs.eof
		strNoticeTable = strNoticeTable & "<table cellSpacing=0 cellPadding=2 width=""100%"" border=0 bordercolor=black>"
			if rs("Severity") > -1 and rs("Severity") < 4 then
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinHeaderColors(rs("Severity")) & ";font-family:verdana;font-size:x-small;font-weight:bold;color:white"">" &  rs("Subject") & "</TD></TR>"
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinBodyColors(rs("Severity"))& ";font-family:verdana;font-size:xx-small;color:black"">" & rs("Body") & "<br/><br/></TD></TR>"
			else
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinHeaderColors(0) & ";font-family:verdana;font-size:x-small;font-weight:bold;color:white"">" &  rs("Subject") & "</TD></TR>"
				strNoticeTable = strNoticeTable &  "<TR><TD style=""background-color:" & BulletinBodyColors(0)& ";font-family:verdana;font-size:xx-small;color:black"">" & rs("Body") & "<br/><br/></TD></TR>"
			end if
		strNoticeTable = strNoticeTable &  "</TABLE><br/>"
		rs.movenext
	loop
	rs.close
    set rs = nothing
    cnExcalibur.close
    set cnExcalibur = nothing



%>

   <h2>Unable to connect to Sudden Impact</h2><br />
   <%=strNoticeTable%>
</body>
</html>
