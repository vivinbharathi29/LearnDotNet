<%@ Language=VBScript %>
<%Option Explicit%>
<%
Dim startTime, finishTime, duration
startTime = Now()

%>

<HTML>
<HEAD>
<Title>Agency Matrix Test</Title>
<!-- #include file="AgencyPivot.asp" -->

</HEAD>
<BODY>
<h3>PM View</h3>
<table width=100%  border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>
<%
'Call DrawPMViewMatrix(356)
%>
</table>
<BR><BR><h3>DM View</h3>
<table width=100%  border=1 bordercolor=tan cellpadding=2 cellspacing=1 bgColor=ivory>
<%
'Call DrawDMViewMatrix(5127,"","","")
Call DrawDMViewMatrix(5131)

finishTime = Now()

duration = DateDiff("s", startTime, finishTime)

%>
</table>
<BR><BR>
<p><font size=1>Rendered in <%=duration%> seconds</font></p>
<p><font size=1>Generated on <%=formatdatetime(now)%></font></p>
</BODY>
</HTML>
