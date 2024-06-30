<%@ Language=VBScript %>
<% Option Explicit %>
<%
	
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
	  
%>
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<%	If Request("ID") <> "" Then %>
<title>Update Schedule Item</title>
</head>
<frameset ROWS="*,70" ID="TopWindow">
	<frame noresize frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="MilestoneMain.asp?ID=<%=Request("ID")%>&action=<%=Request("action")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<frame noresize frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="MilestoneButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
</frameset>
<%	ElseIf Request("PVID") <> "" Then %>
<title>Select Schedule Items</title>
</head>
<frameset ROWS="*,70" ID="TopWindow">
	<frame noresize frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="MilestoneListMain.asp?PVID=<%=Request("PVID")%>&ScheduleID=<%=Request("ScheduleID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<frame noresize frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="MilestoneButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
</frameset>
<%	ElseIf Request("ScheduleID") <> "" Then %>
<title>Add Custom Item</title>
</head>
<frameset ROWS="*" ID="TopWindow">
	<frame noresize frameborder="0" ID="MainWindow" Name="MainWindow" SRC="../Schedule/MilestoneInsert.asp?ProdVID=<%=Request("ProdVID")%>&ScheduleID=<%=Request("ScheduleID")%>">
</frameset>
<%	Else %>
<body>
<font size="2" face="verdana"><br>Unable to display the requested page because not enough information was supplied.</font>
</body>
<%	End If %>
</html>