<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  
  Dim AvType : AvType = Request("AvType")
  Dim UserName : UserName = Request("UserName")
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<TITLE>Select AVs To Localize</TITLE>
<HEAD>

</HEAD>

<FRAMESET ROWS="*,60" ID=TopWindow >
  <%If AvType = 1  Then%>
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="LocalizeAVsMain.asp?PVID=<%=Request("PVID")%>&strSeriesSummary=<%=Request("strSeriesSummary")%>&BID=<%=Request("BID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
  <%ElseIf AvType = 2  Then%>	
  	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="LocalizeHWKitsMain.asp?PVID=<%=Request("PVID")%>&strSeriesSummary=<%=Request("strSeriesSummary")%>&BID=<%=Request("BID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
  <%ElseIf AvType = 3  Then%>	
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="LocalizeKeyboardsMain.asp?PVID=<%=Request("PVID")%>&strSeriesSummary=<%=Request("strSeriesSummary")%>&BID=<%=Request("BID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
  <%ElseIf AvType = 4  Then%>	
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="LocalizeOSRestoreMediaMain.asp?PVID=<%=Request("PVID")%>&strSeriesSummary=<%=Request("strSeriesSummary")%>&BID=<%=Request("BID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
  <%End If %>
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="LocalizeAVsButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>
</HTML>