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

<frameset rows="*,60" id="TopWindow">
    <frame noresize ID="UpperWindow" Name="UpperWindow" SRC="LocalizeMain_Pulsar.asp?PVID=<%=Request("PVID")%>&strSeriesSummary=<%=Request("strSeriesSummary")%>&BID=<%=Request("BID")%>&CategoryID=<%=Request("CategoryID")%>&ShowAllLocs=<%=Request("ShowAllLocs")%>&Releases=<%=Request("Releases")%>&RTPDate=<%=Request("RTPDate")%>&EMDate=<%=Request("EMDate")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>"/>
	<frame noresize ID="LowerWindow" Name="LowerWindow" SRC="LocalizeAVsButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" />
</frameset>
</HTML>