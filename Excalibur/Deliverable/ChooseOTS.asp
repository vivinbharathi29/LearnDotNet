<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Choose Observations Fixed</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChooseOTSMain.asp?UserID=<%=Request("UserID")%>&VersionID=<%=Request("VersionID")%>&ID=<%=Request("ID")%>&OldIDList=<%=request("OldIDList")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ChooseOTSButtons.asp">
</FRAMESET>

</HTML>