<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<TITLE>Deliverable Features</TITLE>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChooseFeaturesMain.asp?RootID=<%=request("RootID")%>&IDList=<%=request("IDList")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ChooseFeaturesButtons.asp">
</FRAMESET>

</HTML>