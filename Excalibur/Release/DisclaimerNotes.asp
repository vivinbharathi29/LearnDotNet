<%@ Language=VBScript %>
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  
%>

<HTML>
	<TITLE>Cycle Disclaimer Notes</TITLE>
<HEAD>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="EditDisclaimerNotes.asp?ReleaseID=<%=request("ReleaseID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="Buttons.asp">
	</FRAMESET>
</HTML>