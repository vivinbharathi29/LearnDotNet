<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Share Profile</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProfileShareMain.asp?ID=<%=Request("ID")%>">
<!--	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProfileShareButtons.asp">-->
</FRAMESET>

</HTML>