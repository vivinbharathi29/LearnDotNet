	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

	%>
<HTML>
<HEAD>
<title>Choose Product Type</title>


</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
    <%if Request("app")="PulsarPlus" then%>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="AddProgramMain.asp?<%=Request.QueryString %>">
        <%else%>
        <FRAME ID="MyWindow" Name="MyWindow" SRC="AddProgramMain.asp<%=Request.QueryString %>">
        <%end if%>
</FRAMESET>
</HTML>
