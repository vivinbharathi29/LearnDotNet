<%@ Language=VBScript %>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<HTML>

<TITLE>AVs Missing Deliverable Root</TITLE>

<FRAMESET ROWS="*" ID=TopWindow >
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="AvsMissingDelRoot.aspx?<%=Request.QueryString %>">
</FRAMESET>
	
</HTML>