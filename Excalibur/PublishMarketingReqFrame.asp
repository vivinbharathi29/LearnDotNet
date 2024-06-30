<%@ Language=VBScript %>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<HTML>

<TITLE>Publish Marketing Requirements</TITLE>

<FRAMESET ROWS="*" ID=TopWindow >
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PublishMarketingReq.aspx?<%=Request.QueryString %>">
</FRAMESET>
	
</HTML>