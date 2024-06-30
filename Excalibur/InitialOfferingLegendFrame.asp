<%@ Language=VBScript %>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<HTML>

<TITLE>Initial Offering Legend</TITLE>

<FRAMESET ROWS="*" ID=TopWindow >
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="InitialOfferingLegend.aspx">
</FRAMESET>
	
</HTML>