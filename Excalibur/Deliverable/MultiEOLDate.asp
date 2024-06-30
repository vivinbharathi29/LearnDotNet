<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<TITLE>Update Deliverable EOA Information</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MultiEOLDateMain.asp?TypeID=<%=Request("TypeID")%>&IDList=<%=Request("IDList")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MultiEOLDateButtons.asp">
</FRAMESET>

</HTML>