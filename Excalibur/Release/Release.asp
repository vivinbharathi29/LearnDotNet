<%@ Language=VBScript %>
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  
%>

<HTML>
	<TITLE>Add Product Release</TITLE>
<HEAD>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="EditRelease.asp?ID=<%=request("ID")%>&ProductTypeID=<%=request("ProductTypeID")%>&BusinessSegmentID=<%=request("BusinessSegmentID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="Buttons.asp">
	</FRAMESET>
</HTML>