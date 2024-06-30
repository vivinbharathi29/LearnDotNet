<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Supplier Code</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="SupplierCodeMain.asp?CategoryID=<%=request("CategoryID")%>&VendorID=<%=request("VendorID")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="SupplierCodeButtons.asp" scrolling=no>
</FRAMESET>

</HTML>