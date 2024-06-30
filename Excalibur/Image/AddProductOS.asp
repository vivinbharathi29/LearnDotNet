<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

	%>


<HTML>
<TITLE>Add Product OS for Preinstall</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="AddProductOSMain.asp?ProductID=<%=request("ProductID")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="AddProductOSButtons.asp">
</FRAMESET>

</HTML>