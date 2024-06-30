<%@ Language=VBScript %>


	<%
	

  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<title>Update Branding</title>
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="BrandUpdateMain_Pulsar.asp?ProductID=<%=request("ProductID")%>&BrandID=<%=request("BrandID")%>&ExcludeIDList=<%=request("ExcludeIDList")%>&BSegmentID=<%=request("BSegmentID")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="brandupdatebuttons.asp">
</FRAMESET>
</HTML>
