<%@ Language=VBScript %>

<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>



<HTML>
<HEAD>
<TITLE>Rev Images</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="RevImagesMain.asp?ProductID=<%=Request("ProductID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>

</HTML>
