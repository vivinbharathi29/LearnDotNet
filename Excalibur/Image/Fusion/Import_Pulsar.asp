<%@ Language=VBScript %>

<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>



<HTML>
<HEAD>
<TITLE>Import Images</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
	<FRAMESET ROWS="400,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ImportMain_Pulsar.asp?ProductID=<%=Request("ProdID")%>&ProductReleaseID=<%=Request("ProductReleaseID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="lowerWindow" Name="LowerWindow" SRC="ImportButtons_Pulsar.asp?ProductID=<%=Request("ProdID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>

</HTML>
