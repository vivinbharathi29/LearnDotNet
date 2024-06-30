<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Edit FCS Target Dates</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="FCSEditMain_Pulsar.asp?ID=<%=Request("ID")%>&isFromPulsarPlus=<%=Request("isFromPulsarPlus")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="FCSButtons.asp?isFromPulsarPlus=<%=Request("isFromPulsarPlus")%>">
</FRAMESET>

</HTML>