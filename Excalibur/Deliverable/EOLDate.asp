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
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="EOLDateMain.asp?TypeID=<%=Request("TypeID")%>&ID=<%=Request("ID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="EOLDateButtons.asp">
</FRAMESET>

</HTML>