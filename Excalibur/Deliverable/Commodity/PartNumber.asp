<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
	<TITLE>Update Deliverable Version Part Number</TITLE>
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PartNumberMain.asp?VersionID=<%=Request("VersionID")%>">
	</FRAMESET>
</HTML>