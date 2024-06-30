<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<%  
    dim strTitle
    if request("PrepComplete") = "1" then
        strTitle = "Prep Complete"
    else
        strTitle = "Edit Preinstall Properties"
    end if
%>
<HEAD>
<TITLE><%=strTitle%></TITLE>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PreinstallPropertiesMain.asp?ID=<%=Request("ID")%>&PrepComplete=<%=request("PrepComplete")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="PreinstallPropertiesButtons.asp">
</FRAMESET>

</HTML>