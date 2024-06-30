<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
 <head>
     <meta http-equiv="X-UA-Compatible" content="IE=8" />
	<title>Update Qualification Status</title>
    <script type="text/javascript">
        function Cancel()
        {
            closewindow(false);
        }

        function closewindow(reload)
        {
            window.parent.closeModalDialog(reload);
        }
    </script>
 </head>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MultiTestStatusMain.asp?RootID=<%=Request("RootID")%>&ProdID=<%=Request("ProdID")%>&VersionList=<%=Request("VersionList")%>&Type=<%=Request("Type")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MultiTestStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>