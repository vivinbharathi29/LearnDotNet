<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<head>
    <title>Update Subassembly</title>
    <script type="text/javascript">
       
        function Close(RowID) {
            if (RowID == "")
                window.parent.closeModalDialog(true);
            else
                window.parent.SubAssemblyResult(RowID);
        }

        function Cancel() {
            window.parent.closeModalDialog(false);
        }
    </script>
</head>
	<FRAMESET ROWS="*" ID=TopWindow >
        <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="SubAssemblyMain.asp?VersionID=<%=Request("VersionID")%>&ProductID=<%=Request("ProductID")%>&RootID=<%=Request("RootID")%>&IDList=<%=request("IDList")%>&RowID=<%=Request("RowID")%>&TodayPageSection=<%=Request("TodayPageSection")%>">    
	</FRAMESET>
</HTML>