<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
    <title>Deliverable Preference</title>
    <script type="text/javascript">
       
        function Close(reload) {
            window.parent.closeModalDialog(reload);
        }

        function Cancel() {
            Close(false);
        }

        function RepositionPopup() {
            window.parent.reposition();
        }
    </script>
</head>
<frameset rows="*,55" id="TopWindow">
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="RestrictMain.asp?RootID=<%=request("RootID")%>&VersionID=<%=request("VersionID")%>&ProductID=<%=request("ProductID")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="RestrictButtons.asp?RootID=<%=request("RootID")%>&VersionID=<%=request("VersionID")%>&ProductID=<%=request("ProductID")%>" scrolling=no>
</frameset>

</html>