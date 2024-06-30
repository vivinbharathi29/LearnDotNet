<%@ Language=VBScript %>
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  
%>

<html>
<head>
    <title>Add Releases</title>
    <script>
        function getPageLocation(){
            return window.parent.getPageLocation();
        }
    </script>
</head>
<frameset rows="*,60" id="TopWindow">
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProductReleaseMain.asp?ID=<%=request("ID")%>&ProductTypeID=<%=request("ProductTypeID")%>&BusinessSegmentID=<%=request("ProductBusinessSegmentID")%>&isClone=<%=request("isClone")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProductReleaseButtons.asp">
</frameset>

</html>
