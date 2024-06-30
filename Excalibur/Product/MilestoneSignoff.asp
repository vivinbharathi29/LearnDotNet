<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
	<title>Product RTM</title>

    <FRAMESET ROWS="*,65" ID=TopWindow >
    	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MilestoneSignoffMain.asp?ID=<%=Request("ID")%>&ProductRTMID=<%=Request("ProductRTMID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
    	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MilestoneSignoffButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
    </FRAMESET>
</html>