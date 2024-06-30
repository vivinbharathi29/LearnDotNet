<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
	<title>Multiple Products RTM</title>
    <FRAMESET ROWS="*,60" ID=TopWindow >
    
       	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MultipleRTMMain.asp?ID=<%=request("ID")%>&IDS=<%=request("IDS") %>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
    	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MultipleRTMButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
    </FRAMESET>
    
</html>