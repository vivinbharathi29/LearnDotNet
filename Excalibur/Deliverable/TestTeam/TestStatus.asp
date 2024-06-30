<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Update Test Status</TITLE>
<head>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        function Close() {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                window.parent.closeModalDialog(false);
            }            
        }
    </script>
</head>
<FRAMESET ROWS="*,55" ID=TopWindow>
   <FRAME ID="UpperWindow" Name="UpperWindow" SRC="TestStatusMain.asp?VersionID=<%=request("VersionID")%>&ProductID=<%=request("ProductID")%>&FieldID=<%=request("FieldID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="TestStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" scrolling=no>	
</FRAMESET>

</HTML>