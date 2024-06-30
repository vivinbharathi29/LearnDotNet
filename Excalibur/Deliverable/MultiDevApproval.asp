<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<html>
<head>
    <title>Update Developer Approval</title>
    <script src="../Scripts/PulsarPlus.js"></script>
    <script>
        function Cancel()
        {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                window.parent.closeModalDialog(false);
            } 
        }

        function Close(RowIDs)
        {
            if (IsFromPulsarPlus()) {
                window.parent.parent.popupCallBack(1);
                ClosePulsarPlusPopup();
            } else {
                window.parent.DisplayHWDeveloperSignoffResult(RowIDs);
            }            
        }
    </script>
</head>

<frameset rows="*,60" id="TopWindow">
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MultiDevApprovalMain.asp?NewValue=<%=Request("NewValue")%>&txtMultiID=<%=Request("txtMultiID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MultiDevApprovalButtons.asp">
	</frameset>
</html>