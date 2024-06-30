<%@ Language=VBScript %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

%>


<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
    <title>Update Deliverable Status</title>
    <script src="Scripts/PulsarPlus.js"></script>
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

        function Close(RowID, DAS)
        {
            if (IsFromPulsarPlus()) {
                window.parent.parent.popupCallBack(RowID);
                ClosePulsarPlusPopup();
            } else {
                window.parent.DisplayHWDeveloperSignoffResult(RowID + ";" + DAS);
            }            
        }
    </script>
</head>
<frameset rows="*,50" id="TopWindow">
    <FRAME ID="UpperWindow" Name="UpperWindow" SRC="UpdateDevStatusMain.asp?ID=<%=request("ID")%>&StatusID=<%=request("StatusID")%>&TypeID=<%=Request("TypeID")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="UpdateDevStatusButton.asp" scrolling=no>
</frameset>

</html>
