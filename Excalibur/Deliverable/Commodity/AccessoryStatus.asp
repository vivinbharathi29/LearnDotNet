<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
    <title>Update Accessory Status</title>
    <script type="text/javascript">
        function Close(VersionID, strStatus, TodayPageSection, RowID) {
            if (TodayPageSection == "EditAccessoryStatus") {
                window.parent.EditAccessoryStatusResult(RowID, strStatus);
            }
            else if (TodayPageSection == "EditAccessoryStatus2") {
                window.parent.EditAccessoryStatusResult2(RowID);
            }
            else {
                window.parent.EditAccessoryStatusResult(VersionID, strStatus);
            }
        }

        function Cancel() {
            window.parent.closeModalDialog(false);
        }
    </script>
</head>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="AccessoryStatusMain.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&RowID=<%=Request("RowID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="AccessoryStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>