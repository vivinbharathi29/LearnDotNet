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
        var _strStatus = "";
        var _RowID = "";
        var _VersionID = "";
        var _TodayPageSection = "";

        function Cancel() {           
            if (_TodayPageSection == "EditAccessoryStatus2") {
                window.parent.EditAccessoryStatusResult2(_RowID);
            }
            else if (_TodayPageSection == "EditAccessoryStatus") {
                window.parent.EditAccessoryStatusResult(_RowID, _strStatus);
            }
            else {
                window.parent.EditAccessoryStatusResult(_VersionID, _strStatus);
            }
        }

        function closewindow(VersionID, RowID, newStatus, TodayPageSection) {
            if (TodayPageSection == "EditAccessoryStatus2") {
                window.parent.EditAccessoryStatusResult2(RowID);
            }
            else if (TodayPageSection == "EditAccessoryStatus") {
                window.parent.EditAccessoryStatusResult(RowID, newStatus);
            }
            else {
                window.parent.EditAccessoryStatusResult(VersionID, newStatus);
            }
        }

        function SetNewStatus(RowID, VersionID, newStatus, TodayPageSection) {
            _strStatus = newStatus;
            _RowID = RowID;
            _VersionID = VersionID;
            _TodayPageSection = TodayPageSection;
        }

        function repositionParentWindow() {
            window.parent.reposition();
        }
    </script>
</head>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="AccessoryStatusMainPulsar.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&ReleaseID=<%=Request("ReleaseID")%>&RowID=<%=Request("RowID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="AccessoryStatusButtonsPulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>