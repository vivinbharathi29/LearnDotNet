<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
    <title>Update Pilot Status</title>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        function Close(VersionID, ProductDeliverableID, strStatus, TodayPageSection) {
            if (IsFromPulsarPlus()) {
                window.parent.parent.parent.popupCallBack(1);
                ClosePulsarPlusPopup();
            }
            else {
                if (TodayPageSection == "EditPilotStatus") {
                    window.parent.EditPilotStatusResult(ProductDeliverableID, 0, strStatus);
                }
                else {
                    if (strStatus.indexOf("_") != -1)
                        strStatus = strStatus.slice(2);

                    window.parent.EditPilotStatusResult(VersionID + "|" + strStatus);
                }
            }
        }

        function Cancel(TodayPageSection) {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                if (TodayPageSection == "EditPilotStatus") {
                    window.parent.EditPilotStatusResult(0, 0, "");
                }
                else {
                    window.parent.EditPilotStatusResult("");
                }
            }
        }
    </script>
</head>

<frameset rows="*,60" id="TopWindow">
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PilotStatusMain.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="PilotStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</frameset>
</html>