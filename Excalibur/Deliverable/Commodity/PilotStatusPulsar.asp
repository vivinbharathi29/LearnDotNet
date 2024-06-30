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
        var strStatus = "";
        var intProductDeliverableID = 0;
        var intProductDeliverableReleaseID = 0;
        var strTodayPageSection = "";

        function Cancel() {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                if (strTodayPageSection == "EditPilotStatus") {
                    window.parent.EditPilotStatusResult(intProductDeliverableID, intProductDeliverableReleaseID, strStatus);
                }
                else {
                    window.parent.EditPilotStatusResult(strStatus);
                }
            }
        }

        function closewindow(VersionID, ProductDeliverableID, ProductDeliverableReleaseID, newStatus, TodayPageSection) {
            if (IsFromPulsarPlus()) {
                window.parent.parent.parent.popupCallBack(1);
                ClosePulsarPlusPopup();
            }
            else {
                if (TodayPageSection == "EditPilotStatus") {
                    window.parent.EditPilotStatusResult(ProductDeliverableID, ProductDeliverableReleaseID, newStatus);
                }
                else {
                    window.parent.EditPilotStatusResult(newStatus);
                }
            }
        }

        function SetNewStatus(newStatus, ProductDeliverableID, ProductDeliverableReleaseID, TodayPageSection) {
            strStatus = newStatus;
            intProductDeliverableID = ProductDeliverableID;
            intProductDeliverableReleaseID = ProductDeliverableReleaseID;
            strTodayPageSection = TodayPageSection;
        }

        function repositionParentWindow() {
            window.parent.reposition();
        }
    </script>
</head>

<frameset rows="*,60" id="TopWindow">
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PilotStatusMainPulsar.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&ReleaseID=<%=Request("ReleaseID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&ProductDeliverableID=<%=Request("ProductDeliverableID")%>&ShowOnlyTargetedRelease=<%=Request("ShowOnlyTargetedRelease")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="PilotStatusButtonsPulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</frameset>
</html>
