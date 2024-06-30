<%@ Language=VBScript %>


	<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<html>
<head>
    <title>Update Qualification Status</title>
    <script type="text/javascript">
        var strStatus = "";
        function Cancel() {
            window.parent.EditCommodityStatusResult(0, 0, 0, strStatus);
        }

        function closewindow(VersionID, ProductDeliverableID, ProductDeliverableReleaseID, newStatus, TodayPageSection, ProductID) {
            if (TodayPageSection == "EditCommodityStatusSignoff") {
                window.parent.EditCommodityStatusSignoffResult(ProductDeliverableID, ProductDeliverableReleaseID, newStatus);
            }
            if (TodayPageSection == "EditCommodityStatus2") {
                window.parent.EditCommodityStatus2Result(ProductID, VersionID, newStatus);
            }
            else {
                window.parent.EditCommodityStatusResult(VersionID, ProductDeliverableID, ProductDeliverableReleaseID, newStatus);
            }
        }

        function SetNewStatus(newStatus) {
            strStatus = newStatus;
        }

        function repositionParentWindow()
        {
            window.parent.reposition();
        }
    </script>
</head>

<frameset rows="*,60" id="TopWindow">
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="QualStatusMainPulsar.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&ReleaseID=<%=Request("ReleaseID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&ShowOnlyTargetedRelease=<%=Request("ShowOnlyTargetedRelease")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="QualStatusButtonsPulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</frameset>

</html>
