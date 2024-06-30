<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>
<html>
<head>
    <title>Batch Update Qualification Status</title>
	<meta charset="utf-8" />
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        function Cancel()
        {  
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                window.parent.closeModalDialog(false);
            }
        }

        function closewindow(RowIDs, TodayPageSection, StatusID)
        {
            if (TodayPageSection == "MultiAccessory" && StatusID != 1)
            {
                window.parent.MultiAccessoryResult(RowIDs);
            }
            else if (TodayPageSection == "MultiUpdateTestStatus" && StatusID <= 1)
            {
                window.parent.MultiUpdateTestStatusResult(RowIDs);
            }
            else if (TodayPageSection == "MultiCommodity" && StatusID != 1) {
                window.parent.MultiCommodityResult(RowIDs);
            }
            else if (TodayPageSection == "MultiPilot" && StatusID != 1) {
                
            }
            else {
                if (IsFromPulsarPlus()) {
                    ClosePulsarPlusPopup();
                }
                else {
                    window.parent.closeModalDialog(false);
                }
            }
        }

        function RepositionPopup()
        {
            window.parent.reposition();
        }
    </script>
</head>

    <frameset rows="*,60" id="TopWindow">
        <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="MultiTestStatusMainPulsar.asp?RootID=<%=Request("RootID")%>&ProdID=<%=Request("ProdID")%>&VersionList=<%=Request("VersionList")%>&Type=<%=Request("Type")%>&ProductVersionReleaseID=<%=Request("ReleaseID")%>&BSID=<%=Request("BSID")%>&ShowOnlyTargetedRelease=<%=Request("ShowOnlyTargetedRelease")%>&TodayPageSection=<%=Request("TodayPageSection")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>"></FRAME>
	    <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MultiTestStatusButtonsPulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>"></FRAME>
    </frameset>

</html>