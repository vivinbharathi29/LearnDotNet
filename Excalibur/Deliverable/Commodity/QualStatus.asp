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
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        function Close(VersionID, ProductDeliverableID, strStatus, TodayPageSection, ProductID) {

            if (IsFromPulsarPlus()) {

                if (TodayPageSection == "EditCommodityStatusSignoff") {
                    //window.parent.parent.parent.EditCommodityStatus2ReloadCallback($("#txtSuccess").val());
                    window.parent.parent.parent.EditCommodityStatusSignoffResultCallback(ProductDeliverableID, 0, strStatus);
                    ClosePulsarPlusPopup();
                }
                if (TodayPageSection == "EditCommodityStatus2") {
                    window.parent.EditCommodityStatus2Result(ProductID, VersionID, strStatus);
                }
                else {
                    window.parent.EditCommodityStatusResult(VersionID, ProductDeliverableID, 0, strStatus);
                }

            }
            else
            {
                if (TodayPageSection == "EditCommodityStatusSignoff") {
                    window.parent.EditCommodityStatusSignoffResult(ProductDeliverableID, 0, strStatus);
                }
                if (TodayPageSection == "EditCommodityStatus2") {
                    window.parent.EditCommodityStatus2Result(ProductID, VersionID, strStatus);
                }
                else {
                    window.parent.EditCommodityStatusResult(VersionID, ProductDeliverableID, 0, strStatus);
                }

            }
            //if (TodayPageSection == "EditCommodityStatusSignoff") {
            //    window.parent.EditCommodityStatusSignoffResult(ProductDeliverableID, 0, strStatus);
            //}
            //if (TodayPageSection == "EditCommodityStatus2") {
            //    window.parent.EditCommodityStatus2Result(ProductID, VersionID, strStatus);
            //}
            //else {
            //    window.parent.EditCommodityStatusResult(VersionID, ProductDeliverableID, 0, strStatus);
            //}
        }
        
        function Cancel()
        {
            window.parent.EditCommodityStatusResult(0, 0, 0, "");
        }
    </script>
</head>

 <FRAMESET ROWS="*,60" ID=TopWindow >
  <%if Request("app")="PulsarPlus" then%>
<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="QualStatusMain.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&app=<%=Request("app")%>">
  <%else%>
 <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="QualStatusMain.asp?ProdID=<%=Request("ProdID")%>&VersionID=<%=Request("VersionID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
 <%end if%>
 <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="QualStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>
</HTML>