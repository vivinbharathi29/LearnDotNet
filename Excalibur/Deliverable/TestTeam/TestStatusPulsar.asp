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
        var strStatus = "";
        var strTodayPageSection = "";
        var intRowID = 0, intFieldID = 0;

        function Close() {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                if (strTodayPageSection == "EditTestStatus") {
                    window.parent.EditTestStatusResult(intFieldID, intRowID)
                }
                window.parent.closeModalDialog(false);
            }            
        }

        function SetNewStatus(newStatus, TodayPageSection, FieldID, RowID) {
            strStatus = newStatus;
            strTodayPageSection = TodayPageSection;
            intRowID = RowID;
            intFieldID = FieldID;
        }

        function repositionParentWindow() {
            window.parent.reposition();
        }
    </script>
</head>
<frameset rows="*,55" id="TopWindow">
   <FRAME ID="UpperWindow" Name="UpperWindow" SRC="TestStatusMainPulsar.asp?VersionID=<%=request("VersionID")%>&ProductID=<%=request("ProductID")%>&FieldID=<%=request("FieldID")%>&ReleaseID=<%=Request("ReleaseID")%>&TodayPageSection=<%=Request("TodayPageSection")%>&RowID=<%=Request("RowID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="TestStatusButtonsPulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" scrolling=no>	
</frameset>

</HTML>