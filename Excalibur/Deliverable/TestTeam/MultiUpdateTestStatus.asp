<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Batch Update Test Status</TITLE>
<HEAD>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
        function Cancel()
        {
            window.parent.closeModalDialog(false);
        }

        function close(TodayPageSection, RowIDs, Index)
        {
            if (IsFromPulsarPlus()) {
                window.parent.parent.parent.MultiTestStatusCallback(txtSuccess.value);
                ClosePulsarPlusPopup();
            }
            else {
                if (TodayPageSection == "") {
                    window.parent.closeModalDialog(false);
                }
                else if (TodayPageSection == "MultiEditTestStatus") {
                    window.parent.MultiEditTestStatusResult(Index, RowIDs);
                }
            }
           
        }
    </script>
</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>       
    <FRAME ID="UpperWindow" Name="UpperWindow" SRC="MultiUpdateTestStatusMain.asp?IDList=<%=request("IDList")%>&ProductID=<%=request("ProductID")%>&TodayPageSection=<%=request("TodayPageSection")%>&Index=<%=request("Index")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="MultiUpdateTestStatusButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" scrolling=no>
</FRAMESET>

</HTML>