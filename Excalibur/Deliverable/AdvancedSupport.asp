<%@ Language=VBScript %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>


<html>
<head>
    <title>Advanced Deliverable Support</title>
    <script src="../Scripts/PulsarPlus.js"></script>
    <script id="clientEventHandlersJS" language="javascript">
    <!--

        function Cancel() {
            
            var IsFromPulsarPlus = false;
            
            try {
                IsFromPulsarPlus = IsFromPulsarPlus();
            }
            catch(ex)
            { }

            if (IsFromPulsarPlus) {
                ClosePulsarPlusPopup();
            } else {
                try {
                    window.parent.closeModalDialog(false);
                }
                catch (ex) { 
					window.close();
				}
            }
        }

        function Close(RowID) {

            var IsFromPulsarPlus = false;

            try {
                IsFromPulsarPlus = IsFromPulsarPlus();
            }
            catch (ex)
            { }

            if (IsFromPulsarPlus) {

                window.parent.parent.parent.popupCallBack(1);
                ClosePulsarPlusPopup();

            } else {
                try {
                    window.parent.DisplayAdvancedSupportResult(RowID);
                }
                catch (ex) {
                    window.returnValue = RowID;
                    window.close();
                }
            }
        }


    //-->
    </script>
</head>
<frameset rows="*,55" id="TopWindow">
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="AdvancedSupportMain.asp?ProdRootID=<%=request("ProdRootID")%>&VersionID=<%=request("VersionID")%>&RootID=<%=request("RootID")%>&ProductID=<%=request("ProductID")%>&ProductDeliverableReleaseID=<%=request("ProductDeliverableReleaseID")%>&RowID=<%=request("RowID")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="AdvancedSupportButtons.asp" scrolling=no>
</frameset>

</html>