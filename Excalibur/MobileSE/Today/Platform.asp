<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
  Dim AppRoot
  AppRoot = Session("ApplicationRoot")

	%>

<HTML>
    <%if request("ID") = "" then %>
        <TITLE>Add Base Unit Group</TITLE>
    <%else%>
        <TITLE>Update Base Unit Group</TITLE>
    <%end if%>
<HEAD>
    <script src="../../includes/client/jquery-1.11.0.min.js"></script>
    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        function AddPlatformCompleted() {
            window.parent.ClosePlatFormDetail("Dialog1", true);
        }
        function CloseDialog1() {
            try {
                window.parent.CloseDialog1();
            }
            catch (e) {
                window.close();
            }
        }
        function CloseDialog2() {
            try {
                window.parent.CloseDialog2();
            }
            catch (e) {
                window.close();
            }
        }        
        function ClosePlatFormDetail(DialogID, Refresh)
        {
            window.parent.ClosePlatFormDetail(DialogID, Refresh);
        }
        function OpenDialog1(url, title, sWidth, sHeight, resizable, modal, sbuttons, disableScrollBar) {
            window.parent.OpenDialog1(url, title, sWidth, sHeight, resizable, modal, sbuttons, disableScrollBar);
        }
        function OpenDialog2(url, title, sWidth, sHeight, resizable, modal, sbuttons, disableScrollBar) {
            try{
                window.parent.OpenDialog2(url, title, sWidth, sHeight, resizable, modal, sbuttons, disableScrollBar);
            }
            catch(e)
            {
                var returnValue = window.showModalDialog(url, title, "dialogwidth:" + sWidth + "px; dialogheight:" + sHeight + "px");
                if (returnValue == 1)
                    window.location.reload();
            }
        }
    </script>
</HEAD>
<FRAMESET  ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="PlatformMain.asp?ID=<%=request("ID")%>&ProductVersionId=<%=request("ProductVersionID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>&FollowMKTName=<%=request("FollowMKTName")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="PlatformButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>&FollowMKTName=<%=request("FollowMKTName")%>">
</FRAMESET>

</HTML>
