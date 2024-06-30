<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

	%>


<HTML>
<HEAD>
<TITLE>Reject Deliverable Request</TITLE>
<script type="text/javascript" src="../Scripts/PulsarPlus.js"></script>
<script>
    function cancel(){
        window.parent.closeModalDialog(false);
    }

    function closewindow(strValue,ID){
     if (IsFromPulsarPlus()) {      
        ClosePulsarPlusPopup();
         window.parent.parent.parent.popupCallBack(strValue);
    }
    else {
        window.parent.SingleConfirmResult(strValue,ID);
    }
       
    }
</script>
</HEAD>
<FRAMESET ROWS="*,50" ID=TopWindow>
    <FRAME ID="UpperWindow" Name="UpperWindow" SRC="DevRejectProductRequestMain.asp?ID=<%=request("ID")%>&NewValue=<%=request("NewValue")%>&Remaining=<%=request("Remaining")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="DevRejectProductRequestButton.asp" scrolling=no>
</FRAMESET>

</HTML>
