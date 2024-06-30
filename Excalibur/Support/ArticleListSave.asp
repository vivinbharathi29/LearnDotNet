<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value != "0")
		{
		    if (GetQueryStringValue("Type") == "Article") {
		        window.parent.parent.parent.ShowArtileListCallBack(txtSuccess.value);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        //window.returnValue = txtSuccess.value;
		        //window.parent.close();
		        parent.window.parent.ShowArtileList_return(txtSuccess.value);		        parent.window.parent.modalDialog.cancel(false);
		    }
			}
		}
	
}
//-->
</SCRIPT>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</STYLE>
</HEAD>



<BODY onload="window_onload();">

<INPUT type="text" style="display:" id=txtSuccess name=txtSuccess value="<%=server.htmlencode(request("chkArticle"))%>">

</BODY>
</HTML>




