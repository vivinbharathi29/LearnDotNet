<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<% if Request("ID") <> "" then %>
	<TITLE>Update Product Requirement</TITLE>
<% else%>
	<TITLE>Select Product Requirements</TITLE>
<% end if%>
<HEAD>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onunload() {
	if (typeof(	window.opener) != "undefined")
		{
		//window.opener.location.reload(true);
		}
}

//function window_onblur() {
//	self.focus(); 
//}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onunload>
<!--
 window_onunload()
-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=window EVENT=onblur>
<!--
 //window_onblur()
//-->
</SCRIPT>
</HEAD>
<%if request("ID") <> "" then%>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME frameborder="0" noresize ID="UpperWindow" Name="UpperWindow" SRC="RequirementMain.asp?ID=<%=Request("ID")%>&ProdID=<%=request("ProdID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
		<FRAME frameborder="0" noresize ID="LowerWindow" Name="LowerWindow" SRC="RequirementButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
	</FRAMESET>
<%elseif request("ProdID") <> "" then%>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME frameborder="0" noresize ID="UpperWindow" Name="UpperWindow" SRC="RequirementListMain.asp?ID=<%=Request("ProdID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
		<FRAME frameborder="0" noresize ID="LowerWindow" Name="LowerWindow" SRC="RequirementButtons.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
	</FRAMESET>
<%else%>
	<font size=2 face=verdana><BR>Unable to display the requested page because not enough information was supplied.</font>
<%end if%>

</HTML>