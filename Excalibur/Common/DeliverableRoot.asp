<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Deliverable Lookup</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="DeliverableRootMain.asp?AddressList=<%=request("AddressList")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="DeliverableRootButtons.asp" scrolling=no>
</FRAMESET>

</HTML>