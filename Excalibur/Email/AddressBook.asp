<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Address Book</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="AddressBookMain.asp?AddressList=<%=request("AddressList")%>&ShowAll=<%=request("ShowAll")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="AddressBookButtons.asp" scrolling=no>
</FRAMESET>

</HTML>