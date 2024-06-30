<%@ Language=VBScript %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>

<HTML>
<HEAD>
<TITLE>Mobile Tools Support</TITLE>

</HEAD>
<FRAMESET ROWS="*,65" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="SupportMain.asp?cboProject=<%=request("cboProject")%>&cboCategory=<%=request("cboCategory")%>&txtRequired=<%=request("txtRequired")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="SupportButtons.asp">
</FRAMESET>

</HTML>