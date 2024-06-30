<HTML>
<HEAD>
<%if request("TypeID")= "1" then%>
	<title>System Board ID</title>
<%else%>
	<title>Machine PNP ID</title>
<%end if%>
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="ProductIDMain.asp?TypeID=<%=Request("TypeID")%>&IDList=<%=Request("IDList")%>">
	<FRAME ID="MyButtons" Name="MyButtons" SRC="ProductIDButtons.asp">
</FRAMESET>
</HTML>
