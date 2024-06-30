<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Edit Scorecard</TITLE>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="RootScorecardMain.asp?ID=<%=Request("ID")%>&Action=<%=Request("Action")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="RootScorecardButtons.asp">
</FRAMESET>

</HTML>