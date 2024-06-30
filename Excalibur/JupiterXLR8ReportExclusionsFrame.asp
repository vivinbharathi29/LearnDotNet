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
<TITLE>Jupiter XLR8 Report Exclusions</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,50" ID=TopWindow>
    <FRAME ID="UpperWindow" Name="UpperWindow" SRC="JupiterXLR8ReportExclusionsMain.asp">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="JupiterXLR8ReportExclusionsButton.asp" scrolling=no>
</FRAMESET>

</HTML>
