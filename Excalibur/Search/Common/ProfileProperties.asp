<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<HEAD>
<%if Request("ID") = "" then%> 
    <%if request("ReportType") = "1" then%>
        <TITLE>Add Custom Report</TITLE>
    <%else%>
        <TITLE>Add Profile</TITLE>
    <%end if%>
<%else%>
    <%if request("ReportType") = "1" then%>
        <TITLE>Update Custom Report</TITLE>
    <%else%>
        <TITLE>Update Profile</TITLE>
    <%end if%>
<%end if%>
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProfilePropertiesMain.asp?ID=<%=Request("ID")%>&ReportType=<%=request("ReportType")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProfilePropertiesButtons.asp">
</FRAMESET>

</HTML>