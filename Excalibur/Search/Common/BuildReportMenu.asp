<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(){
        if (typeof(window.parent.CustomMenuOptions) !="undefined")
            window.parent.CustomMenuOptions.innerHTML = document.body.innerHTML;
    }

//-->
</SCRIPT>
</HEAD>
<BODY onload=window_onload();><%
    dim CurrentUserID

    CurrentUserID = clng(request("CurrentUserID"))
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    CustomStatusReports = ""
	rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	do while not rs.EOF
        CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'""><font face=Arial size=2><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></font></DIV>"
        rs.MoveNext
	loop
	rs.Close
		
	rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	do while not rs.EOF
        CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'""><font face=Arial size=2><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></font></DIV>"
    	rs.MoveNext
	loop
	rs.Close

	rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",7",cn,adOpenForwardOnly
	do while not rs.EOF
        CustomStatusReports = CustomStatusReports & "<DIV  onmouseover=""this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'"" onmouseout=""this.style.background='white';this.style.color='black'""><font face=Arial size=2><SPAN onclick=""javascript:StatusReport("  & rs("ID") & ");"">&nbsp;&nbsp;&nbsp;" & replace(rs("ProfileName")," ","&nbsp;") & "&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></font></DIV>"
    	rs.MoveNext
	loop
	rs.Close
if CustomStatusReports <> "" then
    CustomStatusReports = "<hr width=""95%"">" & CustomStatusReports & "<hr width=""95%"">"
else
    CustomStatusReports = "<hr width=""95%"">"
end if

    set rs = nothing
    cn.Close
    set cn = nothing
%><%=CustomStatusReports%></BODY>
</HTML>




