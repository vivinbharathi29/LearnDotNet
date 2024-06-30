<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<%
    dim strTitle
    if Request("GroupID") = "" then  
        strTitle = "Add Group"
    else
        strTitle = "Edit Group"
    end if
%>

<HTML>
<HEAD>
<TITLE><%=strTitle%></TITLE>
</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProfileShareGroupMain.asp?GroupID=<%=Request("GroupID")%>">
</FRAMESET>

</HTML>