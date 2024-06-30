<%@ Language=VBScript %>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<HTML>
<%If Request("AddNew") = 1 then %>
<TITLE>Add Workflow</TITLE>
<% else%>
<TITLE>View Workflow Status</TITLE>
<% end if%>

<FRAMESET ROWS="*" ID=TopWindow >
<%If Request("AddNew") = 1 then %>
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="DCRWorkflow.aspx?<%=Request.QueryString %>">
<% else%>
    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="DCRWorkflowStatus.aspx?<%=Request.QueryString %>">
<% end if%>
</FRAMESET>
	
</HTML>