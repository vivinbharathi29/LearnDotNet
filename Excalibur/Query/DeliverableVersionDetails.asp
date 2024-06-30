<%@ Language=VBScript %>

<%
' Redirect to new ASP.NET page
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "DeliverableVersionDetails.aspx?" & request.querystring
Response.End
%>

