<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  
  Dim Singles : Singles = Request("Singles") 
%>
	
<HTML>
<HEAD>
<title>PhWeb AV Action Items</title>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
<% If Singles = 0 Then %>
		<FRAME ID="MyWindow" Name="MyWindow" SRC="PDMFeedbackMainMultiples2.asp?<%=Request.QueryString %>">
<% Else %>
		<FRAME ID="MyWindow" Name="MyWindow" SRC="PDMFeedbackMain2.asp?<%=Request.QueryString %>">
<% End If %>
</FRAMESET>
</HTML>
