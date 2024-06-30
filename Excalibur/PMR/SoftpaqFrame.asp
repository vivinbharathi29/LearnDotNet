<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

	%>


<HTML>
    <head>
	    <TITLE><%=request("Title")%></TITLE>
    </head>

    <body>
        <iframe frameborder="0" height="100%" width="100%" src="<%=request("Page")%>?<%=replace(replace(replace(Request.QueryString,"%20"," "),"Page=" & request("Page") & "&",""),"Title=" & request("Title") & "&","")%>"></iframe>
    </body>
</HTML>
