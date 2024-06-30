<%@ Language="VBScript" %>

<%

test = "1"

test2 = split(test, ",")

response.write ubound(test2) & "<BR>"

For i = 0 to ubound(test2)
    response.write test2(i) & "<BR>"
Next

%>