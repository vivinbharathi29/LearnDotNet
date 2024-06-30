<%@ Language=VBScript %>
<html>
    <head>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


    function Test(strTest) {
        var patt = new RegExp("^ *[0-9]{5,8} *( *,{1} *[0-9]{5,8} *)*$")
       
        var result = patt.test(strTest);
        alert (result);
    }

    function Body_Onload() {
        Test('600000 ,888837');
    }
//-->
</SCRIPT>
    
    </head>
	<body onload="Body_Onload();">
<%

'Dim MyNumber, strMyString

'MyString = "300 South Street"

'response.write val(MyString)

'Response.Write datediff("d", NOW(), "4/10/2009")
'Response.End

'On Error Resume Next
    'response.Write "<p>"
    'Response.Write "<b>ClientCertificate:</b>" & Request.ClientCertificate & "</br>"
    'Response.Write "<b>Cookies:</b>" & Request.Cookies & "</br>"
    'Response.Write "<b>Form:</b>" & Request.Form & "</br>"
    'Response.Write "<b>QueryString:</b>" & Request.QueryString & "</br>"
    'Response.Write "<b>TotalBytes:</b>" & Request.TotalBytes & "</br>"
    'response.Write "</p>"

    'response.Write "<p>"
    'REsponse.Write err.number & "<br>"
    'response.Write err.Description & "<br>"
    'response.Write "</p>"

    
 '   response.Write "<h2>Cookies</h2>"
 '   response.Write "<p>"
 '   for each cookie in request.Cookies
 '           Response.Write cookie & " : " & request.Cookies(cookie) & "<br>"
 '   next
 '   response.Write "</p>"
 '
'	Dim x
'    response.Write "<h2>Server Variables</h2>"
'    response.Write "<p>"
'	for each x in Request.ServerVariables
'        Response.Write x & ":" & Request.ServerVariables(x) & "<br />"
'    next
'    response.Write "</p>"
'
'    function Val(strText)
'        dim strOutput
'        dim i

 '       strOutput = ""
 '       for i = 1 to len(trim(strText))
 '           if isnumeric(mid(strText,i,1)) then
 '               strOutput = strOutput & mid(trim(strText),i,1)
 '           else
 '               exit for
 '           end if
 '       next
 '       Val = strOutput
 '   end function

%>

</body></html>