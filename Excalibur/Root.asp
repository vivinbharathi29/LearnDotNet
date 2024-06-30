<%@ Language=VBScript %>
<%
    if request("ID") = "" then
            'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Version Properties&Page=/Pulsar/Component/AddNewRoot&app=" & request("app") 
            Response.Redirect "/Pulsar/Component/AddNewRoot?app=" & request("app") 
    else
            'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Version Properties&Page=/Pulsar/Component/Root/" & request("ID")  & "&app=" & request("app")
            Response.Redirect "/Pulsar/Component/Root?id=" & request("ID")  & "&showMegamenu=false"
    end if

  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
