<%@ Language=VBScript %>

	<%

    'if lcase(Session("LoggedInUser")) = "auth\dwhorton" or lcase(Session("LoggedInUser")) = "auth\wgomero"  or lcase(Session("LoggedInUser")) = "auth\emartinez"   or lcase(Session("LoggedInUser")) = "auth\dismusco"   or lcase(Session("LoggedInUser")) = "auth\buidi"  or lcase(Session("LoggedInUser")) = "auth\letha"   or lcase(Session("LoggedInUser")) = "auth\carrolld"  or lcase(Session("LoggedInUser")) = "auth\lyoung"   or lcase(Session("LoggedInUser")) = "auth\lyoung" then
        'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Workflow&Page=/Pulsar/Component/ReleaseWidget&ComponentVersionId=" & request("ID") & "&ActionType=" & clng(request("Action"))-1  & "&app=" & request("app")
        Response.Redirect "/Pulsar/Component/ReleaseWidget?ComponentVersionId=" & request("ID") & "&ActionType=" & clng(request("Action"))-1  & "&app=" & request("app")
        response.end
    'end if

  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Release Deliverable Version</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="Releasemain.asp?Action=<%=request("Action")%>&ID=<%=request("ID")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="ReleaseButtons.asp" scrolling=no>
</FRAMESET>

</HTML>
