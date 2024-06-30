<%@ Language=VBScript %>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="style/general.css">
<meta  name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script  id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

function LoadFileWindow_onload() {
	window.location.replace("FileBrowseError.asp");
}

function window_onload() {
	window.parent.location.replace(text1.value);
}

function window_onload2() {
	window.location.replace(text1.value);
}

//-->
</script>
</HEAD>

<%
'Get User
dim CurrentDomain
dim Currentuser
CurrentUser = lcase(Session("LoggedInUser"))

if instr(currentuser,"\") > 0 then
	CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
	Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
end if

if request("DeliverablePath") <> "" then 
	if request("Path2Location")="" and request("Path3Location")="" and request("TDCImagePath")="" then %>
		<BODY LANGUAGE=javascript onload="return window_onload()">
	<% else %>
		<BODY LANGUAGE=javascript onload="return window_onload2()">
	<% end if %>
<%elseif request("TDCImagePath")<> "" then 
    if (CurrentDomain <> "asiapacific" and CurrentDomain <> "excaliburweb")then %>
		<BODY LANGUAGE=javascript onload="return window_onload2()">
	<% else
	   if request("Path2Location")="" and request("Path3Location")="" then %>
			<BODY LANGUAGE=javascript onload="return window_onload()">
	   <%else%> 
			<BODY LANGUAGE=javascript onload="return window_onload2()">
	   <%end if%>
	<% end if %>
<% end if %>
<%
'Response.Flush()
%>

<!--WOStatisticsStore -->
<%Request("Instr1")%>

<%
'if the current user is not TDC or ODM in Asia, the default display screen shows the Houston path
if (CurrentDomain <> "asiapacific" and CurrentDomain <> "excaliburweb") then
	if Request("DeliverablePath")<>"" then	%>
	<%
	'Response.write Request("DeliverablePath")
	%>
		<Iframe style="display:none" id=LoadFileWindow SRC="file://<%=Request("DeliverablePath")%>" style="Width:100%;Height:100%" LANGUAGE=javascript onerror="return LoadFileWindow_onload()" onload="return LoadFileWindow_onload()">
		</Iframe>
	<%elseif request("TDCImagePath")<>"" then%>
		<Iframe style="display:none" id=LoadFileWindow SRC="file://<%=Request("TDCImagePath")%>" style="Width:100%;Height:100%" LANGUAGE=javascript onerror="return LoadFileWindow_onload()" onload="return LoadFileWindow_onload()">
		</Iframe>
	<%end if%>
<%else 
	if request("TDCImagePath")<>"" then%>
		<Iframe style="display:none" id=LoadFileWindow SRC="file://<%=Request("TDCImagePath")%>" style="Width:100%;Height:100%" LANGUAGE=javascript onerror="return LoadFileWindow_onload()" onload="return LoadFileWindow_onload()">
		</Iframe>
	<%elseif Request("DeliverablePath")<>"" then%>
		<Iframe style="display:none" id=LoadFileWindow SRC="file://<%=Request("DeliverablePath")%>" style="Width:100%;Height:100%" LANGUAGE=javascript onerror="return LoadFileWindow_onload()" onload="return LoadFileWindow_onload()">
		</Iframe>
	<%end if%>
<%end if%>
<!--
if (parent.frames.myFrame &amp;&amp; (!parent.frames.myFrame.document.readyState 
|| parent.frames.myFrame.document.readyState == "complete")) 
-->
<% if request("DeliverablePath") <> "" then %>
	<INPUT type="hidden" id=text1 name=text1 value="<%=Request("DeliverablePath")%>" />
<% elseif request("TDCImagePath") <> "" then %>
	<INPUT type="hidden" id=text1 name=text1 value="<%=Request("TDCImagePath")%>" />
<% end if %>
</BODY>
</HTML>
