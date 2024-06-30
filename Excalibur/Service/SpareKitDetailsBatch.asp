<%@  language="VBScript" %>
<!--#include file="../includes/DataWrapper.asp"-->
<!--#include file="../includes/no-cache.asp"-->
<% 

    Dim strPVID : strPVID=Request.QueryString("PVID")
    Dim strSFPN : strSFPN=Request.QueryString("SFPN")
    Dim strSKIDs: strSKIDs=Request.QueryString("SKIDs")

    
	dim ConfigErrors
	
	Set MyBrow=Server.CreateObject("MSWC.BrowserType")
		
	ConfigErrors = ""
	if lcase(MyBrow.browser) <> "ie" or clng(left(MyBrow.version,1)) < 4 then
		ConfigErrors = ConfigErrors & "Internet Explorer 4.0 or greater is required<br>"		
	end if
	
	if not MyBrow.frames then
		ConfigErrors = ConfigErrors & "Frames must be enabled<br>"
	end if
	if not MyBrow.tables then
		ConfigErrors = ConfigErrors & "Table support is required.<br>"
	end if
	if not MyBrow.cookies then
		ConfigErrors = ConfigErrors & "Cookie support is required.<br>"
	end if
	if not MyBrow.javascript then
		ConfigErrors = ConfigErrors & "Javascript support is required.<br>"
	end if
	
	Dim pageTitle


    If pageTitle <> "" Then 
        pageTitle = "Spare Kit Details - Batch Update - " & pageTitle
    Else
        pageTitle = "Spare Kit Details - Batch Update"
    End If

%>
<html>
<head>
    <title><%=pageTitle %></title>
</head>
<frameset rows="*,55" id="TopWindow">
<frame id="UpperWindow" name="UpperWindow" src="SpareKitDetailsMainBatch.asp?PVID=<%=strPVID%>&SFPN=<%=strSFPN%>&SKIDS=<%=strSKIDs%>">
<frame ID="LowerWindow" name="LowerWindow" src="SpareKitDetailsButtonsBatch.asp?PVID=<%=strPVID%>&SFPN=<%=strSFPN%>&SKIDS=<%=strSKIDs%>" scrolling="no">
</frameset>
</html>
