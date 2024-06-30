<!-- #include file="Users.inc" -->
<%
function GetGlobalFooter()
	dim sHTML
	
	sHTML = ""

	sHTML = sHTML & "<table BORDER=0 CELLSPACING=0 CELLPADDING=5 width=""100%"">" & vbCrLF
	sHTML = sHTML & "<tr>" & vbCrLF
	' add border:none so that Report pages work okay too since they use CSS for TD
	sHTML = sHTML & "<td colspan=3 width=""100%"" style=""border:none"">" & vbCrLf
	sHTML = sHTML & "<a style=""text-decoration: underline;color: blue;"" href=""#top"">"
	sHTML = sHTML & "<img src=""" & Session("ServerPath") & "/library/images/gifs/top.gif"" alt=""Top"" border=""0"" WIDTH=""10"" HEIGHT=""7"">"
	sHTML = sHTML & "<font SIZE=1>Top</font></a></td></tr>" & vbCrLF

	sHTML = sHTML & "<tr><td colspan=3 width=""100%"" style=""border:none""><font SIZE=1>"
	sHTML = sHTML & Users_CreateUserInfo(False)
	sHTML = sHTML & "</font></td></tr>" & vbCrLF

	sHTML = sHTML & "<tr><td width=""33%"" style=""border:none""><font SIZE=1>"
	if Session("LastUpdated") <> "" then
		sHTML = sHTML & "Last Updated on " & Session("LastUpdated")
	else
		'get the last updated date from the database
		call GetLastUpdatedDate()
		if Session("LastUpdated") <> "" then
			sHTML = sHTML & "Last Updated on " & Session("LastUpdated")
		end if
	end if
	sHTML = sHTML & " </font></td>" & vbCrLF

	sHTML = sHTML & "<td width=""33%"" style=""border:none"" align=center><font SIZE=1>"
	sHTML = sHTML & "All times displayed are Houston time (CST)</font></td>" & vbCrLF
	
	sHTML = sHTML & "<td width=""34%"" style=""border:none"" align=right><font SIZE=1>HP Restricted&nbsp;&nbsp;&nbsp;"
	sHTML = sHTML & "&copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.</font></td>" & vbCrLF
	sHTML = sHTML & "</tr>" & vbCrLF
	sHTML = sHTML & "</table>" & vbCrLf
	GetGlobalFooter = sHTML 
end function

function GetConfidentialFooter()
	dim sHTML
	
	sHTML = ""

	sHTML = sHTML & "<table BORDER=0 CELLSPACING=0 CELLPADDING=5 width=""100%"">" & vbCrLF
	sHTML = sHTML & "<tr>" & vbCrLF
	' add border:none so that Report pages work okay too since they use CSS for TD
	sHTML = sHTML & "<td colspan=3 width=""100%"" style=""border:none"">" & vbCrLf
	sHTML = sHTML & "<a style=""text-decoration: underline;color: blue;"" href=""#top"">"
	sHTML = sHTML & "<img src=""" & Session("ServerPath") & "/library/images/gifs/top.gif"" alt=""Top"" border=""0"" WIDTH=""10"" HEIGHT=""7"">"
	sHTML = sHTML & "<font SIZE=1>Top</font></a></td></tr>" & vbCrLF

	sHTML = sHTML & "<tr><td colspan=3 width=""100%"" style=""border:none""><font SIZE=1>"
	sHTML = sHTML & Users_CreateUserInfo(False)
	sHTML = sHTML & "</font></td></tr>" & vbCrLF

	sHTML = sHTML & "<tr><td width=""33%"" style=""border:none""><font SIZE=1>"
	if Session("LastUpdated") <> "" then
		sHTML = sHTML & "Last Updated on " & Session("LastUpdated")
	else
		'get the last updated date from the database
		call GetLastUpdatedDate()
		if Session("LastUpdated") <> "" then
			sHTML = sHTML & "Last Updated on " & Session("LastUpdated")
		end if
	end if
	sHTML = sHTML & " </font></td>" & vbCrLF

	sHTML = sHTML & "<td width=""33%"" style=""border:none"" align=center><font SIZE=1>"
	sHTML = sHTML & "All times displayed are Houston time (CST)</font></td>" & vbCrLF
	
	sHTML = sHTML & "<td width=""34%"" style=""border:none"" align=right><font SIZE=1>HP Confidential-PSG Business PC&nbsp;&nbsp;&nbsp;"
	sHTML = sHTML & "&copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.</font></td>" & vbCrLF
	sHTML = sHTML & "</tr>" & vbCrLF
	sHTML = sHTML & "</table>" & vbCrLf
	GetConfidentialFooter = sHTML 
end function

function 
	dim sHTML
	sHTML = GetGlobalFooter()
	Response.Write( sHTML )
end function

function InsertConfidentialFooter()
	dim sHTML
	sHTML = GetConfidentialFooter()
	Response.Write( sHTML )
end function

function InsertHelpGlobalFooter()
	response.write "<p><font size=1>HP Restricted<br>" & vbCrLf
	response.write "&copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.</br>" & vbCrLF
	response.write "</font></p>" & vbCrLf
end function

function InsertXMLConfidentialGlobalFooter()
	response.write "<Footer x:Data=""HP Confidential &copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.""/>"
end function

function InsertXMLRestrictedGlobalFooter()
	response.write "<Footer x:Data=""HP Restricted &copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.""/>"
end function

function RestrictedCopyrightTextHTML()
	RestrictedCopyrightTextHTML =  "<Footer x:Data=""HP Restricted &copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.""/>"
end function

function ConfidentialCopyrightTextHTML()
	ConfidentialCopyrightTextHTML =  "<Footer x:Data=""HP Confidential &copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P.""/>"
end function

function LCMCopyrightTextHTML()
	'need to remove the &copy symbol because it is not displayable in SoftPaq text file
	'LCMCopyrightTextHTML = "&copy; Copyright " & cstr(year(now)) & " HP Development Company, L.P."
	LCMCopyrightTextHTML = "Copyright " & cstr(year(now)) & " HP Development Company, L.P."
end function

sub GetLastUpdatedDate()
	'get the Last Updated Date from the database depending on the web server we're on
	dim oRecords, oError, oRs
	
	Session("LastUpdated") = ""	
	set oRecords = Server.CreateObject("JF_S_Admin.ISAdmin")
	set oError = oRecords.WebServer_GetLastUpdatedDate(Application("Repository"), Application("WebServer"), oRs )
	if not oError is nothing then
		Response.Write Errors_GetHTML(oError)
		set oError = nothing
	else
		if oRs.RecordCount > 0 then 
			Session("LastUpdated") = FormatDateTime(oRs("LastUpdated"), vbLongDate)
		end if
	end if
	
	oRs.Close
	set oRs = nothing
	set oRecords = nothing
	set oError = nothing
end sub
%>
