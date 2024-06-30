<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<%
	dim rs
	dim cn
	dim strError
	dim strName
	dim strDeliverablePath
	dim strDeveloperEmail
	dim strDeveloper
	dim strInstruction
	dim strReplicater
	dim strReplicateOnly
	dim strPath2Loaction
	dim strPath2Description
	dim strPath3Loaction
	dim strPath3Description
	dim strTDCImagePath
	dim strArchived
	dim strArchivedPath2
	dim strArchivedPath3
	dim strArchivedTDC
	
	Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim RequestID : RequestID = regEx.Replace(Request("ID"), "")

	
	if RequestID <> "" then
		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		rs.Open "spGetVersionproperties " & clng(RequestID),cn,adOpenForwardOnly

		'dim i
		'for i = 0 to rs.Fields.count -1
		'Response.Write rs.Fields(i).Name & "<BR>"
		'next

		if rs.EOF and rs.BOF then 
			strError = "Could not found deliverable"
		else
			strError=""
			strName=rs("Name") & " " & rs("Version")
			if rs("Revision") & "" <> "" then
				strName = strName & "," & rs("Revision") & ""
			end if
			if rs("Pass") & "" <> "" then
				strName = strName & "," & rs("Pass") & ""
			end if
			strDeveloperEmail = rs("DeveloperEmail")
			strDeveloper = rs("Developer")
			strDeliverablePath = trim(rs("ImagePath") & "")
			strPath2Location = rs("Path2Location")
			strPath2Description = rs("Path2Description")
			strPath3Location = rs("Path3Location")
			strPath3Description = rs("Path3Description")
			strTDCImagePath = trim(rs("TDCImagePath") & "")
			strArchived = rs("Archived")
			strArchivedPath2 = rs("ArchivedPath2")
			strArchivedPath3 = rs("ArchivedPath3")
			strArchivedTDC = rs("ArchivedTDC")
			strReplicater = rs("Replicater") & ""
			strReplicateOnly = trim(rs("AR") & "")
		end if
	
		rs.close
		cn.Close
		set cn=nothing
		set rs=nothing	
	end if
			
	if RequestID = "" and request("DeliverablePath") <> "" then
		strError=""
		strName = request("DeliverablePath")
		strDeliverablePath = request("DeliverablePath")
	end if
	if RequestID = "" and request("DeliverablePath") = "" and request("TDCImagePath") = "" then 
		strError = "You did not specify a path or ID to download"
	end if
	if request("Instr1")  <> "" then
		strInstruction = "<BR><font color=red>" & request("Instr1") & "</font><BR>"
	else
		strInstruction = ""
	end if
	
	if request("TDCImagePath") <> "" then
		strTDCImagePath = request("TDCImagePath") 
	end if
	
	if request("Path2Location") <> "" then
		strPath2Location = request("Path2Location")
	end if
	
	if request("Path2Description") <> "" then
		strPath2Description = request("Path2Description")
	end if

	if request("Path3Location") <> "" then
		strPath3Location = request("Path3Location")
	end if
	
	if request("Path3Description") <> "" then
		strPath3Description = request("Path3Description")
	end if
	
	if strReplicateOnly = "1" then
		if strReplicater <> "" then
			strInstructions =  "<BR>This deliverable is only available from the developer, the release team, and/or " & strReplicater & " (replicater).<BR>" & strInstructions
		else
			strInstructions =  "<BR>This deliverable is only available from the developer, the release team, and/or the Replicater.<BR>" & strInstructions
		end if
		strDeliverablePath="<font color=red>Files not available online.</font>"
	end if
%>	

<% if strArchived = "1" or strArchivedPath2 = "1" or strArchivedPath3 = "1" or strArchivedTDC = "1" then%>
<HTML>
	<TITLE>Browse Files</TITLE>
	<HEAD>
        <meta http-equiv="X-UA-Compatible" content="IE=8" />
		<LINK rel="stylesheet" type="text/css" href="style/general.css">
	</HEAD>
	<Body>
		This deliverable has been archived, please contact the <A HREF="mailto:psgsoftpaqsupport@hp.com;twn.pdc.nb-releaselab@hp.com">Release Team</A>.
	</body>
</HTML>
<% elseif strReplicateOnly = "1" then%>
<HTML>
	<TITLE>Browse Files</TITLE>
	<HEAD>
        <meta http-equiv="X-UA-Compatible" content="IE=8" />
		<LINK rel="stylesheet" type="text/css" href="style/general.css">
	</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="FileBrowseInfo.asp?Instr1=<%=strInstructions%>&DisplayError=<%=strError%>&DeliverableName=<%=strName%>&DeliverableID=<%=RequestID%>&DeveloperEmail=<%=strDeveloperEmail%>&Developer=<%=strDeveloper%>&DeliverablePath=<%=strDeliverablePath%>">
	</FRAMESET>

<!--	<Body>
		The selected deliverable is only available from the replicater.  Please contact the developer or release team for assistance.
	</body>-->
</HTML>
	
<% elseif strDeliverablePath <> "" or (strDeliverablePath = "" and strTDCImagePath <> "") then%>
<HTML>
	<TITLE>Download</TITLE>
	<HEAD>
        <meta http-equiv="X-UA-Compatible" content="IE=8" />
	</HEAD>
	<%if request("Instr1") <> "" then%>
		<FRAMESET ROWS="150,*" ID=TopWindow >
	<%else%>
		<FRAMESET ROWS="90,*" ID=TopWindow>
	<%end if%>
	<FRAME noresize id="UpperWindow" name="UpperWindow" src="FileBrowseInfo.asp?Instr1=<%=strInstruction%>&DisplayError=<%=strError%>&Path2Location=<%=strPath2Location%>&Path2Description=<%=strPath2Description%>&Path3Location=<%=strPath3Location%>&Path3Description=<%=strPath3Description%>&DeliverableName=<%=strName%>&DeliverableID=<%=RequestID%>&DeveloperEmail=<%=strDeveloperEmail%>&Developer=<%=strDeveloper%>&TDCImagePath=<%=strTDCImagePath%>&DeliverablePath=<%=strDeliverablePath%>">
	<!--<FRAME noresize id="LowerWindow" name="LowerWindow" src="File://<%=strDeliverablePath%>">-->
	<FRAME noresize id="LowerWindow" name="LowerWindow" src="FileBrowseResult.aspx?DisplayError=<%=strError%>&Path2Location=<%=strPath2Location%>&Path3Location=<%=strPath3Location%>&Path3Description=<%=strPath3Description%>&TDCImagePath=<%=strTDCImagePath%>&DeliverablePath=<%=strDeliverablePath%>&DeliverableName=<%=strName%>&DeliverableID=<%=RequestID%>&DeveloperEmail=<%=strDeveloperEmail%>">
	</FRAMESET>

</HTML>
<%elseif trim(strDeliverablePath) = "" and trim(strTDCImagePath) = "" then%>
<HTML>
	<TITLE>Browse Files</TITLE>
	<HEAD>
        <meta http-equiv="X-UA-Compatible" content="IE=8" />
		<LINK rel="stylesheet" type="text/css" href="style/general.css">
	</HEAD>
	<Body>
		No deliverable path was specified by the developer for this file.
	</body>
</HTML>
<%else%>
<HTML>
	<TITLE>Browse Files</TITLE>
	<HEAD>
        <meta http-equiv="X-UA-Compatible" content="IE=8" />
		<LINK rel="stylesheet" type="text/css" href="style/general.css">
	</HEAD>
	<Body>
		The selected deliverable is not released yet so no download page is available yet.
	</body>
</HTML>
<%end if%>

