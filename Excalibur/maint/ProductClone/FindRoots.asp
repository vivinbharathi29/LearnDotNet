<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
	

function cmdNext_onclick() {
	frmMain.submit();
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
body
{
	font-faily: Verdana;
	font-size: xx-small;
}
TD
{
	font-faily: Verdana;
	font-size: xx-small;
}

</STYLE>
<BODY>
<form id=frmMain action=FindVersions.asp method=post>
<%

	dim strIDList
	dim strLastReq
	strLastReq = ""
	strIDList = ""

	if request("lstTarget") = "" or request("lstSource") = "" then

	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		cn.CommandTimeout = 5400
		set rs = server.CreateObject("ADODB.recordset")
	
		dim strSource
		dim strTarget
		dim strSourceName
		dim strTargetName
	
		strSource = clng(request("lstSource"))
		strTarget = clng(request("lstTarget"))
		
		rs.Open "spGetProductVersion " & strSource,cn,adOpenStatic
		if rs.EOF or rs.bof then
			strSourceName = ""
		else
			strSourceName = rs("DotsName")
		end if
		rs.Close

		rs.Open "spGetProductVersion " & strTarget,cn,adOpenStatic
		if rs.EOF or rs.bof then
			strTargetName = ""
		else
			strTargetName = rs("DotsName")
		end if
		rs.Close
		
		Response.Write "<h3>Clone From " & strSourceName & " to " & strTargetName & ".</h3>"

		Response.Write "<b>Step 2: Choose Destination Requirement.</b><BR>"
			Response.Write "<INPUT checked type=""radio"" name=optReq ID=optCloneReq value=1> Copy to matching requirement - Add Requirement if missing.<BR>"
			Response.Write "<INPUT type=""radio"" name=optReq ID=optCloneReqTBD  value=2> Copy to matching requirement - Add to TBD if missing.<BR>"
			Response.Write "<INPUT height=16 width=16 type=""radio"" name=optReq ID=optCloneTBD  value=3> Copy to TBD only.<BR><BR>"
		
		Response.Write "<b>Step 3: Choose Roots to Add.</b><BR>"
		
	
		rs.Open "spCloneFindMissingRoots " & strSource & "," & strTarget,cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			Response.Write "<TABLE>"
		end if
		do while not rs.EOF 
			if strLastReq <> "" and strLastReq <> rs("Requirement") then
				Response.Write "<BR>" & strLastReq & ":" & strIDList
				strIDList = ""
			end if
			strLastReq = rs("Requirement")
			strIDList = strIDList & "," & rs("ID")
			Response.Write "<TR><TD><INPUT height=16 width=16 type=""checkbox"" checked id=chkID name=chkID value=""" & rs("ID") & """></TD><TD>" & rs("Name") & "</TD><TD>" & rs("Requirement") & "</TD></TR>"
			rs.MoveNext
		loop
		Response.Write "<BR>" & strLastReq & ":" & strIDList
		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>"
		end if
		rs.Close	
	
		set rs = nothing
		cn.Close
		set cn = nothing


	end if	'Verify input params

%>
<BR>
<INPUT type="button" value="Next" id=cmdNext name=cmdNext LANGUAGE=javascript onclick="return cmdNext_onclick()"> <BR>
<font size=1 face=verdana color=red>Note: The selected root deliverables will be copied to the destination product when you click this button.  There is no "undo" function for this operation.</font>


<INPUT type="hidden" id=lstSource name=lstSource value="<%=strSource%>">
<INPUT type="hidden" id=lstSourceName name=lstSourceName value="<%=strSourceName%>"><BR>
<INPUT type="hidden" id=lstTarget name=lstTarget value="<%=strTarget%>">
<INPUT type="hidden" id=lstTargetName name=lstTargetName value="<%=strTargetName%>">
</form>
</BODY>
</HTML>
