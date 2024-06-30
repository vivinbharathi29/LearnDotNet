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

</STYLE>
<BODY>
<form action=FindVersionsToRemove.asp method=post ID=frmMain>
<%
	dim strIDList

	if request("lstTarget") = "" or request("lstSource") = "" then

	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		cn.CommandTimeout = 5400
		set rs = server.CreateObject("ADODB.recordset")
	
	%><!-- #include file="incSaveRootAdd.asp" --><%
	
		dim strSource
		dim strTarget
		dim strSourceName
		dim strTargetName
		dim strVersion
	
		strSource = clng(request("lstSource"))
		strTarget = clng(request("lstTarget"))
		strSourceName = request("lstSourceName")
		strTargetName = request("lstTargetName")

		Response.Write "<h3>Clone From " & strSourceName & " to " & strTargetName & ".</h3>"

		Response.Write "<b>Step 4: Choose Versions to Add or Target.</b><BR>"
		
	
		rs.Open "spCloneFindMissingVersions " & strSource & "," & strTarget,cn,adOpenStatic
		do while not rs.EOF 
			strVersion = rs("name") & " [" & rs("Version")
			if trim(rs("Revision") & "") <> "" then
				strversion = strVersion & "," & rs("Revision")
			end if
			if trim(rs("Pass") & "") <> "" then
				strversion = strVersion & "," & rs("Pass")
			end if
			strversion = strVersion & "]"
			Response.Write "<INPUT height=16 width=16 type=""checkbox"" checked id=chkID name=chkID value=""" & rs("ID") & """>" & strVersion & "<BR>"
			strIDList = strIDList & "," & rs("ID")
			rs.MoveNext
		loop
		rs.Close	
	
		if strIDList <> "" then
			strIDList = mid(strIDList,2)
		end if 
	
		set rs = nothing
		cn.Close
		set cn = nothing


	end if	'Verify input params

	Response.Write "<BR><BR>" & strIDList

%>
<BR>
<INPUT type="button" value="Next" id=cmdNext name=cmdNext LANGUAGE=javascript onclick="return cmdNext_onclick()"> <BR>
<font size=1 face=verdana color=red>Note: The selected deliverable versions will be copied to the destination product when you click this button.  There is no "undo" function for this operation.</font>


<INPUT type="hidden" id=lstSource name=lstSource value="<%=strSource%>">
<INPUT type="hidden" id=lstSourceName name=lstSourceName value="<%=request("lstSourceName")%>"><BR>
<INPUT type="hidden" id=lstTarget name=lstTarget value="<%=strTarget%>">
<INPUT type="hidden" id=lstTargetName name=lstTargetName value="<%=request("lstTargetName")%>">
</form>
</BODY>
</HTML>
