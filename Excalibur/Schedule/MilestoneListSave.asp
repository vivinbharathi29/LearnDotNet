<%@  language="VBScript" %>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->

<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" language="javascript">
<!--

    function window_onload() {

        if (typeof (txtSuccess) != "undefined") {

            if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
                // For Reload PulsarPlusPmView Tab
                parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

                // For Closing current popup
                parent.window.parent.closeExternalPopup();
            }
            else {

                if (txtSuccess.value == "1") {
                    if (parent.window.parent.document.getElementById('modal_dialog')) {
                        parent.window.parent.modalDialog.cancel(true);
                    } else {
                        window.returnValue = 1;
                        window.parent.close();
                    }
                } else {
                    document.write("<BR><font size=2 face=verdana>Unable to update the schedule.</font>");
                }
            }
        } else {
            document.write("<BR><font size=2 face=verdana>Unable to update the schedule.</font>");
        }
        
    }

    //-->
    </script>
</head>
<body language="javascript" onload="return window_onload();">
    <%

	dim strSelected
	dim strTag
	dim SelectedArray
	dim TagArray
	dim i
	dim strAddList
	dim strRemoveList
	dim AddArray
	dim RemoveArray
	dim cn
	dim cm
	dim RowsChanged

'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim m_IsSysAdmin
	Dim m_IsProgramManager
	Dim m_IsSysEngProgramManager
    Dim m_IsSEPMProductsEditor
	Dim m_IsSysTeamLead
	Dim m_EditModeOn
	
	m_EditModeOn = False
	
	Dim Security
	Dim sUserFullName
	
	Set Security = New ExcaliburSecurity
	
	
	m_IsSysAdmin = Security.IsSysAdmin()
'
' Debug Section
'
'	If Security.CurrentUserID = 1396 Then
'		m_IsSysAdmin = False
'		Security.CurrentUserID = 1288
'		Response.Write Security.CurrentUserID
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSysEngProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSystemTeamLead(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If
	
	m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
	m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	sUserFullName = Security.CurrentUser()
	
	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

	strSelected = ", " & request("chkSelected") & ","
	strTag = ", " & request("chkTag") & ","
	SelectedArray = split(request("chkSelected"),",")
	TagArray = split(request("chkTag"),",")
	Response.Write strSelected & "<br>"
	Response.Write "<hr>"
	Response.Write strTag & "<br>"
	Response.Write "<hr>"
	
	strAddList = ""
	strRemoveList = ""
	
	for i = lbound(SelectedArray) to ubound(SelectedArray) 
		if instr(strTag,", " & trim(SelectedArray(i)) & ",") <> 0 then
			strAddList = strAddList & "," & trim(SelectedArray(i))
		end if
	next

	for i = lbound(TagArray) to ubound(TagArray) 
		if instr(strSelected,", " & trim(TagArray(i)) & ",") = 0 then
			strRemoveList = strRemoveList & "," & trim(TagArray(i))
		end if
	next

	if strAddList <> "" then
		strAddList  = mid(strAddList,2)
	end if	

	if strRemoveList <> "" then
		strRemoveList  = mid(strRemoveList,2)
	end if	

	Response.Write strAddList & "<BR>"
	Response.Write "<hr>"
	Response.Write strRemoveList & "<br>"
	Response.Write "<hr>"

	FoundErrors = false	
	
	if strAddList <> "" or strRemoveList <> "" then


		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		
		cn.BeginTrans

		if strAddList <> "" then
			AddArray = split(strAddList,",")
			
			for i = lbound(AddArray) to ubound(AddArray)
				if trim(AddArray(i)) <> "" then

					Set cmd = dw.CreateCommandSP(cn, "usp_UpdateScheduleData")
					dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, AddArray(i)
					dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
					dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, "Y"
					rowschanged = dw.ExecuteNonQuery(cmd)
					Response.Write rowschanged & "<BR>"
					
					Set cmd = nothing

'					if rowschanged <> 1 then
'						FoundErrors = true
'					end if
		
				end if
			next
		end if

		if (not FoundErrors) and strRemoveList <> "" then
			RemoveArray = split(strRemoveList,",")
			
			for i = lbound(RemoveArray) to ubound(RemoveArray)
				if trim(RemoveArray(i)) <> "" then
				
					Set cmd = dw.CreateCommandSP(cn, "usp_UpdateScheduleData")
					dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, RemoveArray(i)
					dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
					dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, "N"
					rowschanged = dw.ExecuteNonQuery(cmd)
					Response.Write rowschanged & "<BR>"

					Set cmd = nothing

'					if rowschanged <> 1 then
'						FoundErrors = true
'					end if
		
				end if
			next
		end if

		if not FoundErrors then
			cn.CommitTrans
		else
			cn.RollbackTrans
		end if
		set cn = nothing
	end if


	if FoundErrors then
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
	else
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
	end if

    %>
    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
</body>
</html>



