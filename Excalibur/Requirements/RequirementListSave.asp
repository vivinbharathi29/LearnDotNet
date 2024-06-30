<%@  language="VBScript" %>
<!-- #include file = "../includes/noaccess.inc" -->
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" language="javascript">
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
                // For Reload PulsarPlusPmView Tab
                parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

                // For Closing current popup
                parent.window.parent.closeExternalPopup();
            }
            else {

                var iframeName = parent.window.name;
                if (iframeName != '') {
                    parent.window.parent.CloseRequirementsDialog(1);
                } else {
                    if (parent.window.parent.document.getElementById('modal_dialog')) {
                        //save value and return to parent page: ---
                        parent.window.parent.modalDialog.cancel(true);
                    } else {
                        window.returnValue = 1;
                        window.parent.close();
                    }
                }
            }
        }
        else {
            document.write("<BR><font size=2 face=verdana>Unable to update the requirement list.</font>");
        }
    }
    else {
        document.write("<BR><font size=2 face=verdana>Unable to update the requirement list.</font>");
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

	
	strSelected = ", " & request("chkSelected") & ","
	strTag = ", " & request("chkTag") & ","
	SelectedArray = split(request("chkSelected"),",")
	TagArray = split(request("chkTag"),",")
	
	strAddList = ""
	strRemoveList = ""
	
	for i = lbound(SelectedArray) to ubound(SelectedArray) 
		if instr(strTag,", " & trim(SelectedArray(i)) & ",") = 0 then
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


	FoundErrors = false	
	
	if strAddList <> "" or strRemoveList <> "" then


		set cn = server.CreateObject("ADODB.connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		
		cn.BeginTrans

	
		if strAddList <> "" then
		
			AddArray = split(strAddList,",")
			
			for i = lbound(AddArray) to ubound(AddArray)
				if trim(AddArray(i)) <> "" then

					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
		
					cm.CommandText = "spAddRequirement2ProductWeb"	

					Set p = cm.CreateParameter("@RequirementID", 3,  &H0001)
					p.Value = AddArray(i)
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
					p.Value = request("txtID")
					cm.Parameters.Append p
					
					cm.Execute rowschanged

					if rowschanged <> 1 then
						FoundErrors = true
					end if
		
					set cm = nothing
				end if
			next
		end if

		if  (not FoundErrors) and strRemoveList <> "" then
		
			RemoveArray = split(strRemoveList,",")
			
			for i = lbound(RemoveArray) to ubound(RemoveArray)
				if trim(RemoveArray(i)) <> "" then

					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
		
					cm.CommandText = "spRemoveRequirementFromProduct"	

					Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
					p.Value = request("txtID")
					cm.Parameters.Append p
					
					Set p = cm.CreateParameter("@RequirementID", 3,  &H0001)
					p.Value = RemoveArray(i)
					cm.Parameters.Append p

					cm.Execute rowschanged

					if cn.Errors.count > 1 then
						FoundErrors = true
					end if
		
					set cm = nothing
				end if
			next
		end if


		if not FoundErrors then
			cn.committrans
		else
			cn.rollback
		end if
		cn.close
		set cn = nothing
	end if


	if FoundErrors then
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
	else
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
	end if

    %>
    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</body>
</html>
