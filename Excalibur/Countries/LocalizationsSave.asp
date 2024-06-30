<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->


<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	    if (txtSuccess.value == "1") {
	        var pulsarplusDivId = document.getElementById('hdnTabName');
	        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
	            // For Reload PulsarPlusPmView Tab
	            parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

	            // For Closing current popup
	            parent.window.parent.closeExternalPopup();
	        }
	        else {

	            var iframeName = parent.window.name;
	            if (iframeName != '') {
	                parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
	            } else {
	                window.returnValue = txtSummary.value;
	                window.parent.close();
	            }
	        }
	    }
//		else
//			document.write ("<BR><font size=2 face=verdana>Unable to update the country list.</font>");
		}
//	else
//		document.write ("<BR><font size=2 face=verdana>Unable to update the country list.</font>");
}

//-->
</script>
</head>
<body onload="return window_onload();">
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
	dim DcrID

	DcrID = Trim(Request.Form("cboDcr"))
	If DcrID = "" Then
		DcrID = NULL
	End If
	
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	Set rs = cm.Execute 

	set cm=nothing


	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=1"
	else
		UserName = rs("Name")
	end if 
	rs.Close

	
	strSelected = ", " & request("chkSelected") & ","
	strTag = ", " & request("chkTag") & ","
	SelectedArray = split(request("chkSelected"),",")
	TagArray = split(request("chkTag"),",")
	
	strAddList = ""
	strRemoveList = ""
	strUpdateList = ""
	
	for i = lbound(SelectedArray) to ubound(SelectedArray) 
		if instr(strTag,", " & trim(SelectedArray(i)) & ",") = 0 then
			strAddList = strAddList & "," & trim(SelectedArray(i))
		end if
	next

	for i = lbound(SelectedArray) to ubound(SelectedArray) 
		if instr(strTag,", " & trim(SelectedArray(i)) & ",") > 0 then
			strUpdateList = strUpdateList & "," & trim(SelectedArray(i))
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

	If strUpdateList <> "" Then
		strUpdateList = mid(strUpdateList, 2)
	End If

	FoundErrors = false	
	
	cn.BeginTrans

	if strAddList <> "" then
		
		AddArray = split(strAddList,",")
			
		for i = lbound(AddArray) to ubound(AddArray)
			
			If trim(AddArray(i)) <> "" Then

				PwrCord = Trim(Request.Form("PwrCord" & AddArray(i)))
				Keyboard = Trim(Request.Form("Kbd" & AddArray(i)))
				KWL = Trim(Request.Form("KWL" & AddArray(i)))
				DocKit = Trim(Request.Form("DocKit" & AddArray(i)))
				Media = Trim(Request.Form("Media" & AddArray(i)))
				
				If PwrCord = "" Then
					PwrCord = NULL
				End If
				
				If Keyboard = "" Then
					Keyboard = NULL
				End If
				
				If KWL = "" Then
					KWL = NULL
				End If
				
				If DocKit = "" Then
					DocKit = NULL
				End If
				
				If Media = "" Then
					Media = NULL
				End If

				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType = adCmdStoredProc
				cm.ActiveConnection = cn
		
				cm.CommandText = "usp_InsertProdBrandCountryLocalization"	

				Set p = cm.CreateParameter("@p_ProductBrandCountryID", adInteger)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@p_LocalizationID", adInteger)
				p.Value = AddArray(i)
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_UserName", adVarChar, adParamInput, 20)
				p.Value = UserName
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_DcrID", adInteger)
				p.Value = DcrID
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_PowerCord", adVarChar, adParamInput, 20)
				p.Value = PwrCord
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_Keyboard", adVarChar, adParamInput, 15)
				p.Value = Keyboard
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_KWL", adVarChar, adParamInput, 7)
				p.Value = KWL
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_DocKit", adVarChar, adParamInput, 30)
				p.Value = DocKit
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_Media", adVarChar, adParamInput, 25)
				p.Value = Media
				cm.Parameters.Append p
					
				cm.Execute rowschanged

				if rowschanged <> 1 then
					FoundErrors = true
				end if
		
				set cm = nothing
			end if
		next
	end if

	if (not FoundErrors) and strRemoveList <> "" then
		
		RemoveArray = split(strRemoveList,",")
			
		for i = lbound(RemoveArray) to ubound(RemoveArray)
			if trim(RemoveArray(i)) <> "" then

				set cm = server.CreateObject("ADODB.Command")
				cm.CommandType =  &H0004
				cm.ActiveConnection = cn
		
				cm.CommandText = "usp_DeleteProdBrandCountryLocalization"	

				Set p = cm.CreateParameter("@p_ProdBrandCountryLocalizationID", adInteger)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@p_ProductBrandCountryID", adInteger)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@p_LocaliaztionID", adInteger)
				p.Value = RemoveArray(i)
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_DcrID", adInteger)
				p.Value = DcrID
				cm.Parameters.Append p
					
				Set p = cm.CreateParameter("@p_UserName", adVarChar, adParamInput, 20)
				p.Value = UserName
				cm.Parameters.Append p
					
				cm.Execute rowschanged

				if cn.Errors.count > 1 then
					FoundErrors = true
				end if
		
				set cm = nothing
			end if
		next
	end if

	if (Not FoundErrors) And strUpdateList <> "" then
		
		UpdateArray = split(strUpdateList,",")
			
		for i = lbound(UpdateArray) to ubound(UpdateArray)
            
			If (trim(UpdateArray(i)) <> "") Then
				
				PwrCord = Trim(Request.Form("PwrCord" & UpdateArray(i)))
				hidPwrCord = Request.Form("hidPwrCord" & UpdateArray(i))
				Keyboard = Trim(Request.Form("Kbd" & UpdateArray(i)))
				hidKeyboard = Request.Form("hidKbd" & UpdateArray(i))
				KWL = Trim(Request.Form("KWL" & UpdateArray(i)))
				hidKWL = Request.Form("hidKWL" & UpdateArray(i))
				DocKit = Trim(Request.Form("DocKit" & UpdateArray(i)))
				hidDocKit = Request.Form("hidDocKit" & UpdateArray(i))
				Media = Trim(Request.Form("Media" & UpdateArray(i)))
				hidMedia = Request.Form("hidMedia" & UpdateArray(i))
				
				If PwrCord <> hidPwrCord Or Keyboard <> hidKeyboard Or _
					KWL <> hidKWL Or DocKit <> hidDocKit Or Media <> hidMedia Then
					
					If PwrCord = "" Then
						PwrCord = NULL
					End If
					
					If Keyboard = "" Then
						Keyboard = NULL
					End If
					
					If KWL = "" Then
						KWL = NULL
					End If
					
					If DocKit = "" Then
						DocKit = NULL
					End If
					
					If Media = "" Then
						Media = NULL
					End If
					
					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
			
					cm.CommandText = "usp_UpdateProdBrandCountryLocalization"	

					Set p = cm.CreateParameter("@p_ProdBrandCountryLocalizationID", adInteger)
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@p_ProductBrandCountryID", adInteger)
					p.Value = request("txtID")
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@p_LocalizationID", adInteger)
					p.Value = UpdateArray(i)
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@p_UserName", adVarChar, adParamInput, 20)
					p.Value = UserName
					cm.Parameters.Append p
					
					Set p = cm.CreateParameter("@p_DcrID", adInteger)
					p.Value = DcrID
					cm.Parameters.Append p
					
					Set p = cm.CreateParameter("@p_PowerCord", adVarChar, adParamInput, 20)
					p.Value = PwrCord
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@p_Keyboard", adVarChar, adParamInput, 15)
					p.Value = Keyboard
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@p_KWL", adVarChar, adParamInput, 7)
					p.Value = KWL
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@p_DocKit", adVarChar, adParamInput, 30)
					p.Value = DocKit
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@p_Media", adVarChar, adParamInput, 25)
					p.Value = Media
					cm.Parameters.Append p
					
					cm.Execute rowschanged

					if rowschanged = 0 then
						FoundErrors = true
					end if
		
					set cm = nothing
				End If
			end if		
		next
	end if
	if not FoundErrors then
		cn.CommitTrans
	else
		cn.RollbackTrans
	end if
	cn.close
	set cn = nothing
	'end if


	if FoundErrors then
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
	else
		Response.Write "<INPUT type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
	end if

%>

<input type="text" id="txtSummary" name="txtSummary" value="<%=strLocalization%>" />
    <input type="hidden" id="hdnTabName" name="hdnTabName" value="<%= Request("pulsarplusDivId")%>" />
</body>
</html>



