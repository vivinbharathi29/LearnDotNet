<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
		{
		    var pulsarplusDivId = document.getElementById('hdnTabName').value;
		    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
		        parent.window.parent.reloadFromPopUp(pulsarplusDivId);
		        parent.window.parent.closeExternalPopup();
		    }
		    else {
		    if (parent.window.parent.document.getElementById('modal_dialog')) {
		        parent.window.parent.modalDialog.cancel(true);
		     }
		    else {
		            window.parent.close();
		        }

		    }
        }
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update the country list.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update the country list.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();">
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

    if strSelected <> "" then
        set cn = server.CreateObject("ADODB.connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
		
		dim CurrentDomain
		dim CurrentUserPartner
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

		cn.BeginTrans    	
	    
        if strAddList <> "" then		
			AddArray = split(strAddList,",")			
			for i = lbound(AddArray) to ubound(AddArray)
				if trim(AddArray(i)) <> "" then

					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
		
					cm.CommandText = "usp_InsertProdBrandCountry"	

					Set p = cm.CreateParameter("@p_CountryID", adInteger)
					p.Value = AddArray(i)
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@p_ProductBrandID", adInteger)
					p.Value = request("ProdBrandID")
					cm.Parameters.Append p
					
					Set p = cm.CreateParameter("@p_DcrID", adInteger)
					If Request("cboDcr") <> "" Then
						p.Value = request("cboDcr")
					End If
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

		if (not FoundErrors) and strRemoveList <> "" then
		
			RemoveArray = split(strRemoveList,",")
			
			for i = lbound(RemoveArray) to ubound(RemoveArray)
				if trim(RemoveArray(i)) <> "" then

					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
		
					cm.CommandText = "usp_DeleteProdBrandCountry"	

					Set p = cm.CreateParameter("@p_CountryID", adInteger)
					p.Value = RemoveArray(i)
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@p_ProductBrandID", adInteger)
					p.Value = request("ProdBrandID")
					cm.Parameters.Append p
					
					Set p = cm.CreateParameter("@p_DcrID", adInteger)
					If Request("cboDcr") <> "" Then
						p.Value = request("cboDcr")
					End If
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
       if (not FoundErrors) and cint(request("IsPulsarProduct")) = 1 then 'save the updated releases for all the selected countries
            dim intCountryID
            dim strID
            dim strAddReleaseList : strAddReleaseList = ""
            dim strRemoveReleaseList : strRemoveReleaseList = ""
            dim rsReleases
            dim strReleases : strReleases = ""
            dim ReleaseArray
               
            set rsReleases = server.CreateObject("ADODB.recordset")
            set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "usp_Product_GetProductReleases"

		    Set p = cm.CreateParameter("@p_intProductVersionID",adInteger)
		    p.Value = request("ProdID")
		    cm.Parameters.Append p		    

            rsReleases.CursorType = adOpenForwardOnly
		    rsReleases.LockType=AdLockReadOnly
		    Set rsReleases = cm.Execute 
            set cm = nothing 
          
            Do while not rsReleases.EOF
                strReleases = strReleases & "," & rsReleases("ReleaseID")
                rsReleases.MoveNext
            Loop
            
            if strReleases <> "" then
		        strReleases  = mid(strReleases,2)
	        end if	
            
            rsReleases.Close   
            ReleaseArray = split(strReleases,",")

            strAddReleaseList = "<?xml version='1.0' encoding='iso-8859-1' ?><AddReleaseList>"
            strRemoveReleaseList = "<?xml version='1.0' encoding='iso-8859-1' ?><RemoveReleaseList>"
            for i = lbound(SelectedArray) to ubound(SelectedArray)
                intCountryID = trim(SelectedArray(i))
                for j = lbound(ReleaseArray) to ubound(ReleaseArray)
                    strID = intCountryID & "-" & ReleaseArray(j) 
                    'check if the release checkbox value has changed for the country
                    if (cint(request("txtnew" & strID)) < 2) then'only update the localization for the fully checked or unchecked releases not for partially checked ones
                        if ((cint(request("txtold" & strID)) <> cint(request("txtnew" & strID))) and (cint(request("txtnew" & strID))) = 1) then 'add to the add list 
                            strAddReleaseList = strAddReleaseList & "<CountryRelease><CountryID>" & intCountryID & "</CountryID>" & "<ReleaseID>" & ReleaseArray(j) & "</ReleaseID></CountryRelease>"
                        elseif ((cint(request("txtold" & strID)) <> cint(request("txtnew" & strID))) and (cint(request("txtnew" & strID))) = 0) then 'add to the delete list
                            strRemoveReleaseList = strRemoveReleaseList & "<CountryRelease><CountryID>" & intCountryID & "</CountryID>" & "<ReleaseID>" & ReleaseArray(j) & "</ReleaseID></CountryRelease>"
                        end if
                    end if    
                next
            next     
            strAddReleaseList = strAddReleaseList & "</AddReleaseList>"
            strRemoveReleaseList = strRemoveReleaseList & "</RemoveReleaseList>"
    	   
            set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
		
			cm.CommandText = "usp_Product_ProdBrandCountry_UpdateReleases"	
            
            Set p = cm.CreateParameter("@p_intProductBrandID", adInteger)
			p.Value = request("ProdBrandID")
			cm.Parameters.Append p
					
			Set p = cm.CreateParameter("@p_xmlAddReleaseList", adLongVarChar, adParamInput, len(strAddReleaseList))
			p.Value = strAddReleaseList		
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_xmlRemoveReleaseList", adLongVarChar, adParamInput, len(strRemoveReleaseList))
			p.Value = strRemoveReleaseList		
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@p_chrUserName", adVarChar, adParamInput, 250)
			p.Value = UserName
			cm.Parameters.Append p

			cm.Execute rowschanged

			if cn.Errors.count > 1 then
				FoundErrors = true   
			end if
		
			'set cm = nothing
        end if
		if not FoundErrors then
			cn.CommitTrans
		else
			cn.RollbackTrans
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
    <input type="hidden" id="hdnTabName" class="hdnTabClass" value="<%=Request("pulsarplusDivId")%>" />
</BODY>
</HTML>



