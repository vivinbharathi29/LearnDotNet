<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<script src="../Scripts/Pulsar2.js"></script>
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "NoUpdate")
        {
            if (isFromPulsar2()) {
                closePulsar2Popup(false);
            }
		    else if (IsFromPulsarPlus()) {
		        ClosePulsarPlusPopup();
		    }
		    else {
		        window.parent.Cancel();
		    }
		}
		else if (txtSuccess.value != "Error")
        {
            if (isFromPulsar2()) {
                closePulsar2Popup(true);
            }
		    else if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.popupCallBack(1);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        var rowid = document.getElementById("txtRowID").value;
		        window.returnValue = txtSuccess.value;
		        window.parent.Close(rowid);
		    }
		}
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update supported products.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update supported products.</font>");
}


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">

<%=">" & request("chkSupport")%><BR>
<%=">" & request("txtProductID")%><BR>
<%=">" & request("txtVersionsLoaded")%><BR><BR><BR>

<%
	dim strAdd
	dim strRemove
	dim IDLoadedArray
	dim IDSelectedArray
	dim strItem
	dim strSuccess
	dim blnError
	dim cn
	dim cm
	dim blnAnyAdded
	dim blnAddRoot
	dim TotalSelected
	dim ProdDelID
	 
	if trim(request("txtVersionsLoaded")) = "" and trim(request("chkSupport")) = "" then
		strSuccess = "NoUpdate"
	else
    
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.Recordset")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		
		
		'Get User
		dim CurrentDomain
		dim CurrentUser
		dim CurrentUserID

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

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") & ""
		else
			CurrentUserID = 0
		end if
		rs.Close
	
		if trim(request("txtRootID")) = "0" then
			blnAddRoot = false
		else
			rs.open "spGetProductDelRootID " &  clng(request("txtProductID")) & "," & clng(request("txtRootID")) & "," & clng(request("txtReleaseID")) ,cn,adOpenStatic
			if rs.EOF and rs.BOF then
				blnAddRoot = true
			else
				blnAddRoot = false
				strProdRootID = rs("ID")
                strProductVersionReleaseID = rs("ProductVersionReleaseID")
			end if
			rs.close
		end if
		
		cn.BeginTrans
		blnError = false


		'Prep Strings
		IDLoadedArray = split(request("txtVersionsLoaded"),",")
		IDSelectedArray = split(request("chkSupport"),",")
	
		TotalSelected = ubound(IDSelectedArray) + 1 + clng(request("txtLockedCount"))
		
		'Find ones to Add
		blnAnyAdded = false
		for each strItem in IDSelectedArray
			if not InArray(IDLoadedArray,strItem) then
				Response.write "<BR>Add:" & strItem  
				blnAnyAdded = true
				
	         	set cm = server.CreateObject("ADODB.Command")
				            
	            cm.ActiveConnection = cn
	            cm.CommandText = "spLinkVersion2Product"
	            cm.CommandType = &H0004
	                
	            Set p = cm.CreateParameter("@ProductVersionID", 3, &H0001)
				p.Value = clng(request("txtProductID"))
	            cm.Parameters.Append p
	                    
	            Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
				p.Value = clng(strItem)
	            cm.Parameters.Append p
		            
                if strProductVersionReleaseID > 0 then
                    Set p = cm.CreateParameter("@ProductVersionReleaseID", 3, &H0001)
				    p.Value = clng(strProductVersionReleaseID)
	                cm.Parameters.Append p
            
                    Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		            p.Value = Currentuser
		            cm.Parameters.Append p

                    Set p = cm.CreateParameter("@UserID",3, &H0001)
			        p.Value = clng(CurrentUserID)
			        cm.Parameters.Append p
                end if

		        cm.Execute recordseffected
	                    
	            Set cm = Nothing

				if recordseffected = 0 or cn.Errors.count > 0 Then
					blnError = true
					exit for
				End If    
                		
			end if
		next
		
		'Add Root To Product if Necessary
		if blnAnyAdded and blnAddRoot then
			Response.Write "<BR>Add Root To Product"
         	set cm = server.CreateObject("ADODB.Command")
				            
            cm.ActiveConnection = cn
            cm.CommandText = "spAddDelRoot2Product"
            cm.CommandType = &H0004
	                
            Set p = cm.CreateParameter("@ProductVersionID", 3, &H0001)
	        p.Value = clng(request("txtProductID"))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
	        p.Value = clng(request("txtRootID"))
	        cm.Parameters.Append p

            if strProductVersionReleaseID > 0 then
                Set p = cm.CreateParameter("@ProductVersionReleaseID", 3, &H0001)
				p.Value = clng(strProductVersionReleaseID)
	            cm.Parameters.Append p

                Set p = cm.CreateParameter("@UserID",3, &H0001)
			    p.Value = clng(CurrentUserID)
			    cm.Parameters.Append p
            
                Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		        p.Value = Currentuser
		        cm.Parameters.Append p        
            end if

		    cm.Execute recordseffected
	                    
	        If recordseffected = 0 or cn.Errors.count > 0 then
				FoundErrors = true
	        End If
	        Set cm = Nothing
	        
			strProdRootID = "" 'No need to do the next section in this case
		end if

		'Automatically Approve this root because some versions were added
		if blnAnyAdded and isnumeric(strProdRootID) and strProdRootID <> 0 and strProdRootID <> "" then
			Response.Write "<BR>Automatically approve Root"

			set cm = server.CreateObject("ADODB.Command")
		
			cm.ActiveConnection = cn

            if clng(request("txtProductDeliverableReleaseID")) = 0 then 
			    cm.CommandText = "spUpdateDeveloperNotificationStatus"
            	cm.CommandType = &H0004
						
			    Set p = cm.CreateParameter("@PDID",3, &H0001)
			    p.Value = clng(strProdRootID)
			    cm.Parameters.Append p
            else 
                cm.CommandText = "spUpdateDeveloperNotificationStatusPulsar"
            	cm.CommandType = &H0004
						
                Set p = cm.CreateParameter("@PDID",3, &H0001)
			    p.Value = clng(strProdRootID)
			    cm.Parameters.Append p

                Set p = cm.CreateParameter("@PDRID",3, &H0001)
			    p.Value = clng(request("txtProductDeliverableReleaseID"))
			    cm.Parameters.Append p
            end if
			
			Set p = cm.CreateParameter("@Status",16, &H0001)
			p.Value = 1
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@TargetNotes", 200, &H0001, 256)
			p.Value = ""
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Type",16, &H0001)
			p.Value = 1
			cm.Parameters.Append p
			
			Set p = cm.CreateParameter("@UserID",3, &H0001)
			p.Value = clng(CurrentUserID)
			cm.Parameters.Append p
					
			cm.Execute recordseffected
			Set cm = Nothing
			
			if recordseffected <> 1 then
				blnError = true
			end if
			
		end if


		'Find ones to remove
		if not blnError then
			for each strItem in IDLoadedArray
				if not InArray(IDSelectedArray,strItem) then
					Response.write "<BR>Remove:" & strItem
					
	          		set cm = server.CreateObject("ADODB.Command")				            
					cm.ActiveConnection = cn

                    if clng(request("txtProductDeliverableReleaseID")) = 0 then 
			            cm.CommandText = "spUnLinkVersionFromProduct"
                        cm.CommandType = &H0004
                        
                        Set p = cm.CreateParameter("@ProdID", 3, &H0001)
					    p.Value = clng(request("txtProductID"))
					    cm.Parameters.Append p
	                    
					    Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
					    p.Value = clng(strItem)
					    cm.Parameters.Append p

                    else 
                        cm.CommandText = "spUnLinkVersionFromProductRelease"
                        cm.CommandType = &H0004
                        
                        Set p = cm.CreateParameter("@ProdID", 3, &H0001)
					    p.Value = clng(request("txtProductID"))
					    cm.Parameters.Append p
	                    
					    Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
					    p.Value = clng(strItem)
					    cm.Parameters.Append p

                        Set p = cm.CreateParameter("@ProductDeliverableReleaseID", 3, &H0001)
					    p.Value = clng(request("txtProductDeliverableReleaseID"))
					    cm.Parameters.Append p
                    end if				
	                					
	                    
					cm.Execute recordseffected
	                    
					Set cm = Nothing

					If recordseffected = 0 or cn.Errors.count > 0 Then
						blnError = true
						exit for
					End If
					
					  
				end if
			next
		end if
	
	
		if blnError then
			cn.RollbackTrans
			strSuccess = "Error"
		else
			cn.CommitTrans
			strSuccess = TotalSelected
		end if
	
	end if
	
	
	
	
	
		function InArray(MyArray,strFind)
			dim strElement
			dim blnFound
			
			blnFound = false
			for each strElement in MyArray
				if trim(strElement) = trim(strFind) then
					blnFound = true
					exit for
				end if
			next
			InArray = blnFound
		end function
	
%>


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input type="hidden" id="txtRowID" name="txtRowID" value="<%=request("txtRowID")%>" />
</BODY>
</HTML>
