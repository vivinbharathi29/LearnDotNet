<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

	%>
<!-- #include file= "../includes/EmailQueue.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->



<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    if (typeof (txtSuccess) != "undefined") {
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Reload PulsarPlusPmView Tab
            parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (txtSuccess.value != "0") {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = txtSuccess.value;
                    window.parent.opener = 'X';
                    window.parent.open('', '_parent', '')
                    window.parent.close();
                }
            }
        }
    }
//-->
</SCRIPT>

</HEAD>

<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Processing....
<%


	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 


	dim cn
	dim rs
	dim strSQL
	dim strFirstName
	dim strDeliverable
	dim strMSG
	dim DelCount
	dim strLastDev
	dim strLastDevMan
	dim strDelText
	dim strConfirmMSG
	dim strProductName
	dim strPMEmail
	dim strPMName
	dim strOS
	'dim strBrands
	dim strFullName
	dim strNotify
	dim i
	dim SeriesArray
	dim strSeriesName
	dim strVersionNote
	dim strVersion
	dim VersionArray
	dim strImage
	dim ImageArray
    dim strDelList 
    dim strExecptions 
	
	strSuccess = "1"
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	dim CurrentUser
	dim CurrentUserEmail
	dim CurrentUserName
	dim CurrentUserID

	'Get User
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

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing
	
	CurrentUserID = ""
	if not (rs.EOF and rs.BOF) then
		CurrentUserEmail = rs("Email") & ""
		CurrentUserName = rs("Name") & ""
	    CurrentUserID = rs("ID") & ""
	end if
	rs.Close



    if currentuserid = "" then
        response.Write "You are not authorized to use this page."
    else
		
		cn.begintrans
		
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		cm.CommandText = "spAddProductRTM"	

        Set p = cm.CreateParameter("ProductRTMID", 3,  &H0001)
		p.Value = clng("0")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("ProductVersionID", 3,  &H0001)
		p.Value = clng(request("txtProductID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Title", 200,  &H0001,120)
		p.Value = left(request("txtRTMName"),120)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RTMDate", 135, &H0001)
		if trim(request("txtRTMDate")) <> "" and isdate(trim(request("txtRTMDate"))) then
			p.Value = CDate(trim(request("txtRTMDate")))
		else
			p.value = null
		end if
		cm.Parameters.Append p
    
		Set p = cm.CreateParameter("@Comments", 200,  &H0001,2147483647)
		p.Value = request("txtRTMComments")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@BIOSComments", 200,  &H0001,2147483647)
		p.Value = trim(request("txtBIOSComments") & "")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RestoreComments", 200,  &H0001,2147483647)
		p.Value = trim(request("txtRestoreComments") & "")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ImageComments", 200,  &H0001,2147483647)
		p.Value = trim(request("txtImageComments") & "")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PatchComments", 200,  &H0001,2147483647)
		p.Value = trim(request("txtPatchComments") & "")
		cm.Parameters.Append p

    	Set p = cm.CreateParameter("@FWComments", 200,  &H0001,2147483647)
		p.Value = trim(request("txtFWComments") & "")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Attachment1", 200,  &H0001,300)
		p.Value = left(trim(request("txtAttachmentPath1") & ""),300)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("SubmittedByID", 3,  &H0001)
		p.Value = clng(Currentuserid)
		cm.Parameters.Append p

        Set p = cm.CreateParameter("@RTMAsDraft", 3,  &H0001)	
		p.Value = cbool("0")		
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
    	cm.Parameters.Append p
	
		cm.Execute RowsEffected

        if RowsEffected <> 1 then
            strSuccess = "0"
        else
            strSuccess = cm("@NewID")
        end if
        set cm = nothing
        
        'Link Alerts to RTM
        if strSuccess <> "0" then
            CommentArray = Array(request("txtBuildLevelComments"),request("txtDistributionComments"),request("txtCertificationComments"),request("txtWorkflowComments"),request("txtAvailabilityComments"),request("txtDeveloperComments"),request("txtRootComments"),request("txtOTSPrimaryComments"),request("txtOTSRelatedComments"))
            for i= 1 to 8  
    		    set cm = server.CreateObject("ADODB.Command")
	    	    cm.CommandType =  &H0004
		        cm.ActiveConnection = cn
    		
		        cm.CommandText = "spUpdateProductRTMAlert"	
    
    
		        Set p = cm.CreateParameter("@ProductRTMID", 3,  &H0001)
		        p.Value = clng(strSuccess)
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@SectionID", 3,  &H0001)
		        p.Value = i
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@UserID", 3,  &H0001)
		        p.Value = CurrentUserID
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		        p.Value = clng(request("txtProductID"))
		        cm.Parameters.Append p

		        Set p = cm.CreateParameter("@Comments", 200,  &H0001,500)
		        p.Value = left(CommentArray(i-1),500)
		        cm.Parameters.Append p

                Set p = cm.CreateParameter("@UpdateRTMAlert", 3,  &H0001)
                p.Value =1
		        cm.Parameters.Append p

        		cm.Execute RowsEffected

                if RowsEffected <> 1 then
                    strSuccess = "0"
                    exit for
                end if
                set cm = nothing
            next
        end if

        'Link Patches to RTM
        if strSuccess <> "0" and trim(request("chkPatch")) = "1" then
            VersionArray = split(request("chkPatchList"),",")
            for each strVersion in VersionArray

    		    set cm = server.CreateObject("ADODB.Command")
	    	    cm.CommandType =  &H0004
		        cm.ActiveConnection = cn
    		
		        cm.CommandText = "spAddProductRTMDeliverable"	
    
		        Set p = cm.CreateParameter("@ProductRTMID", 3,  &H0001)
		        p.Value = clng(strSuccess)
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@DeliverableVersionID", 3,  &H0001)
		        p.Value = clng(trim(strVersion))
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@TypeID", 3,  &H0001)
		        p.Value = 3
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@Details", 200,  &H0001,300)
	            p.Value = ""
		        cm.Parameters.Append p

        		cm.Execute RowsEffected

                if RowsEffected <> 1 then
                    strSuccess = "0"
                    exit for
                end if
                set cm = nothing
                
            next

        end if
        
        'Link System BIOS to RTM
        if strSuccess <> "0" and trim(request("chkBIOS")) = "1" then
            VersionArray = split(request("chkBIOSList"),",")
            for each strVersion in VersionArray

    		    set cm = server.CreateObject("ADODB.Command")
	    	    cm.CommandType =  &H0004
		        cm.ActiveConnection = cn
    		
		        cm.CommandText = "spAddProductRTMDeliverable"	
    
		        Set p = cm.CreateParameter("@ProductRTMID", 3,  &H0001)
		        p.Value = clng(strSuccess)
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@DeliverableVersionID", 3,  &H0001)
		        p.Value = clng(trim(strVersion))
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@TypeID", 3,  &H0001)
		        p.Value = 1
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@Details", 200,  &H0001,300)
                if request("optPhaseIn") = "0" then
		            p.Value = "Immediate (Rework All Units)"
                elseif request("optPhaseIn") = "1" then
		            p.Value = "Phase-in"
                else
		            p.Value = "Web Release Only"
		        end if
		        cm.Parameters.Append p

        		cm.Execute RowsEffected

                if RowsEffected <> 1 then
                    strSuccess = "0"
                    exit for
                end if
                set cm = nothing
                
            next

        end if


        'Link Restore Solution to RTM
        if strSuccess <> "0" and trim(request("chkRestore")) = "1" then
            VersionArray = split(request("chkRestoreList"),",")
            for each strVersion in VersionArray

    		    set cm = server.CreateObject("ADODB.Command")
	    	    cm.CommandType =  &H0004
		        cm.ActiveConnection = cn
    		
		        cm.CommandText = "spAddProductRTMDeliverable"	
    
		        Set p = cm.CreateParameter("@ProductRTMID", 3,  &H0001)
		        p.Value = clng(strSuccess)
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@DeliverableVersionID", 3,  &H0001)
		        p.Value = clng(trim(strVersion))
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@TypeID", 3,  &H0001)
		        p.Value = 2
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@Details", 200,  &H0001,300)
	            p.Value = ""
		        cm.Parameters.Append p

        		cm.Execute RowsEffected

                if RowsEffected <> 1 then
                    strSuccess = "0"
                    exit for
                end if
                set cm = nothing
                
            next

        end if
        
        
        'Link Images to RTM
        if strSuccess <> "0" and trim(request("chkImages")) = "1" then
            ImageArray = split(request("chkImage"),",")
            for each strImage in ImageArray

    		    set cm = server.CreateObject("ADODB.Command")
	    	    cm.CommandType =  &H0004
		        cm.ActiveConnection = cn
    		
		        cm.CommandText = "spAddProductRTMImage"	
    
		        Set p = cm.CreateParameter("@ProductRTMID", 3,  &H0001)
		        p.Value = clng(strSuccess)
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
		        p.Value = clng(trim(strImage))
		        cm.Parameters.Append p

        		Set p = cm.CreateParameter("@RTMDate", 135, &H0001)
        		if trim(request("txtRTMDate")) <> "" and isdate(trim(request("txtRTMDate"))) then
		        	p.Value = CDate(trim(request("txtRTMDate")))
		        else
			        p.value = null
		        end if
		        cm.Parameters.Append p
    
        		cm.Execute RowsEffected

                if RowsEffected <> 1 then
                    strSuccess = "0"
                    exit for
                end if
                set cm = nothing
                
                'Lock Deliverable List
			    rs.Open "spListDeliverablesInImage " & clng(trim(strImage)) & ",0" ,cn,adOpenStatic 
			    strDelList = ""
			    strExecptions = ""
			    do while not rs.EOF
				    strImageID = clng(trim(strImage))
				    if ( rs("Preinstall") or rs("Preload") or rs("ARCD") or rs("SelectiveRestore") ) and rs("InImage") and ( trim(rs("Images") & "") = "" or instr(", " & rs("Images") & ",", ", " & strImageID & ",")>0  or instr( rs("Images") , "(" & strImageID & "=")>0 )  then
					    strDelList = strDelList & ", " & rs("ID")
					    if instr(rs("Images"),"(" & trim(clng(trim(strImage))) & "=") > 0 then
						    strExecptions = strExecptions & ";(" & rs("ID") & "=" & GetExceptions(rs("Images") & "", clng(trim(strImage)) & "") & ")"
					    end if
				    end if
				    rs.MoveNext
			    loop
			    rs.Close
			    if strDelList <> "" then
				    strDelList = mid(strDelList,3) & ","
			    end if
			    if strExecptions <> "" then
				    strExecptions = mid(strExecptions,2)
				    strDelList = strDelList & ":" & strExecptions
			    end if

			    set cm = server.CreateObject("ADODB.Command")
			    cm.CommandType =  &H0004
			    cm.ActiveConnection = cn
		
			    cm.CommandText = "spUpdateImageLockedDeliverableList"	

			    Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			    p.Value = clng(trim(strImage)) 
			    cm.Parameters.Append p

			    Set p = cm.CreateParameter("@DeliverableList", 201, &H0001, 2147483647)
			    p.value = strDelList
			    cm.Parameters.Append p

			    cm.Execute rowschanged

			    set cm = nothing		
			
			
			    if rowschanged <> 1 then
				    FoundErrors = true
				    exit for
			    end if

                'End Lock Deliverable List

            next

        end if

        'Remember the Notification List
        on error resume next    
	    set cm = server.CreateObject("ADODB.Command")
   	    cm.CommandType =  &H0004
        cm.ActiveConnection = cn
  		
        cm.CommandText = "spUpdateRTMNotificationList"	
  
        Set p = cm.CreateParameter("@ID", 3,  &H0001)
        p.Value = clng(request("txtProductID"))
        cm.Parameters.Append p

        Set p = cm.CreateParameter("@RTMNotifications", 200,  &H0001,2147483647)
        p.Value = trim(request("txtNotify"))
        cm.Parameters.Append p
        
   		cm.Execute RowsEffected

        set cm = nothing
        on error goto 0        

               
        if strSuccess = "0" then
            cn.rollbacktrans
        else
            cn.committrans
        end if
        
    end if
	
    if strSuccess <> "0" then
    	Set oMessage = New EmailQueue
	
        if trim(request("txtCurrentUserEmail")) = "" then
	        omessage.from = "max.yu@hp.com"
	    else
	        omessage.from = request("txtCurrentUserEmail")
	    end if
	    if clng(request("txtProductID")) = 100 or trim(request("txtNotify")) = "" then
	        oMessage.To= "max.yu@hp.com"
	    else
	        oMessage.To= request("txtNotify") & ";max.yu@hp.com" 
	    end if
	    oMessage.Subject = request("txtProductName") & " RTM Notification - " & request("txtRTMName")
    	oMessage.Importance = cdoNormal
	    oMessage.HTMLBody = "<STYLE>.EmbeddedTable TBODY TD{FONT-FAMILY: Verdana;}.EmbeddedTable TBODY TD{Font-Size: xx-small;}</style>" & trim(request("txtEmailPreview"))
        if trim(request("txtAttachmentPath1")) <> "" then
            on error resume next
            oMessage.AddAttachment(request("txtAttachmentPath1"))
            on error goto 0
        end if
	    oMessage.Send 
	    Set oMessage = Nothing 			
    end if

	set rs = nothing
    cn.Close
    set cn = nothing
	
%>
Done</font><br>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
</BODY>
</HTML>
