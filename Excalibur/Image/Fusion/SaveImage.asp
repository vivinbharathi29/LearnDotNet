<%@ Language=VBScript %>
<% Option Explicit %>

<!-- #include file="../../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (txtSuccess.value == "1") {
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.closeExternalPopup();
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);                
          }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
                /*window.returnValue = 1;
                window.parent.close();*/
            }
        }
        else
            document.write("Unable to save Image Definition.");
		
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
    

<%


	dim RegionArray
	dim TagArray
	Dim ImageIDArray
	Dim ImageNameArray
	dim i
	dim cn
	dim cm
	dim p
	dim blnRegionsChanged
	dim blnDefinitionChanged
	dim rowschanged
	dim FoundErrors
	dim NewID
	dim strLog
	dim strChangeLog
	dim strRegionChangeLog
	dim strChangeMail
	dim strRegionChangeMail
	dim strProductDrop
	dim oMessage
	dim strBody
	dim strProductName
	dim strPreinstallTeam
	dim rs
	dim rs2
	dim strDelList
	dim strExecptions
	dim strImageID
	dim strProductEmail
	dim strImageDefinitionId
    dim strBrandsAdded
    dim strBrandsRemoved
    dim BrandIDTagArray
    dim BrandIDArray
    dim UpdateArray
    dim PublishArray
    dim PublishTagArray
    dim PublishUpdated

    response.write "BrandTag: " & request("txtBrandIDTag") & "<br>"
    response.write "Brands: " & request("chkBrands") & "<br>"

    response.write "TagPublish: " & request("tagPublish") & "<br>"
    response.write "Publish: " & request("chkPublish") & "<br>"
    
    PublishArray = split(request("chkPublish"),",")
    PublishTagArray = split(request("tagPublish"),",")

    if replace(request("chkPublish")," ","") <> replace(request("tagPublish")," ","") then
        PublishUpdated = true
    else
        PublishUpdated = false
    end if

    BrandIDTagArray = split(request("txtBrandIDTag"),",")
    BrandIDArray = split(request("chkBrands"),",")

    strBrandsRemoved = ""
    for i = 0 to ubound(BrandIDTagArray)
        if not inarray(BrandIDArray,BrandIDTagArray(i)) then
            strBrandsRemoved = strBrandsRemoved & "," & BrandIDTagArray(i)
        end if
    next

    if trim(strBrandsRemoved) <> "" then
        strBrandsRemoved = mid(strBrandsRemoved,2)
    end if

    strBrandsAdded = ""
    for i = 0 to ubound(BrandIDArray)
        if not inarray(BrandIDTagArray,BrandIDArray(i)) then
            strBrandsAdded = strBrandsAdded & "," & BrandIDArray(i)
        end if
    next

    if trim(strBrandsAdded) <> "" then
        strBrandsAdded = mid(strBrandsAdded,2)
    end if

    response.write "Brands Removed: " & strBrandsRemoved & "<br>"
    response.write "Brands Added: " & strBrandsAdded & "<br>"

'		if trim(request("txtBrandTag")) <>  trim(request("txtBrandText")) then
'			strLog = strLog & "<TR><TD>Brand</TD>"
'			strLog = strLog &  "<TD>" & request("txtBrandTag") & "&nbsp;</TD>"
'			strLog = strLog & "<TD>" & request("txtBrandText") & "&nbsp;</TD></TR>"	
'			strChangeLog = strChangeLog & "Brand Changed from " & request("txtBrandTag") & " to " & request("txtBrandText") & "; "			
'		end if



	
	strProductDrop = request("txtProductDrop")
	
	blnRegionsChanged = request("txtTag") <> request("chkRegion")

	blnDefinitionChanged =  request("tagImageType") <> request("cboImageType") or request("tagComments") <> request("txtComments")  or request("tagRTMDate") <> request("txtRTMDate") or request("tagOS") <>  request("cboOS") or request("tagProductDrop") <> strProductDrop or request("cboStatus") <> request("tagStatus") or request("cboDriveDefinition") <> request("tagDriveDefinition") or request("tagOSRelease") <>  request("cboOSRelease")

	NewID = trim(request("txtDisplayedID"))

    'PBI 19513 / Task 20989; Bug 22201/ Task 22901 
    Dim bCopyWithTarget
    Dim iCopyID
    If Request("inpCopyWithTarget") <> "" Then
        If Request("inpCopyWithTarget") = "True" Then
            bCopyWithTarget = True
            iCopyID = Request("txtCopyID")
	    Else
		    bCopyWithTarget = False
            iCopyID = 0
	    End If
    Else
       bCopyWithTarget = False 
       iCopyID = 0
    End If

	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	cn.BeginTrans

	FoundErrors = false	

	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")	
	
	cm.CommandText = "spGetProductVersionname"	

	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.Value = request("txtProdID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	if not(rs.EOF and rs.BOF) then
		strProductName = rs("Name") & ""
		strPreinstallTeam = rs("preinstallteam") & ""
		strProductEmail = rs("Email") & ""
	end if
	rs.Close
	set rs=nothing

	strLog = ""
	strChangeLog = ""
	strChangeMail = ""	

	if  blnDefinitionChanged then
	
		if request("tagProductDrop") <> strProductDrop then
			strLog = strLog & "<TR><TD>Product Drop</TD>"
			strLog = strLog &  "<TD>" & request("tagProductDrop") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & strProductDrop & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Product Drop Changed from " & request("tagProductDrop") & " to " & strProductDrop & "; "			
		end if
		
		if request("tagOS") <>  request("cboOS") then
			strLog = strLog & "<TR><TD>Operating System</TD>"
			strLog = strLog &  "<TD>" & request("txtOSTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtOSText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "OS Changed from " & request("txtOSTag") & " to " & request("txtOSText") & "; "
		end if
        response.write "TagStatus: " & request("tagStatus") & "<BR>"
        response.write "cboStatus: " & request("cboStatus") & "<BR>"
		if request("tagStatus") <>  request("cboStatus") then
			strLog = strLog & "<TR><TD>Status</TD>"
			strLog = strLog &  "<TD>" & request("txtStatusTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtStatusText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Status Changed from " & request("txtStatusTag") & " to " & request("txtStatusText") & "; "
		end if

		if request("tagRTMDate") <>  request("txtRTMDate") then
			strLog = strLog & "<TR><TD>RTM Date</TD>"
			strLog = strLog &  "<TD>" & request("tagRTMDate") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtRTMDate") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "RTM Date Changed from " & request("txtRTMDate") & " to " & request("txtRTMDate") & "; "
		end if
		
		if request("tagDriveDefinition") <> request("cboDriveDefinition") then
			strLog = strLog & "<TR><TD>Master SKU Component</TD>"
			strLog = strLog &  "<TD>" & request("tagDriveDefinition") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("cboDriveDefinition") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Master Sku Component Changed from " & request("tagDriveDefinition") & " to " & request("cboDriveDefinition") & "; "
		end if

		if request("tagComments") <>  request("txtComments") then
			strLog = strLog & "<TR><TD>Comments</TD>"
			strLog = strLog &  "<TD>" & request("tagComments") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtComments") & "&nbsp;</TD></TR>"			
		end if

        if request("tagOSRelease") <>  request("cboOSRelease") then
			strLog = strLog & "<TR><TD>Operating System Version</TD>"
			strLog = strLog &  "<TD>" & request("txtOSReleaseTag") & "&nbsp;</TD>"
			strLog = strLog & "<TD>" & request("txtOSReleaseText") & "&nbsp;</TD></TR>"			
			strChangeLog = strChangeLog & "Operating System Version Changed from " & request("txtOSReleaseTag") & " to " & request("txtOSReleaseText") & "; "
		end if

		strChangeMail = strProductName & ", " & strProductDrop & ", Image Definition has changed: "
		strChangeMail = strChangeMail & strChangelog
		

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		if trim(request("txtDisplayedID")) = "" then
			Response.Write "Adding<BR>"
			cm.CommandText = "spAddImageDefinition"	

			Set p = cm.CreateParameter("@ProdID", 3,  &H0001)
			p.Value = -1 'request("txtProdID")
			cm.Parameters.Append p
		else
			Response.Write "Updating<BR>"
			cm.CommandText = "spUpdateImageDefinition"	
			
			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = request("txtDisplayedID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@ImageChange", 200, &H0001, 500)
			p.Value = left(strChangeLog, 500)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@RegionChange", 200, &H0001, 500)
			p.Value = ""
			cm.Parameters.Append p

		end if
		
		Set p = cm.CreateParameter("@SKU", 200, &H0001, 20)
		p.value = left(strProductDrop,20)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@BrandID", 3, &H0001)
		p.Value = 74 'request("cboBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSID", 3, &H0001)
		p.Value = request("cboOS")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SWID", 3, &H0001)
		p.Value = 32 'request("cboSW")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ImageType", 200, &H0001, 10)
		p.value = ""'left(request("cboType"),10)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ImageTypeID", 3, &H0001)
		p.value = clng(request("cboImageType"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@StatusID", 16, &H0001)
		p.Value = request("cboStatus")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RTMDate", 135, &H0001)
		if trim(request("txtRTMDate")) = "" then
			p.Value = null
		else
			p.Value = request("txtRTMDate")
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Comments", 200, &H0001, 50)
		p.value = left(request("txtComments"),50)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@ImageDriveDefinitionId", 3,  &H0001)
		If Request("cboDriveDefinition") = "" Then
            p.Value = NULL
    	Else
		    p.Value = request("cboDriveDefinition")
		End If
		cm.Parameters.Append p

        Set p = cm.CreateParameter("@OSReleaseId", 3, &H0001)
        if (isnull(request("cboOSRelease")) or request("cboOSRelease")="" ) then
            p.Value = Null
        else
		    p.Value = request("cboOSRelease")
        end if
		cm.Parameters.Append p

		if trim(request("txtDisplayedID")) = "" then
			Set p = cm.CreateParameter("@NewID", 3,  &H0002)
			cm.Parameters.Append p
		end if
		
		cm.Execute rowschanged
		if rowschanged <> 1 then
			FoundErrors = true
		else
			if trim(request("txtDisplayedID")) = "" then
				NewID = cm("@NewID")
			end if
		end if
		
		set cm = nothing
		
	end if


    strRegionChangeLog = ""


    '*******************Link Image Definition to Product if adding or if ProductDrop updated ***********************************
    if (trim(request("txtDisplayedID")) = "" or request("tagProductDrop") <> strProductDrop) and trim(request("txtProdID"))  <> "" and (not FoundErrors) then
    	    set cm = server.CreateObject("ADODB.Command")
		    cm.ActiveConnection = cn
		    cm.CommandType =  &H0004
    	    cm.CommandText = "spUpdateImageDefinitionProductLink"
    
		    Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
		    p.Value = NewID
		    cm.Parameters.Append p
    
    	    Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		    p.Value = clng(request("txtProdID") )
		    cm.Parameters.Append p

    	    Set p = cm.CreateParameter("@ProductDrop", 200,  &H0001,30)
		    p.Value = left(request("txtProductDrop"),30 )
		    cm.Parameters.Append p
    
    	    cm.Execute rowschanged
            
            set cm=nothing
    end if


    if trim(strBrandsRemoved) <> "" and (not FoundErrors) then
    
        UpdateArray = split(strBrandsRemoved,",")

        for i = 0 to ubound(UpdateArray)
    	    set cm = server.CreateObject("ADODB.Command")
		    cm.ActiveConnection = cn
		    cm.CommandType =  &H0004
    	    cm.CommandText = "spUpdateImageDefinitionRemoveBrand"
    
		    Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
		    p.Value = NewID
		    cm.Parameters.Append p
    
    	    Set p = cm.CreateParameter("@BrandID", 3,  &H0001)
		    p.Value = UpdateArray(i)
		    cm.Parameters.Append p
    
    	    cm.Execute rowschanged
            set cm=nothing
        next
    end if


    if trim(strBrandsAdded) <> "" and (not FoundErrors)  then
    
        UpdateArray = split(strBrandsAdded,",")

        for i = 0 to ubound(UpdateArray)
    	    set cm = server.CreateObject("ADODB.Command")
		    cm.ActiveConnection = cn
		    cm.CommandType =  &H0004
    	    cm.CommandText = "spUpdateImageDefinitionAddBrand"
    
		    Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
		    p.Value = NewID
		    cm.Parameters.Append p
    
    	    Set p = cm.CreateParameter("@BrandID", 3,  &H0001)
		    p.Value = UpdateArray(i)
		    cm.Parameters.Append p
    
    	    cm.Execute rowschanged
            set cm=nothing
        next
    end if

    if PublishUpdated and trim(request("txtDisplayedID")) <> "" and (not FoundErrors) then

		ImageIDArray = split(request("txtimageIDList"),",")

        for i = 0 to ubound(ImageIDArray)
            if inarray(PublishTagArray,ImageIDArray(i)) and not inarray(PublishArray,ImageIDArray(i)) then
                'Remove
                
        	    set cm = server.CreateObject("ADODB.Command")
	    	    cm.ActiveConnection = cn
		        cm.CommandType =  &H0004
    	        cm.CommandText = "spUpdateImagePublished"
    
		        Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
		        p.Value = clng(request("txtDisplayedID"))
		        cm.Parameters.Append p

		        Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
		        p.Value = ImageIDArray(i)
		        cm.Parameters.Append p
    
    	        Set p = cm.CreateParameter("@Published", 11,  &H0001)
		        p.Value = 0
		        cm.Parameters.Append p
    
    	        cm.Execute rowschanged
                set cm=nothing

            elseif not inarray(PublishTagArray,ImageIDArray(i)) and inarray(PublishArray,ImageIDArray(i)) then
                'Add
        	    set cm = server.CreateObject("ADODB.Command")
	    	    cm.ActiveConnection = cn
		        cm.CommandType =  &H0004
    	        cm.CommandText = "spUpdateImagePublished"

		        Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
		        p.Value = clng(request("txtDisplayedID"))
		        cm.Parameters.Append p
    
		        Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
		        p.Value = ImageIDArray(i)
		        cm.Parameters.Append p
    
    	        Set p = cm.CreateParameter("@Published", 11,  &H0001)
		        p.Value = 1
		        cm.Parameters.Append p
    
    	        cm.Execute rowschanged
                set cm=nothing

            end if
        next

    end if


	if blnRegionsChanged and (not FoundErrors) then
		RegionArray = split(request("chkRegion"),",")
		TagArray = split(request("txtTag"),",")
		ImageIDArray = split(request("txtimageIDList"),",")
		ImageNameArray = split(request("txtimageNameList"),",")
	
			for i = lbound(ImageIDArray) to ubound(ImageIDArray)
				if inarray(RegionArray,ImageIDArray(i)) and not inarray(TagArray,ImageIDArray(i)) then
                    response.write "Add: " & ImageIDArray(i) & "<br>"
					strRegionChangeLog = strRegionChangeLog & "Added Region " & ImageNameArray(i) & "; "

    					set cm = server.CreateObject("ADODB.Command")
	    				cm.ActiveConnection = cn
		    			cm.CommandType =  &H0004
						cm.CommandText = "spAddRegion2ImageFusion"

						Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
						p.Value = NewID
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
						p.Value = ImageIDArray(i)
						cm.Parameters.Append p
						
						Set p = cm.CreateParameter("@Priority", 200,  &H0001,6)
						p.Value = "1"
						cm.Parameters.Append p

   						Set p = cm.CreateParameter("@Published", 11,  &H0001)
                        if inarray(PublishArray,ImageIDArray(i)) then
						    p.Value = 1
                        else
						    p.Value = 0
                        end if
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@ProductDrop", 200,  &H0001,30)
						p.Value = left(request("txtProductDrop"),30 )
						cm.Parameters.Append p

    					cm.Execute rowschanged
    	
    					if rowschanged <> 1 then
    						FoundErrors = true
    						exit for
    					end if
                        
                        set cm=nothing

                elseif inarray(TagArray,ImageIDArray(i)) and not inarray(RegionArray,ImageIDArray(i)) then
                    response.write "Remove: " & ImageIDArray(i) & "<br>"
					strRegionChangeLog = strRegionChangeLog & "Removed Region " & ImageNameArray(i) & "; "

    					set cm = server.CreateObject("ADODB.Command")
	    				cm.ActiveConnection = cn
		    			cm.CommandType =  &H0004
						cm.CommandText = "spRemoveRegionFromImage"

						Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
						p.Value = NewID
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@RegionID", 3,  &H0001)
						p.Value = ImageIDArray(i)
						cm.Parameters.Append p
                        
                        cm.Execute rowschanged
    	
    					if rowschanged <> 1 then
    						FoundErrors = true
    						exit for
    					end if
                        
                        set cm=nothing
                end if
    
			next			
		
	end if
	

	'Snapshot Deliverable List if releasing to factory	
	if trim(request("tagStatus")) = "1" and trim(request("cboStatus")) = "2"  and (not FoundErrors) then
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")


		rs.Open "Select * from images with (NOLOCK) where ImagedefinitionID = " &  request("txtDisplayedID"),cn,adOpenStatic
		do while not rs.EOF
            if trim(rs("LockedDeliverableList") & "") = "" then
			    rs2.Open "spListDeliverablesInImage " & rs("ID") & ",0" ,cn,adOpenStatic 
			    strDelList = ""
			    strExecptions = ""
			    do while not rs2.EOF
				    strImageID = rs("ID")
				    if ( rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore") ) and rs2("InImage") and ( trim(rs2("Images") & "") = "" or instr(", " & rs2("Images") & ",", ", " & strImageID & ",")>0  or instr( rs2("Images") , "(" & strImageID & "=")>0 )  then
					    strDelList = strDelList & ", " & rs2("ID")
					    if instr(rs2("Images"),"(" & trim(rs("ID")) & "=") > 0 then
						    strExecptions = strExecptions & ";(" & rs2("ID") & "=" & GetExceptions(rs2("Images") & "", rs("ID") & "") & ")"
					    end if
				    end if
				    rs2.MoveNext
			    loop
			    rs2.Close
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
			    p.Value = rs("ID")
			    cm.Parameters.Append p

			    Set p = cm.CreateParameter("@DeliverableList", 201, &H0001, 2147483647)
			    p.value = strDelList
			    cm.Parameters.Append p

			    'cm.Execute rowschanged

			    set cm = nothing		
			
			    '****Uncomment
			    'if rowschanged <> 1 then
				 '   FoundErrors = true
				 '   exit do
			    'end if
                '***uncomment
            end if			
			rs.MoveNext
		loop
		rs.Close

		set rs2 = nothing
		set rs = nothing	
	end if
	
	
	
	strRegionChangeMail = strProductName & ", " & strProductDrop & ", Region has changed: "
	strRegionChangeMail = strRegionChangeMail & strRegionChangelog
	

	Response.Write "Updating Log<BR>"

	if (blnRegionsChanged or blnDefinitionChanged) and (not FoundErrors)  then
		
		if trim(request("txtDisplayedID")) <> "" then
			'*****UNcomment
            'cn.Execute "spImageModified " & NewID,rowschanged
			'if rowschanged <> 1 then
			'	FoundErrors = true
			'end if
            '*****Uncomment
		end if
				
		if (not founderrors) and trim(request("cboDCR")) <> "" then
			set cm = server.CreateObject("ADODB.Command")
			cm.ActiveConnection = cn
			cm.CommandType =  &H0004
			cm.CommandText = "spAddIMageLog"

			Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
			p.Value = clng(request("txtCurrentUserID"))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@DCRID", 3,  &H0001)
			if request("cboDCR") <> "" and isnumeric(request("cboDCR")) then
				p.Value = clng(request("cboDCR"))
			else
				p.Value = null
			end if
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = NewID
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@Details", 200,  &H0001,7500)
			if trim(request("txtDisplayedID")) = "" then
				p.Value = "Added"
			else
				p.Value = left(strLog,7500)
			end if
			cm.Parameters.Append p
		
    '****Uncomment*****			

			'cm.Execute rowschanged
	
			'if rowschanged <> 1 then
			'	FoundErrors = true
			'end if
'****Uncomment*****			
		
		end if
	end if


	if FoundErrors then
		Response.Write "Rolling Back<BR>"

		cn.RollbackTrans
		%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="0"><%
	else
		cn.CommitTrans
		Response.Write "Save Complete<BR>"
		%><INPUT type="hidden" id=txtSuccess name=txtSuccess value="1"><%
		
		'update Region Changes
		if 	strRegionChangeLog <> "" and request("txtDisplayedID") <> "" then
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
	
			cm.CommandText = "spUpdateImageDefinitionRegionChange"	

			Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
			p.Value = request("txtDisplayedID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@RegionChange", 200, &H0001, 500)
			p.Value = left(strRegionChangeLog, 500)
			cm.Parameters.Append p

'			cm.Execute 
			Set cm=nothing
		end if

		If request("txtProdID") <> 100 AND strProductEmail = "1" Then
		    if strChangeLog <> "" or strRegionChangeLog <> "" then
			    'send Email to Preinstall DB team when there is a definition or region change
			    Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			    'Set oMessage.Configuration = Application("CDO_Config")		
            
			    if request("txtCurrentUserEmail") = "" then
				    oMessage.From = "max.yu@hp.com"
			    else
				    oMessage.From = request("txtCurrentUserEmail")
			    end if
            
			    if strPreinstallTeam = 1 then
				    oMessage.To = "houportpreindb@hp.com" 
			    elseif strPreinstallTeam = 2 then
				    oMessage.To = "TWN.PDC.NB-preinPM@hp.com"
			    else
				    oMessage.To = "max.yu@hp.com"
			    end if 
    		
			    oMessage.Subject = "Image Change Notification"
            
			    strBody = "<b><font size=3 face=verdana>" & "Image and Region Changes" & "</font></b>"
			    strBody = strBody & "<br><br>"
			    strBody = strBody & strChangeMail & "<br><br>"
			    strBody = strBody & strRegionChangeMail & "<br>"
    		
			    oMessage.HTMLBody = "<font size=2 face=verdana><Font color=red><b>NOTE: This email was generated by a change to the <u>definiton</u> of the images defined for the product.</b></font><BR><BR>" & strBody & "</font>"
    						
'			    oMessage.Send 
			    Set oMessage = Nothing 		
		    end if
		End If

	end if
	
    '-----------------------------------------------------------------
	'Procedure: usp_Image_CopyWithTargeting
    'Purpose:   Copy Image Definition with Targeting
	'Modified:  Harris, Valerie (5/31/2016) - PBI 19513 / Task 20989 - if Copy With Targeting, update the Product's Deliverable and Root table's Images with Image Summary
	'           Harris, Valerie (7/4/2016) - Bug 22201/ Task 22901 - Pulsar Test: Copy with Targeting Failure
	'-----------------------------------------------------------------
    If bCopyWithTarget = True Then
        set cm = server.CreateObject("ADODB.Command")
        cm.ActiveConnection = cn
        cm.CommandType =  &H0004
        cm.CommandText = "usp_Image_CopyWithTargeting"
    
        Set p = cm.CreateParameter("@p_ProductID", 3,  &H0001)
        p.Value = clng(request("txtProdID"))
        cm.Parameters.Append p

        Set p = cm.CreateParameter("@p_ImageDefinitionID", 3,  &H0001)
        p.Value = clng(NewID)
        cm.Parameters.Append p

        Set p = cm.CreateParameter("@p_CopiedImageDefinitionID", 3,  &H0001)
        p.Value = clng(iCopyID)
        cm.Parameters.Append p
        
        cm.Execute rowschanged
            
        set cm=nothing
    End If

	cn.Close
	set cm = nothing
	set cn = nothing
		
Function GetExceptions(strList, strItem)
	dim strTemp
	strTemp = mid(strList, instr(strList,"(" & strItem & "=") + len(strItem) + 2) 
	strTemp = left(strTemp,instr(strTemp,")")-1)
	GetExceptions = strTemp
end function	

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
</BODY>
</HTML>
