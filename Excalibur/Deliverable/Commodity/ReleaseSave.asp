<%@ Language=VBScript %>
<!-- #include file="../../includes/emailwrapper.asp" -->
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
                window.parent.parent.parent.popupCallBack(1);
            }
            else {
                window.parent.returnValue = 1;
                window.parent.close();
            }
        }
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%
	dim strSQL
	dim cn
    dim rs
	dim FoundErrors
    dim IDArray
    dim i
   	dim j
    dim CurrentUser
    dim CurrentUserID
    dim strItem
    dim strSupported
    dim strLanguageList
    dim Lang
    dim strFilename
    dim strDeveloperReassignments
    dim ReassignmentArray
    dim ValueArray
    dim ValueSet    
    strDeveloperReassignments = ""
    
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.ConnectionTimeout = 180
	cn.Open


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

	if not (rs.EOF and rs.BOF) then
		CurrentUserEmail = rs("Email")
		CurrentUserID = rs("ID")
	end if
	rs.Close


    response.Write CurrentUserEmail


    cn.BeginTrans
    blnFailed =false

    
    IDArray = split(request("chkID"),",")
    for i = 0 to ubound(IDArray)
        response.Write "<BR><BR>Processing Item " & i+1 
        response.Write "<BR>" & IDArray(i)
        response.Write "<BR>" & request("txtEmail" & trim(IDArray(i)))
        response.Write "<BR>" & request("txtLocation" & trim(IDArray(i)))
        response.Write "<BR>" & request("txtComments" & trim(IDArray(i)))
        response.Write "<BR>" & request("txtExecutionEngineer" & trim(IDArray(i)))
        response.Write "<BR>" & request("txtExecutionEngineerName" & trim(IDArray(i)))
        response.Write "<BR>" & request("SelectedMilestoneName" & trim(IDArray(i)))
        response.Write "<BR>" & request("SelectedMilestoneID" & trim(IDArray(i)))
        response.Write "<BR>" & request("NextMilestoneID" & trim(IDArray(i)))
        
        strSupported = ""
		rs.Open "spGetOTSIDsByDelVersion " & clng(IDArray(i)),cn,adOpenForwardOnly
		strItem = ""
		if not (rs.EOF and rs.BOF) then
			do while not rs.EOF
				strItem = strItem & ", " & rs("OTSNumber") 
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3)
			end if
			strSupported = strSupported & "OBSERVATIONS FIXED: " & strItem & "<BR>"
		end if
		
		rs.Close

		rs.Open "spGetProductsForVersion " & clng(IDArray(i)),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strSupported = strSupported & "PRODUCTS: Product Independent<BR>"
		else
			strItem = ""
			do while not rs.EOF
				strItem = strItem & ", " & rs("Family") & " " & rs("Version")
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3) 
			end if
			strSupported = strSupported & "PRODUCTS: " & strItem & "<BR>"
		end if
		
		rs.Close
		
		rs.Open "spGetSelectedOS " & clng(IDArray(i)),cn,adOpenForwardOnly
		strItem = ""
		if rs("ID")=16 then
			strSupported = strSupported
		elseif rs.EOF and rs.BOF then
			strSupported = strSupported & "OPERATING SYSTEMS: OS Independent<BR>"
		else
			do while not rs.EOF
				strItem = strItem & ", " & rs("Name") 
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3)
			end if
			if trim(strItem) <> "" then
				strSupported = strSupported & "OPERATING SYSTEMS: " & strItem & "<BR>"
			else
				strSupported = strSupported & "OPERATING SYSTEMS: OS Independent<BR>"
			end if
		end if
		
		rs.Close
		
		rs.Open "spGetSelectedLanguages " & clng(IDArray(i)),cn,adOpenForwardOnly
		strItem = ""
		strLanguageList = ""
		do while not rs.EOF
			strItem = strItem & ", " & rs("Abbreviation") 
			strLanguageList = strLanguageList & "," & rs("Name") & "," & rs("ID")
			rs.MoveNext
		loop
		if len(strItem)>0 then
			strItem = mid(strItem,3)
		end if
		if strItem = "XX" then
		    strSupported = strSupported & "COUNTRIES: Country Independent<BR>"
		else
		    strSupported = strSupported & "COUNTRIES: " & strItem & "<BR>"
		end if
		rs.Close
		if len(strLanguageList) > 0 then
		    strLanguageList = mid(strLanguageList,2)
		end if

		rs.open "spGetVersionProperties4Web " & clng(IDArray(i)),cn,adOpenForwardOnly
	    strFilename = ""
		if not(rs.EOF and rs.BOF) then
		    strFilename = trim(rs("Filename") & "")
			if trim(request("txtComments" & trim(IDArray(i)))) <> "" then
				strComments = propercase(request("SelectedMilestoneName" & trim(IDArray(i)))) & " Comments:" & vbcrlf & request("txtComments" & trim(IDArray(i))) & vbcrlf & vbcrlf &  rs("Comments") 
			else
				strComments = rs("Comments")
			end if
			
			strDeliverableReleaseString = request("SelectedMilestoneName" & trim(IDArray(i))) & ": " & rs("Name") & " " & rs("Version")
			
			if rs("Revision") & "" <> "" then
				strDeliverableReleaseString = strDeliverableReleaseString & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strDeliverableReleaseString = strDeliverableReleaseString & "," & rs("Pass")
			end if
			
			strBody = ""
			strBody = strBody & "ID: " & rs("VersionID") & "<BR>"
			strBody = strBody & "NAME: " & rs("DeliverableName") & "<BR>"
			if rs("PartNumber") & "" <> "" then
				strBody = strBody & "PART NUMBER: " & rs("PartNumber") & "<BR>"
			end if

		    if rs("VENDOR") & "" <> "" and lCase(rs("VENDOR")) & "" <> "hp" and lCase(rs("VENDOR")) & "" <> "compaq" then
			    strBody = strBody & "VENDOR: " & rs("VersionVendor") & "<BR>"
		    end if
		
			if rs("ModelNumber") & "" <> "" then
				strBody = strBody & "MODEL NUMBER: " & rs("ModelNumber") & "<BR>"
			end if
		
			strBody = strBody & "HW VERSION: " & rs("Version") & "<BR>"

			if rs("Revision") & "" <> "" then
				strBody = strBody & "FW VERSION: " & rs("Revision") & "<BR>"
			end if
			
			if rs("Pass") & "" <> "" then
				strBody = strBody & "PASS: " & rs("Pass") & "<BR>"
			end if

			if rs("vendorVersion") & "" <> "" then
				strBody = strBody & "VENDOR VERSION: " & rs("Vendorversion") & "<BR><BR>"
			end if
			
			if rs("Changes") & "" <> "" then
				strBody = strBody & "MODIFICATIONS/ENHANCEMENTS: " & rs("changes") & "<BR><BR>"
			end if

			if trim(request("txtLocation" & trim(IDArray(i)))) <> "" then
				strBody = strBody & "DELIVERABLE LOCATION: " & request("txtLocation" & trim(IDArray(i))) & "<BR>"
			elseif rs("ImagePath") <> "" then
				strBody = strBody & "DELIVERABLE LOCATION: " & rs("ImagePath") & "<BR>"
			else
				strBody = strBody & "DELIVERABLE LOCATION: " & "N/A" & "<BR>"
			end if

			if rs("CDPartNumber") & "" <> "" then
				strBody = strBody & "CD PART NUMBER: " & rs("CDPartNumber") & "<BR>"
			end if
			

			if rs("SampleDate") & "" <> "" then
				strBody = strBody & "SAMPLES AVAILABLE: " & rs("SampleDate") & "<BR><BR>"
			end if
			
			if rs("DevManager") & "" <> "" then
				strBody = strBody & "DEVELOPMENT MANAGER: " & rs("devmanager") & "<BR>"
			end if

            if trim(request("txtExecutionEngineer" & trim(IDArray(i)))) <> "" and trim(request("txtExecutionEngineer" & trim(IDArray(i)))) <> "0" then
				strBody = strBody & "DEVELOPER: " & request("txtExecutionEngineerName" & trim(IDArray(i))) & "<BR><BR>"
			elseif rs("developer") & "" <> "" then
				strBody = strBody & "DEVELOPER: " & rs("developer") & "<BR><BR>"
			end if

		
		
			strBody = strBody & strSupported & "<BR><BR>"
			
			if trim(strComments) <> "" then
				strBody = strBody & "COMMENTS: <BR>" & strComments & "<BR><BR>"
			else 
				strBody = strBody & "COMMENTS: N/A <BR>" & "<BR><BR>"
			end if
			

			if rs("Description") <> "" then
				strBody = strBody & "DESCRIPTION: <BR>" & rs("Description") & "<BR><BR>"
			end if
			

			if trim(rs("Notes") & "") <> "" then
				strBody = strBody & "NOTES: <BR>" & rs("Notes") & "<BR><BR>"
			end if
			

		end if
		rs.close
	
	    'Update Language_Delver 
    	Response.write "<BR>LanguageList: " & strLanguageList & "<BR>"
    	LangArray = split(strLanguageList,",")
    	for j = 0 to ubound(LangArray) step 2
    		Lang = langArray(j)
    		if trim(Lang) <> "" then
    			Select case trim(request("txtFunction"))
	    			case "2" 'Fail/Cancel
		    			cn.Execute "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) & "," & 0 & ",1", RowsEffected
			    	    response.Write ">>>" & "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) & "," & 0 & ",1"
			    	case else 'Release
			    	    if clng(request("NextMilestoneID" & trim(IDArray(i)))) = 0 then
			    	    	cn.Execute "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) & ",1,0",RowsEffected
			    	        response.Write ">>>" & "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) & ",1,0"
			            else
			    	    	cn.Execute "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("NextMilestoneID" & trim(IDArray(i)))) & ",0,0",RowsEffected
			    	        response.Write ">>>" & "spupdateLanguageLocation " & clng(IDArray(i)) & "," & clng(langArray(j+1)) & "," & clng(request("NextMilestoneID" & trim(IDArray(i)))) & ",0,0"
			            end if
			    end select
			    if RowsEffected <> 1 then
				    blnfailed = true
				    exit for
			    end if
		    end if
	    next 

        'Reset if a failure occurred previously
     	if (not blnFailed) and trim(request("txtFunction")) <> "2" then
		    cn.Execute "spresumeafterfailure " & clng(IDArray(i)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) ,RowsEffected		
		    if cn.Errors.count > 0 then
			    blnfailed = true
    		end if
	    end if

    	if (not blnFailed) and trim(request("txtFunction")) <> "2" then
	    	cn.Execute "spupdatemilestonerelease " & clng(IDArray(i)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i))))  ,RowsEffected		
		    if cn.Errors.count > 0 then
			    blnfailed = true
		    end if
        end if
 
 
 
     	if (not blnFailed) then
	    	set cm = server.CreateObject("ADODB.Command")
					            
		    cm.ActiveConnection = cn
		    cm.CommandText = "spUpdateDeliverable4Release"
		    cm.CommandType = &H0004
		                
		    Set p = cm.CreateParameter("@DelID", 3, &H0001)
		    p.Value = clng(IDArray(i))
		    cm.Parameters.Append p

    		Set p = cm.CreateParameter("@NextMilestoneID", 3, &H0001)
	    	p.Value = request("NextMilestoneID" & trim(IDArray(i)))
		    cm.Parameters.Append p

    		Set p = cm.CreateParameter("@Filename", 200, &H0001,60)
	    	p.Value = left(strFilename,60)
    		cm.Parameters.Append p

    		Set p = cm.CreateParameter("@Transfer", 200, &H0001,255)
			p.Value = left(request("txtLocation" & trim(IDArray(i))),255)
	    	cm.Parameters.Append p

    		Set p = cm.CreateParameter("@MD5", 200, &H0001,50)
	    	p.Value = ""
    		cm.Parameters.Append p

		    Set p = cm.CreateParameter("@BaseFilePath", 200, &H0001,1024)
		    p.Value = ""
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@CVASubPath", 200, &H0001,255)
		    p.Value = ""
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@ReleasePriority", 16, &H0001)
		    p.Value = Null
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@ReleasePriorityJustification", 200, &H0001,80)
		    p.Value = Null
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@Comments", 201, &H0001, 2147483647)
			p.value =  strComments
    		cm.Parameters.Append p
	            
	    	cm.Execute RowsEffected
	
    		If cn.Errors.count > 0 Then
	    		blnFailed = true
		    End If
		    set cm = nothing
		    set p = nothing
	    end if

    	if (not blnFailed) and trim(request("txtFunction")) = "2" then
	    	cn.Execute "spupdatemilestonefailure " & clng(IDArray(i)) & "," & clng(request("SelectedMilestoneID" & trim(IDArray(i)))) ,RowsEffected		
		    if cn.Errors.count > 0 then
			    blnfailed = true
		    end if
	    end if

        if (not blnFailed) and trim(request("txtFunction")) <> "2" and clng(request("NextMilestoneID" & trim(IDArray(i)))) = 0  then
	    	cn.Execute "spLogReleaseToProducts " & clng(IDArray(i)) & "," & clng(currentuserid) ,RowsEffected		
		    if cn.Errors.count > 0 then
			    blnFailed = true
		    end if
        end if

    	if (not blnFailed) then
	    	cn.Execute "spUpdateDeliverableLocation " &  clng(IDArray(i)) ,RowsEffected		
		    if cn.Errors.count > 0 then
			    blnFailed = true
    		end if
	    end if

        
        if not blnFailed then
            if trim(request("txtFunction")) = "2" then
    			strSubject = "Deliverable Failed or Cancelled in " & strDeliverableReleaseString
            else
    			strSubject = "Deliverable Release from " & strDeliverableReleaseString
            end if
            
			if strTo = "" then 'Error message could be sent to admin.  If so, don't override with regular notification list
   				strFullNotify = ""
   				NotifyArray = split(request("txtEmail" & trim(IDArray(i))),";")
   				for each NotifyAddress in NotifyArray
		   			if trim(NotifyAddress) <> "" then
   						strFullNotify = strFullNotify & ";" & NotifyAddress
   					end if
   				next
				if strFullNotify = "" then
					strFullNotify = "max.yu@hp.com" 
				else
					strFullNotify = mid(strFullNotify,2)
				end if
   				strTo = strFullNotify 
			end if	
            
			if trim(CurrentUserEmail) = "" then
				CurrentUserEmail = "max.yu@hp.com"
			end if
			
   			dim strMailHeader	
  		
			strMailHeader = "<font face=Arial size=2>"
            if trim(request("txtFunction")) = "2" and trim(request("txtComments" & trim(IDArray(i)))) <> "" then
				strMailHeader = strMailHeader & trim(server.HTMLEncode(trim(request("txtComments" & trim(IDArray(i)))))) & "<BR><BR>"
			end if	
			strMailHeader = strMailHeader & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur.asp"">Open Pulsar</a></font><br><br>"
			strMailHeader = strMailHeader & "<font face=Arial size=2>"
				
			strBody = replace(replace(strMailHeader & strBody,vbcrlf,"<BR>"),"""","&QUOT;")
			Set oMessage = New EmailWrapper 
			
			oMessage.From = CurrentUserEmail

			if trim(strTo) = "" then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= replace(strTo,";;",";")
			end if
            'oMessage.BCC = "max.yu@hp.com"
			oMessage.Subject = strSubject
				
			oMessage.HTMLBody = strBody
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 
            
            
        end if

 
        if blnFailed then
            exit for
        end if 
'-------------------------------------------- 
        strDeveloperReassignments = strDeveloperReassignments & "," & IDArray(i) & "|" & request("txtExecutionEngineer" & trim(IDArray(i)))

    next

    if blnFailed then
        cn.RollbackTrans
    else
        cn.CommitTrans
    end if

    'Reassign Developers if necessary
    if trim(strDeveloperReassignments) <> "" then
        strDeveloperReassignments = mid(strDeveloperReassignments,2)
    end if    
    response.Write strDeveloperReassignments
    if trim(strDeveloperReassignments) <> "" and (not blnFailed) then
        ReassignmentArray = split(strDeveloperReassignments,",")
        for each ValueSet in ReassignmentArray
            if valueset <> "" then
                ValueArray = split (ValueSet,"|")
        	    if trim(ValueArray(1)) <> "" and trim(ValueArray(1)) <> "0" and (not blnFailed) then
    	            cn.execute "spUpdateDeliverableDeveloper " & clng(ValueArray(0)) & "," & clng(ValueArray(1))
                    response.Write "<BR>" & "spUpdateDeliverableDeveloper " & clng(ValueArray(0)) & "," & clng(ValueArray(1))
                end if
            end if
        next
    end if
    
    set rs = nothing
    cn.Close
    set cn = nothing
    
    if blnFailed then
        response.Write "<input id=""txtSuccess"" type=""text"" value=""0"">"
    else
        response.Write "<input id=""txtSuccess"" type=""text"" value=""1"">"
    end if  
function ProperCase(sString)
	if sString = "PM" then
		Propercase = sstring
	else
	
		Dim lTemp
		Dim sTemp, sTemp2
		Dim x
		sString = LCase(sString)
		if Len(sString) Then
		    sTemp = Split(sString, " ")
			lTemp = UBound(sTemp)
			For x = 0 To lTemp
				sTemp2 = sTemp2 & UCase(Left(sTemp(x), 1)) & Mid(sTemp(x), 2) & " " 
			Next
			ProperCase = trim(sTemp2)
		Else
			ProperCase = sString
	    End if
	end if
 end function
    
%>
</body>
</html>

