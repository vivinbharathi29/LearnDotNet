<%@ Language=VBScript %>
<%option explicit%>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value == "1")
		{
			window.returnValue = 1;
			window.parent.close();
		}
	//else
		//document.write ("Unable to save HFCN.");
		
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<!--
><%=Request("cboDeliverable")%><BR>
><%=Request("txtDeliverableName")%><BR>
><%=Request("cboCategory")%><BR>
><%=Request("txtVersion")%><BR>
><%=Request("txtRevision")%><BR>
><%=Request("cboDeveloper")%><BR>
><%=Request("txtEmail")%><BR>
><%=Request("txtLocation")%><BR>
><%=Request("OTS2ADD")%><BR>
><%=Request("OTSSummary")%><BR>
><%=Request("txtDependencies")%><BR>
><%=Request("txtChanges")%><BR>
><%=Request("AddProducts")%><BR>
><%=Request("AddProductsNames")%><BR>
-->



<%
	dim strBody
	dim strSubject
	dim strTO
	dim strFooter
	
	strSubject = "NEW HFCN: " & Request("txtTitle") 'Request("txtDeliverableName") & " " &  Request("txtVersion") & " " & ucase(Request("txtRevision")) & " : "
	strTO = Request("txtEmail") 
	if trim(request("cboCategory"))= "12" or trim(request("cboCategory"))= "52" or trim(request("cboCategory"))= "161" or trim(request("cboCategory"))= "166" then
		strTO = strTo & ";David.Crumpler@hp.com;Galen.Hedlund@hp.com"
	end if
	strBody = "TITLE: " & replace(replace(Request("txtTitle"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR>"
	strBody = strBody  & "DELIVERABLE: " & replace(replace(Request("txtDeliverableName"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR>"
	strBody = strBody  & "VERSION: " & replace(replace(ucase(Request("txtVersion")),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR>"
	if Request("txtVendor") <> "" then
		strBody = strBody  & "VENDOR: " & Request("txtVendor") & "<BR>"
	end if
	if Request("txtVendorVersion") <> "" then
		strBody = strBody  & "VENDOR VERSION: " & replace(replace(Request("txtVendorVersion"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR>"
	end if
	if left(request("txtLocation"),2) = "\\" then
		strBody = strBody  & "LOCATION: <a href=""file://" & Request("txtLocation") & """>" & Request("txtLocation") & "</a><BR><BR>"
	else
		strBody = strBody  & "LOCATION: " & replace(replace(Request("txtLocation"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR><BR>"
	end if

	if Request("txtChanges") <> "" then
		strBody = strBody  & "CHANGES: " & "<BR>" & replace(replace(Request("txtChanges"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR><BR>"
	end if


	if request("OTSSummary") <> "" then
		strBody = strBody  & "OBSERVATIONS FIXED: " & "<BR>" & replace(replace(Request("OTSSummary"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR><BR>"
	end if

	if request("txtSpecial") <> "" then
		strBody = strBody  & "SPECIAL NOTES: " & "<BR>" & replace(replace(request("txtSpecial"),vbcrlf,"<BR>"),"""","&QUOT;") & "<BR><BR>"
	end if

	if request("AddProductsNames") = "" then
		strBody = strBody  & "PRODUCTS SUPPORTED: Product Independent" & "<BR><BR>"		
	else
		strBody = strBody  & "PRODUCTS SUPPORTED: " & "<BR>" & mid(Request("AddProductsNames"),2) & "<BR><BR>"
	end if
	
	strFooter = "<HR><font size=1 face=verdana><b>Hardware / Firmware Change Notification</b><BR>"
	strFooter = strFooter & "This HFCN process is the method to use when necessary to notify the test/development groups of pending changes to hardware or firmware<br>"
	strFooter = strFooter & "For more information, please <a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur.asp"">visit our website</a><BR><BR>"
	strFooter = strFooter & "If you have any questions on the HFCN itself, please contact <b><a href=""mailto:" & request("CurrentUserEmail") & """>" & request("CurrentUserName") & "</a></b><BR><BR>"
	strFooter = strFooter & "If you have any questions on the HFCN System, please contact <a href=""mailto:max.yu@hp.com"">Martinez, Efren</a>"
	strFooter = strFooter & "</font>"

	'Create Database Connection
	dim cn
	dim cm
	dim p
	dim ErrorsFound
	dim strSuccess
	dim rowschanged
	dim DisplayedID
	dim NewMilestone
	dim ProdList
	dim strProd
	
		
	set cn = server.CreateObject("ADODB.Connection")
	set cm = server.CreateObject("ADODB.Command")
	
	cn.ConnectionString = Application("PDPIMS_ConnectionString")  
	cn.Open

	cn.BeginTrans
	
	cm.ActiveConnection = cn


	dim RootID
	
	ErrorsFound = false
	
	if instr(Request("cboDeliverable"),":") = 0 then
		RootId = Request("cboDeliverable")
	else
		RootId = left(Request("cboDeliverable"),instr(Request("cboDeliverable"),":")-1)
	end if
	if left(RootID,3) = "ADD" then
		cm.CommandText = "spAddDelRootHFCN"
		cm.CommandType =  &H0004
	
		Set p = cm.CreateParameter("@Name", 200, &H0001, 120)
		p.value = left(Request("txtDeliverableName"),120)
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@CategoryID", 3, &H0001)
		p.Value = request("cboCategory")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@DevManager", 3, &H0001)
		p.Value = request("cboDeveloper")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
		cm.Parameters.Append p
	
		cm.Execute rowschanged
	
		if rowschanged <> 1 then
			Response.Write "<BR>Error Adding Root"
			ErrorsFound = true
			RootID = 0
		else
			Response.Write "<BR>Root Added"
			RootID = cm("@NewID")
		end if
	

		if not ErrorsFound then
	
			'Add Operating Systems Ind
			set cm = nothing
			set cm = server.CreateObject("ADODB.Command")
					            
			cm.ActiveConnection = cn
			cm.CommandText = "spAddOS2DelRoot"
			cm.CommandType = &H0004
		                
			Set p = cm.CreateParameter("@OSID", 3, &H0001)
			p.Value = 16
			cm.Parameters.Append p
		                    
			Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
			p.Value = RootID
			cm.Parameters.Append p
		                    
			cm.Execute rowschanged
		                    
			If rowschanged = 0 Then
				Response.Write "<BR>Error Adding OS to Root"
				ErrorsFound = true
			else
				Response.Write "<BR>OS Added to Root"		
			End If
			
			Set cm = Nothing
		end if
		
		'Add Languages
		if not errorsfound then
			set cm = nothing
       		set cm = server.CreateObject("ADODB.Command")
				            
			cm.ActiveConnection = cn
			cm.CommandText = "spAddLang2DelRootWeb"
			cm.CommandType = &H0004
	              
	                
			Set p = cm.CreateParameter("@LangID", 3, &H0001)
            p.Value = 58
            cm.Parameters.Append p
	                    
			Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
			p.Value = RootID
			cm.Parameters.Append p
	
 			Set p = cm.CreateParameter("@Title", 202, &H0001, 1000)
			p.Value = left(request("txtDeliverableName"),1000)
			cm.Parameters.Append p
	                    
 			Set p = cm.CreateParameter("@Description", 202, &H0001, 1000)
			p.Value = ""
			cm.Parameters.Append p
	                  
	                    
			cm.Execute rowschanged
	                    
			If rowschanged = 0 Then
				Response.Write "<BR>Error Adding Languages to Root"
				ErrorsFound = true
			else
				Response.Write "<BR>Languages Added to Root"		
			End If
			
			Set cm = Nothing
		end if
	
	end if


	if not errorsfound then
		if request("ID") = "" then 'Adding 
			set cm = nothing
       		set cm = server.CreateObject("ADODB.Command")
				            
			cm.ActiveConnection = cn
			cm.CommandText = "spAddDeliverableVersionHFCN"
			cm.CommandType = &H0004
	              
			Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
			p.Value = RootID
			cm.Parameters.Append p
	
 			Set p = cm.CreateParameter("@Version", 200, &H0001, 20)
			p.Value = left(request("txtVersion"),20)
			cm.Parameters.Append p
	                    
			Set p = cm.CreateParameter("@Vendor", 3, &H0001)
			p.Value = request("cboVendor")
			cm.Parameters.Append p

 			Set p = cm.CreateParameter("@VendorVersion", 200, &H0001, 30)
			p.Value = ucase(left(request("txtVendorVersion"),30))
			cm.Parameters.Append p
	                    
			Set p = cm.CreateParameter("@Developer", 3, &H0001)
			p.Value = request("cboDeveloper")
			cm.Parameters.Append p

 			Set p = cm.CreateParameter("@HFCNNotify", 200, &H0001, 255)
			p.Value = left(request("txtEmail"),255)
			cm.Parameters.Append p
	                    
 			Set p = cm.CreateParameter("@HFCNLocation", 200, &H0001, 1000)
			p.Value = left(request("txtLocation"),1000)
			cm.Parameters.Append p

 			Set p = cm.CreateParameter("@OtherDependencies", 200, &H0001, 3000)
			p.Value = left(request("txtDependencies"),3000)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Changes", 201, &H0001, 2147483647)
			p.value = Request("txtChanges")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Title", 201, &H0001, 2147483647)
			p.value = Request("txtTitle") & vbcrlf & vbcrlf & request("txtSpecial")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@NewID", 3,  &H0002)
			cm.Parameters.Append p
	
			cm.Execute rowschanged
	
			if rowschanged < 1 then
				Response.Write "<BR>Error Adding Version"
				ErrorsFound = true
				DisplayedID = 0
			else
				DisplayedID = cm("@NewID")
				Response.Write "<BR>Version Added"		
				cn.Execute "spUpdateDeliverableFilename " & DisplayedID & "," & " 'HFCN_" & DisplayedID & "'",rowschanged
			end if
	                    
			'Add OTS
			
			
			
			if request("OTS2ADD") <> "" and (not ErrorsFound) then
		
				'Add OTS
				dim OTSList
				dim strOTS
							
				if len(request("OTS2ADD")) > 0 then
					if right(request("OTS2ADD"),1) <> "," then
						OTSList = request("OTS2ADD") & ","
					else
						OTSList = request("OTS2ADD") 
					end if
				else
					OTSList = ""
				end if
			
				set cm = nothing
				do while (instr(OTSList,",") and (not ErrorsFound))
					strOTS = left(OTSList,instr(OTSList,",")-1)
		
	         		set cm = server.CreateObject("ADODB.Command")
					            
					cm.ActiveConnection = cn
					cm.CommandText = "spLinkOTS2Version"
					cm.CommandType = &H0004
		                
					Set p = cm.CreateParameter("@OTSNumber",200, &H0001,7)
					p.Value = left(strOTS,7)
					cm.Parameters.Append p
		                    
					Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
					p.Value = DisplayedID
					cm.Parameters.Append p
			                    
					cm.Execute rowschanged
					If rowschanged = 0 Then
						Response.Write "<BR>Error Adding OTS"
						ErrorsFound = true
					else
						Response.Write "<BR>OTS added to version"		
					End If
					Set cm = Nothing
					OTSList = mid(OTSList,instr(OTSList,",")+1)
				loop
			end if
			

			'Add Milestones
			
			if not ErrorsFound then
		
				set cm = nothing
	
	         	set cm = server.CreateObject("ADODB.Command")
				            
	            cm.ActiveConnection = cn
	            cm.CommandText = "spAddMilestones2DeliverableWeb"
				cm.CommandType = &H0004
	                
				Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
	            p.Value = DisplayedID
	            cm.Parameters.Append p
		                    
	            Set p = cm.CreateParameter("@Milestone",200, &H0001,30)
	            p.Value = "Development"
	            cm.Parameters.Append p
	                    
	            Set p = cm.CreateParameter("@MilestoneOrder", 2, &H0001)
	            p.Value = 1
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Planned", 135, &H0001)
	            p.Value = Date()
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Actual", 135, &H0001)
	            p.Value = Date()
		        cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Status", 3, &H0001)
	            p.Value = 3
		        cm.Parameters.Append p
		         
	            Set p = cm.CreateParameter("@ReportMilestone", 16, &H0001)
	            p.Value = 1
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@WorkflowDefinitionID", 3, &H0001)
	            p.Value = 7
	            cm.Parameters.Append p
				
	            Set p = cm.CreateParameter("@NewID", 3, &H0002)
	            cm.Parameters.Append p

		        cm.Execute rowschanged
	            If rowschanged = 0 Then
					Response.Write "<BR>Error Adding Milestones"
	                 ErrorsFound = true
				else
					Response.Write "<BR>Milestone added to version"		
	            End If
					
				NewMilestone = cm("@NewID")
				Set cm = Nothing
			end if


			'Add Operating Systems
			if not ErrorsFound then

				set cm = nothing
          		set cm = server.CreateObject("ADODB.Command")
				            
				cm.ActiveConnection = cn
	            cm.CommandText = "spLinkOS2Version"
				cm.CommandType = &H0004
	                
				Set p = cm.CreateParameter("@OSID", 3, &H0001)
				p.Value = 16
				cm.Parameters.Append p
	                    
				Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
	             p.Value = DisplayedID
				cm.Parameters.Append p
	                    
				cm.Execute rowschanged
	                    
				If rowschanged = 0 Then
					Response.Write "<BR>Error Adding OS to version"
					ErrorsFound = true
				else
					Response.Write "<BR>OS added to version"		
				End If
				
				Set cm = Nothing
				
			end if
	
			'Add Languages
			if not ErrorsFound then
	
				set cm = nothing
         		set cm = server.CreateObject("ADODB.Command")
				            
				cm.ActiveConnection = cn
				cm.CommandText = "spLinkLanguage2Version"
				cm.CommandType = &H0004
	                
				Set p = cm.CreateParameter("@LangID", 3, &H0001)
		        p.Value = 58
	            cm.Parameters.Append p
	                    
				Set p = cm.CreateParameter("@DeliverableVersionID", 3, &H0001)
				p.Value = DisplayedID
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@MilestoneID", 3, &H0001)
				p.Value = NewMilestone
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@WorkflowComplete", 16, &H0001)
				p.Value = 1
				cm.Parameters.Append p
                
				Set p = cm.CreateParameter("@PartNumber", 200, &H0001, 71)
				p.value = ""
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@CDPartNumber", 200, &H0001, 71)
				p.value = ""
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@CDKitNumber", 200, &H0001, 71)
				p.value = ""
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SoftpaqNumber", 200, &H0001, 10)
				p.value = ""
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Supersedes", 200, &H0001, 2048)
				p.value = ""
				cm.Parameters.Append p

				cm.Execute rowschanged
	                    
				If rowschanged = 0 Then
					Response.Write "<BR>Error Adding Language to Version"
	                  ErrorsFound = true
				else
					Response.Write "<BR>Languages added to version"		
				End If
				Set cm = Nothing
			end if

			if not ErrorsFound then
				cn.Execute "spUpdateDeliverableLocation " & DisplayedID, rowschanged
				If rowschanged = 0 Then
					Response.Write "<BR>Error Updating Location"		
				      ErrorsFound = true
				else
					Response.Write "<BR>Location Updated"		
				End If
				Set cm = Nothing
			end if
			
	
	
			'Add Products
			if not ErrorsFound then
				if len(Request("AddProducts")) > 0 then
					ProdList = Request("AddProducts")
					if right(ProdList,1) <> "," then
						ProdList = ProdList &  ","
					end if
				else
					prodList = ""
				end if
		
				set cm = nothing
				do while (instr(ProdList,",") and (not ErrorsFound))
					strProd = left(ProdList,instr(ProdList,",")-1)
	
				 	set cm = server.CreateObject("ADODB.Command")
				            
					cm.ActiveConnection = cn
					cm.CommandText = "spLinkVersion2Product"
					cm.CommandType = &H0004
	                
					Set p = cm.CreateParameter("@ProductVersionID", 3, &H0001)
					p.Value = strProd
					cm.Parameters.Append p
	                    
		            Set p = cm.CreateParameter("@DeliverableRootID", 3, &H0001)
					p.Value = DisplayedID
					cm.Parameters.Append p
		                    
					cm.Execute rowschanged
	                    
					If rowschanged = 0 Then
						Response.Write "<BR>Error Adding Products to Version"
						  ErrorsFound = true
					else
						Response.Write "<BR>Products added to version"		
					End If
					Set cm = Nothing
					ProdList = mid(ProdList,instr(ProdList,",")+1)
				loop
			end if

		
		else 'Updating
		
		
		end if
	end if
	
	if ErrorsFound then
		cn.RollbackTrans
		strSuccess = "0"
		Response.Write "<BR><BR>Unable to save this HFCN."
	else
		cn.CommitTrans
		strSuccess = "1"
		
		
		Response.Write "HFCN saved.  Sending email...<BR>"
		
		'Send Email
		Dim oMessage
		Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
		'Set oMessage.Configuration = Application("CDO_Config")		

		if request("CurrentUserEmail") = "" then
			oMessage.From = "max.yu@hp.com"
		else
			oMessage.From = request("CurrentUserEmail")
		end if

		oMessage.To= strTO
		oMessage.Subject = strSubject
				
		strBody = "<font face=Arial size=2>HFCN ID: " & DisplayedID & "<BR>" & strBody & "</font>"
		oMessage.HTMLBody = strBody & strFooter
		oMessage.DSNOptions = cdoDSNFailure
		oMessage.Send 
		Set oMessage = Nothing 
	end if
	
	set cm = nothing
	set cn=nothing
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
