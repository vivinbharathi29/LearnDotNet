<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (document.getElementById('pulsarplus').value == 'true')
            {
                parent.window.parent.reloadIgGrid();
                parent.window.parent.closeExternalPopup();
            }
            else if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel(true);
            } else {
                window.returnValue = 1;
                window.parent.close();
            }
        } else {
            document.write("<BR><font size=2 face=verdana>Unable to update program.</font>");
        }
    } else {
        document.write("<BR><font size=2 face=verdana>Unable to update program.</font>");
    }
}





//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim cn
	dim cm
	dim strSuccess
	dim p
	dim rowsUpdated
	dim strItem
	dim strPartNumbers
	dim strSQL
	
	strPartNumbers = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	cn.BeginTrans
	strSuccess = "1"
	if request("txtDisplayedID") = "" then
		'Add
		
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		cm.CommandText = "spAddProgram"	

		Set p = cm.CreateParameter("@Name", 200,  &H0001,30)
		p.Value = left(request("txtName"),30)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@FullName", 200,  &H0001,34)
		p.Value = left(request("txtFullName"),34)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("ProgramGroupID", 3,  &H0001)
		p.Value = clng(request("cboProgramGroup"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("OTSCycleName", 200,  &H0001,40)
		p.Value = left(request("txtOTSCycleName"),40)
		cm.Parameters.Append p
			
		Set p = cm.CreateParameter("@Active", 16,  &H0001)
		if request("chkActive") = "on" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		
		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
		cm.Parameters.Append p

		cm.Execute rowschanged
		NewID = cm("@NewID")
		set cm=nothing

		if cn.Errors.count > 0 then
			strSuccess = "0"
		end if	
		
		
		strItemsAdded = ""
		if strSuccess = "1" and trim(NewID) <> "" then		
			ItemArray = split(request("lstProduct"),",")
			for each strItem in ItemArray
				if trim(strItem) <> "" then
					strItemsAdded = strItemsAdded & "," & trim(strItem)
					cn.Execute "spLinkProductToProgram " & clng(NewID) & "," & clng(strItem), rowsUpdated
					if rowsUpdated <> 1 then		
						strSuccess = "0"
						exit for			
					end if
				end if
			next
		end if

		if strItemsAdded <> "" then
			strItemsAdded = mid(strItemsAdded,2)
		end if


	
	else
		'Update
		
		
		dim ItemArray
		dim i
		dim strItemsAdded
		dim strItemsRemoved
		
		strHistory  = ""
	

			'Added
		ItemArray = split(request("lstProduct"),",")
		strItemsAdded = ""
		for i = lbound(ItemArray) to ubound(ItemArray)
			if trim(ItemArray(i)) <> "" then
				if instr("," & trim(request("tagProduct")) & ",","," & trim(ItemArray(i)) & ",") = 0 then
					strItemsAdded = strItemsAdded & "," & trim(ItemArray(i))
				end if
			end if
		next
		if strItemsAdded <> "" then
			strItemsAdded = mid(strItemsAdded,2)
		end if
			
			'Removed
		ItemArray = split(request("tagProduct"),",")
		strItemsRemoved = ""
		for i = lbound(ItemArray) to ubound(ItemArray)
			if trim(ItemArray(i)) <> "" then
				if instr(", " & trim(request("lstProduct")) & ",",", " & trim(ItemArray(i)) & ",") = 0 then
					strItemsRemoved = strItemsRemoved & "," & replace(trim(ItemArray(i)),",","")
				end if
			end if
		next
		if strItemsRemoved <> "" then
			strItemsRemoved = mid(strItemsRemoved,2)
		end if
		
		Response.Write "Added:" & strItemsAdded & "<BR>"
		Response.Write "Removed:" & strItemsRemoved & "<BR>"
		Response.Write "Current:" &  request("chkCurrent") & "<BR>"
		Response.Write "Active:" &  request("chkActive") & "<BR>"
		Response.Write "Name:" & request("txtName") & "<BR>"


		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		cm.CommandText = "spUpdateProgram"	

		Set p = cm.CreateParameter("@ID", 3,  &H0001)
		p.Value = clng(request("txtDisplayedID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Name", 200,  &H0001,30)
		p.Value = left(request("txtName"),30)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@FullName", 200,  &H0001,34)
		p.Value = left(request("txtFullName"),34)
		cm.Parameters.Append p
        
		Set p = cm.CreateParameter("ProgramGroupID", 3,  &H0001)
		p.Value = clng(request("cboProgramGroup"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("OTSCycleName", 200,  &H0001,40)
		p.Value = left(request("txtOTSCycleName"),40)
		cm.Parameters.Append p
			
		Set p = cm.CreateParameter("@Active", 16,  &H0001)
		if request("chkActive") = "on" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		cm.Execute rowschanged

		set cm=nothing

		if cn.Errors.count > 0 then
			strSuccess = "0"
		end if	

		'Add New Products
		ItemArray = split(strItemsAdded,",")		
		for i = lbound(ItemArray) to ubound(ItemArray)
			if trim(ItemArray(i)) <> "" then
				cn.Execute "spLinkProductToProgram " & clng(request("txtDisplayedID")) & "," & clng(ItemArray(i)), rowsUpdated
				if rowsUpdated <> 1 then		
					strSuccess = "0"
					exit for			
				end if
			end if
		next
		
		'Remove Products
		ItemArray = split(strItemsRemoved,",")		
		for i = lbound(ItemArray) to ubound(ItemArray)
			if trim(ItemArray(i)) <> "" then
				cn.Execute "spUnLinkProductFromProgram " & clng(request("txtDisplayedID")) & "," & clng(ItemArray(i)), rowsUpdated
				if rowsUpdated <> 1 then		
					strSuccess = "0"
					exit for			
				end if
			end if
		next

	end if
	
	if trim(request("tagLead")) <> trim(request("cboLead")) then
        if trim(request("txtDisplayedID")) = "" then
	        cn.execute "spUpdateProgramLeadProduct " & clng(NewID) & "," & clng(request("cboLead"))
        else
	        cn.execute "spUpdateProgramLeadProduct " & clng(request("txtDisplayedID")) & "," & clng(request("cboLead"))
	    end if
    end if

	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if


'	if clng(request("cboCycleType")) = 1  and (strItemsAdded <> "" or (request("txtOldCycleName") <> request("txtOTSCycleName")and request("txtOldCycleName") <> "" ) ) then
		'*****ONLY UPDATE OTS FOR COMMERCIAL CYCLES********


'		set cnOTS = server.CreateObject("ADODB.Connection")
'		cnOTS.ConnectionString = Application("OTS_ConnectionString") 
'		cnOTS.IsolationLevel=256
 '   	cnOTS.ConnectionTimeout = 10
'		cnOTS.Open
'
		
		'If Cycle Name changed, then we need to update OTS with the new name
'		if request("txtOldCycleName") <> request("txtOTSCycleName") and request("txtOldCycleName") <> "" then
'			
'			strSQL = "Update CyclePlatform " & _
'					 "set platform = '" & trim(scrubsql(request("txtOTSCycleName"))) & "' " & _
'					 "where (cycle='BNB Cycle' or cycle='BNB Common') " & _
'					 "and Platform='" & trim(scrubsql(request("txtOldCycleName") )) & "' " & _
'				 	 "and organizationid=3 " & _
'				 	 "and Operation='Add' " & _
'					 "and partnumber like 'EXC-%'"
'					 
'			Response.write "<BR>" & strSQL
'				cnOTS.Execute strSQL
'	
'		end if
'		
		'Update OTS - Need to link all deliverables for the Added products to the cycle
'		if strItemsAdded <> "" then
'			dim ProductNameList		
'			set rs = server.CreateObject("ADODB.Recordset")
'			rs.Open "Select Distinct DOTSName as Product from productversion where id in (" & strItemsAdded & ");",cn,adOpenStatic
'			ProductNameList = ""
'			do while not rs.EOF	
'				ProductNameList = ProductNameList & ",'" & rs("Product") & "'"
'				rs.MoveNext
'			loop
'			rs.Close
'			set rs = nothing
'			if ProductNameList <> "" then
'				ProductNameList = mid(ProductNameList,2)
'				
'				Response.Write "<BR>"
'				strSQL =  "Insert Into CyclePlatform (Cycle,Platform,PartNumber,Created, Operation,OrganizationID) " & _
'								"(Select distinct 'BNB Common','" & request("txtOTSCycleName") & "', partnumber,GetDate(), 'Add',3 " & _
'								" from cycleplatform " & _
'								" where organizationid = 3 " & _
'								" and platform in (" & ProductNameList & ") " & _
'								" and partnumber like 'EXC-%' " & _
'								" and operation = 'Add' " & _
'								" and partnumber not in (Select Partnumber " & _
'								" from cycleplatform " & _
'								" where organizationid = 3 " & _
'								" and platform in ('" & request("txtOTSCycleName") & "') " & _
'								" and operation = 'Add') )"
'				Response.Write "<BR>" & strSQl
'			
'				cnOTS.Execute strSQl
'			end if
'		end if
'		
'		cnOTS.Close
'		set cnOTS = nothing
'		
'	end if
	cn.Close
	set cn = nothing



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

%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="text" id=pulsarplus name=pulsarplus value="<%=Request("pulsarplus")%>">
</BODY>
</HTML>
