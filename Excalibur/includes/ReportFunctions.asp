<%

Function ScrubSQL(strWords)     dim badChars     dim newChars 
'	strWords=replace(strWords,"'","''")
    badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update")     newChars = strWords 
    for i = 0 to uBound(badChars)         newChars = replace(newChars, badChars(i), "")     next 
    ScrubSQL = newChars 
End Function 


Function ProcessProductGroupsList( ProductGroups )
	if ProductGroups <> "" then		dim ProductGroupsArray		dim ProductGroupArray		dim strProductGroup		dim lastProductGroup		dim strProductGroupFilter		dim strCycleList		dim strSql				ProductGroupsArray = split(ProductGroups,",")		lastProductGroup = 0		strProductGroupFilter = ""		strCycleList = "" 		for each strProductGroup in ProductGroupsArray			if instr(strProductGroup,":")>0 then				ProductGroupArray = split(strProductGroup,":")				if trim(lastproductgroup) <> "0" and trim(ProductGroupArray(0)) <> "2" and lastproductgroup <> trim(ProductGroupArray(0)) then					strProductGroupFilter = strProductGroupFilter & " ) and  "				end if				if trim(lastproductgroup) <> trim(ProductGroupArray(0)) then					if trim(ProductGroupArray(0)) = "1" then						strProductGroupFilter = strProductGroupFilter & " ( pv.partnerid = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					elseif trim(ProductGroupArray(0)) = "2" then						strCycleList = strCycleList & "," & clng(ProductGroupArray(1))					elseif trim(ProductGroupArray(0)) = "3" then						strProductGroupFilter = strProductGroupFilter & " ( pv.devcenter = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					elseif trim(ProductGroupArray(0)) = "4" then						strProductGroupFilter = strProductGroupFilter & " ( pv.productstatusid = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					end if				else					if trim(ProductGroupArray(0)) = "1" then						strProductGroupFilter = strProductGroupFilter & " or pv.partnerid = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					elseif trim(ProductGroupArray(0)) = "2" then						strCycleList = strCycleList & "," & clng(ProductGroupArray(1))					elseif trim(ProductGroupArray(0)) = "3" then						strProductGroupFilter = strProductGroupFilter & " or pv.devcenter = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					elseif trim(ProductGroupArray(0)) = "4" then						strProductGroupFilter = strProductGroupFilter & " or pv.productstatusid = " & trim(ProductGroupArray(1))						lastproductgroup = trim(ProductGroupArray(0))					end if				end if			end if		next		if strProductGroupFilter <> "" then			strSQl = strSQL & " and ( " & ScrubSQL(strProductGroupFilter) &  ") ) "		end if		if strCycleList <> "" then			strSQl = strSQL & " and pv.id in (Select ProductVersionid from product_program where programid in (" & mid(strCycleList,2) &  ")) "		end if	end if	    ProcessProductGroupsList = strSQL	End Function

'because of no optional perameter in VB Script
Function ProcessProductGroupListGetIds( ProductGroups )
    ProcessProductGroupListGetIds = ProcessProductGroupActive( ProductGroups , " active = 1 ")
End Function

'Herb, PBI 21220, Display with all "active" status.
Function ProcessProductGroupListGetAllIds( ProductGroups )
    ProcessProductGroupListGetAllIds = ProcessProductGroupActive( ProductGroups , " active in (0,1) ")
End Function


Function ProcessProductGroupActive( ProductGroups , strActive)
    Dim strSQL
    Dim strIds
    if trim(strActive) = "" then
        strActive = " active = 1 " 'strActive cannot be empty
    end if
    strSQL = "Select pv.ID FROM ProductVersion pv WHERE " & strActive & _
    ProcessProductGroupsList( ProductGroups )
    
    Dim cn, cm, rs
    Set cn = Server.CreateObject("ADODB.Connection")
   	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	Set rs = server.CreateObject("ADODB.Recordset")
    Set cm = Server.CreateObject("ADODB.Command")
    Set cm.ActiveConnection = cn
    cm.CommandType = 1
    cm.CommandText = strSql
    Set rs = cm.Execute 
	Set cm = Nothing   
	
    Do Until rs.EOF
        strIds = strIds & "," & rs("ID")
        rs.MoveNext
    Loop
	
	rs.Close
	
	ProcessProductGroupActive = strIds
	
	Set rs = Nothing
	Set cn = Nothing

End Function

%>