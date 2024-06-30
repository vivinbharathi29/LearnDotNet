<%

Function ScrubSQL(strWords) 
'	strWords=replace(strWords,"'","''")
    badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
    for i = 0 to uBound(badChars) 
    ScrubSQL = newChars 
End Function 


Function ProcessProductGroupsList( ProductGroups )
	if ProductGroups <> "" then

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