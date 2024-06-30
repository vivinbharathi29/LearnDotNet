<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%

Server.ScriptTimeout = 5400

	dim cn
	dim rs 
	dim strSQL
	dim strSQL1
	dim strSQL2
	dim blnFailed
	dim strTable1
	dim strTable2
	dim strOrderByField
	dim strRowCount1
	dim strRowCount2
	
	
	strtable1 = "datawarehouse.dbo.snapshot_OTS_User_Profile"
	strtable2 = "prs.dbo.ots_OTS_User_Profile"
	strOrderByField = "[UserName]"

'	strtable1 = "prs.dbo.SI_Observation_Tracking"
'	strtable2 = "prs.dbo.SI_Observation_Tracking_New"
'	strOrderByField = "[observationid]"
	
	
'	strtable1 = "datawarehouse.dbo.SI_DevTeam_Observations"
'	strtable2 = "HOUSIREPORT01.SIO.dbo.SI_DevTeam_Observations"
'	strOrderByField = "[observation id]"

'	strtable1 = "datawarehouse.dbo.SI_action"
'	strtable2 = "HOUSIREPORT01.SIO.dbo.SI_action"
'	strOrderByField = "[object id]"
	
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	set rs1 = server.CreateObject("ADODB.Recordset")
	set rs2 = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") 
	cn.IsolationLevel=256
	cn.Open

    strSQl = "Select count(1) as RecordCount " & _
             "From " & strTable1 & " with (NOLOCK) "
    rs.open strSQL,cn
    rowcount1 = rs("RecordCount")
    rs.close

    strSQl = "Select count(1) as RecordCount " & _
             "From " & strTable2 & " with (NOLOCK) "
    rs.open strSQL,cn
    rowcount2 = rs("RecordCount")
    rs.close
    
    
    if rowcount1 > rowcount2 then
        response.Write "<BR>skipping " & rowcount1 - rowcount2
        strSQl1 = "Select * " & _
                 "from " & strtable1 & " with (NOLOCK) " & _
                 "where " & strOrderByField & " in( " & _
                 "                         Select " & strOrderByField & " " & _
                 "                         from " & strtable2 & ") " & _
                 "order by " & strOrderByField
        strSQl2 = "Select * " & _
                 "from " & strtable2 & " with (NOLOCK) " & _
                 "where " & strOrderByField & " in( " & _
                 "                         Select " & strOrderByField & " " & _
                 "                         from " & strtable1 & ") " & _
                 "order by " & strOrderByField
    elseif rowcount2 > rowcount1 then
        response.Write "<BR>skipping " & rowcount2 - rowcount1
        strSQl1 = "Select * " & _
                 "from " & strtable2 & " with (NOLOCK) " & _
                 "where " & strOrderByField & " in( " & _
                 "                         Select " & strOrderByField & " " & _
                 "                         from " & strtable1 & ") " & _
                 "order by " & strOrderByField
        strSQl2 = "Select * " & _
                 "from " & strtable1 & " with (NOLOCK) " & _
                 "where " & strOrderByField & " in( " & _
                 "                         Select " & strOrderByField & " " & _
                 "                         from " & strtable2 & ") " & _
                 "order by " & strOrderByField
    else
        strSQl1 = "Select * " & _
                 "from " & strtable1 & " with (NOLOCK) " & _
                 "order by " & strOrderByField
        strSQl2 = "Select * " & _
                 "from " & strtable2 & " " & _
                 "order by " & strOrderByField
    end if
    response.Write  strSQL1
    response.Flush
    rs1.open strSQL1,cn
    rs2.open strSQL2,cn
    xit = 0
    do while not rs1.eof
        xit = xit  + 1
      ' if xit=100 then
      '  exit do
      ' end if
        for i = 0 to rs1.fields.count -1
            if rs1.fields(i).value <> rs2.fields(i).value then
                if rs1.fields(i).name <> "dateCreated" and rs1.fields(i).name <> "dateModified" and rs1.fields(i).name <> "currwaitdays" and rs1.fields(i).name <> "target_date1" and rs1.fields(i).name <> "dsp_daysOpen" then
                    response.Write "<BR>" & rs1.fields(0).value & ":  " & rs1.fields(i).name & "|" & rs1.fields(i).value & "|" & rs2.fields(i).value
                end if
            end if
        next
        
        response.Flush
    
        rs1.movenext
        rs2.movenext
    loop

    set rs = nothing
    rs1.close
    set rs1 = nothing
    rs2.close
    set rs2 = nothing
	cn.Close
	set cn = nothing

%>
</BODY>
</HTML>
