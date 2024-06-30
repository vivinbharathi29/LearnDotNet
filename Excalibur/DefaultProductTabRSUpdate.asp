<script runat="server" language="vbscript">

    dim EmployeeID, strList, returnValue
    strList = request.QueryString("List")
    EmployeeID = request.QueryString("CurrentUserID")

    dim cn 
	dim cm 
	dim i

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
    
    cn.BeginTrans
    
    set cm = server.CreateObject("ADODB.Command")
				            
	cm.ActiveConnection = cn
    cm.CommandText = "spUpdateDefaultProductTab"
	cm.CommandType = &H0004
	                
	Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
	p.Value = clng(EmployeeID)
	cm.Parameters.Append p
	                    
	Set p = cm.CreateParameter("@DeliverableRootID", 200, &H0001,2000)
	p.Value = trim(left(strList,2000))
	cm.Parameters.Append p
	                    
	cm.Execute rowschanged

	if rowschanged = 1 then
		returnValue = "1"
		cn.committrans
	else
		returnValue = "0"
		cn.rollbacktrans
	end if
    
	set cm = nothing
	cn.close	
	set cn = nothing

    response.Write returnValue
    

</script>
