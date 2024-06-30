<!--#include file="_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "UpdateSetting" ) %>


<script runat="server" language="vbscript">


function UpdateSetting(strValue, EmployeeID, UserSettingID) 
	on error resume next 
	dim cn 
	dim cm 
	dim i

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
    
    cn.BeginTrans
    
    set cm = server.CreateObject("ADODB.Command")
				            
	cm.ActiveConnection = cn
    cm.CommandText = "spUpdateDefaultProductFilter"
	cm.CommandType = &H0004
	                
	Set p = cm.CreateParameter("@EmployeeID", 3, &H0001)
	p.Value = clng(EmployeeID)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@UserSettingID", 3, &H0001)
	p.Value = clng(UserSettingID)
	cm.Parameters.Append p
	                    
	Set p = cm.CreateParameter("@Value", 200, &H0001,8000)
	p.Value = trim(left(strValue,8000))
	cm.Parameters.Append p
	                    
	cm.Execute rowschanged

	if rowschanged = 1 then
		UpdateSetting = "1"
		cn.committrans
	else
		UpdateSetting = "0"
		cn.rollbacktrans
	end if


	set cm = nothing
	cn.close	
	set cn = nothing
end function



</script>
