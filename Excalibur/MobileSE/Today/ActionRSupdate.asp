<!--#include file="../../_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "ProductApprovers" ) %>



<script runat="server" language="vbscript">




function ProductApprovers(ID,NewList) 

	on error resume next 


	dim cn 
	dim cm
	dim i
	dim strResult

	strResult = ""
	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spUpdateApproverList"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@ProdID", 3,  &H0001)
	p.value =ID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ApproverList", 200, &H0001, 255)
	p.value = left(NewList,255)
	cm.Parameters.Append p
	
	cm.Execute rowschanged

	if rowschanged = 1 then
		strResult = "1"
		cn.committrans
	else
		cn.rollbacktrans
	end if
	
	set cm = nothing
	set cn = nothing
	ProductApprovers = strResult
end function 

</script> 

