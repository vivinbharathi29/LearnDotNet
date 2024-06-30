<!--#include file="_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "FindRootName" ) %>


<script runat="server" language="vbscript">



function FindRootName(strName, strID) 

    on error resume next 
    
    dim cn 
    dim rs 
	
	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spCountRootDeliverableWithName"
	
	Set p = cm.CreateParameter("@Name", 200, &H0001,120)
	p.Value = left(strName,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	if trim(strID) = "" then
	    p.Value = 0
	else
	    p.value = strID
	end if
	cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	set cm=nothing
	
	FindRootName = rs("RootCount")

	rs.Close
	set rs = nothing
	cn.Close
	set cn=nothing

end function 

</script> 

