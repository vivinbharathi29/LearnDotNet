<!--#include file="../../_ScriptLibrary/jsrsServer.inc"--> 

<% jsrsDispatch( "getRoot" ) %>

<script runat="server" language="vbscript">

function getRoot(ID)
	on error resume next 

	
	dim cn 
	dim rs 
	dim strConnect
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	rs.open "spGetRootID " & ID ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		getRoot = 0
	else
		getRoot = rs("ID")
	end if
	rs.Close

	set rs = nothing
	cn.Close
	set cn=nothing
	

end function 

</script>
