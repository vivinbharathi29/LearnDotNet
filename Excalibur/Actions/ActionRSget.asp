<!--#include file="../_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "getItem" ) %>


<script runat="server" language="vbscript">


function getItem(ID) 
	on error resume next 

	dim cn 
	dim rs 
	dim strConnect
	dim i
	
	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.open
	

	rs.Open "spListActionRoadmap " & clng(ID),cn,adOpenForwardOnly
	getItem = "<SELECT style=""width=100%"" id=cboRoadmap name=cboRoadmap><option value=""0"" selected>TBD</Option>"
	i=0	do while not rs.EOF
		i=i+1
		getItem = getItem & "<OPTION value=""" & rs("ID") & """>" & i & ". " &  rs("summary") & "</OPTION>"
		rs.MoveNext
	loop
	getItem=getItem & "</select>"
	rs.Close
	
	rs.Open "spGetProductVersion " & clng(ID),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		getItem = getItem & "<INPUT type=""hidden"" id=txtDefaultNotify name=txtDefaultNotify value="""">"
	else
		getItem = getItem & "<INPUT type=""hidden"" id=txtDefaultNotify name=txtDefaultNotify value=""" & rs("ActionNotifyList") & """>" 
	end if
	
	rs.Close
	set rs = nothing
	cn.Close
	set cn=nothing

end function 


</script>
