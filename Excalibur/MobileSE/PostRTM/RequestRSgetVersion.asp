<!--#include file="../../_ScriptLibrary/jsrsServer.inc"--> 

<% jsrsDispatch( "getVersion" ) %>

<script runat="server" language="vbscript">

function getVersion(ID)
	on error resume next 

	
	dim cn 
	dim rs 
	dim strConnect
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	if instr(ID, "*") > 0 then
		ID = left(ID, instr(ID, "*")-1)
	end if
	
	rs.open "spListDeliverableVersions " & ID ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		'getVersion = "Version Not Found : " & ID	
		getVersion = "<SELECT style=""width=100%"" id=cboVersion name=cboVersion LANGUAGE=javascript onclick=""return VersionParent_onclick()"" onkeyup=""return VersionParent_onclick()""><OPTION selected value=0>Next Release</OPTION>"
	else
		getVersion = "<SELECT style=""width=100%"" id=cboVersion name=cboVersion  LANGUAGE=javascript onclick=""return VersionParent_onclick()"" onkeyup=""return VersionParent_onclick()""><OPTION selected value=0>Next Release</OPTION>"
		do while not rs.EOF
			getVersion = getVersion & "<Option value=""" & rs("VersionID") & """>" & rs("Version") & ", " & rs("Revision") & ", " & rs("Pass") & "</Option>"
			rs.MoveNext
		loop
	end if
	rs.Close
	getVersion = getVersion & "</SELECT>"

	set rs = nothing
	cn.Close
	set cn=nothing
	

end function 

</script>
