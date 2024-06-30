<!--#include file="_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "ProfileSharing" ) %>


<script runat="server" language="vbscript">



function ProfileSharing(ID) 

on error resume next 


	dim cn 

	set cn = server.createobject("ADODB.Connection") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	

	cn.execute "spRemoveSharedProfile " & clng(ID) 
	
	if cn.Errors.count > 0 then
		ProfileSharing = "Error"
	else
		ProfileStrings = ""
	end if
	cn.Close
	set cn=nothing

end function 

</script> 

