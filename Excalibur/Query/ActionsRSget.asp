<!--#include file="../_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "ProfileStrings" ) %>


<script runat="server" language="vbscript">




function ProfileStrings(ID) 

on error resume next 


dim cn 
dim rs 
dim i


	dim strProducts
	dim strStatus
	dim strCategories
	'dim Profiles
	
	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	rs.open "spGetReportProfile " & clng(ID) ,cn,adOpenForwardOnly
	
	if rs.EOF and rs.BOF then
		ProfileStrings = "<INPUT type=""hidden"" id=txtQProfileName name=txtQProfileName value="""">"
	else
		ProfileStrings = "<INPUT type=""hidden"" id=txtQProfileName name=txtQProfileName value=""" & rs("profileName") & """><BR>"
		for i = 1 to 69
			ProfileStrings = ProfileStrings & "<INPUT type=""hidden"" id=txtQValue" & i & " name=txtQValue" & i & " value=""" & rs("Value" & i) & """><BR>"
		next
		 
	end if
	rs.Close
	set rs = nothing
	cn.Close
	set cn=nothing

end function 

</script> 

