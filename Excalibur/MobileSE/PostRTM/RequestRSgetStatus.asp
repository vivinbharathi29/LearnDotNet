<!--#include file="../../_ScriptLibrary/jsrsServer.inc"--> 

<% jsrsDispatch( "getStatus" ) %>

<script runat="server" language="vbscript">

function getStatus(VersionID)
	on error resume next 

	dim cn 
	dim rs 
	dim strConnect
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'getStatus="whatever"
	
	getStatus = "<SELECT style=""width=100%"" id=cboVersionStatus name=cboVersionStatus>"
	rs.open "spGetDeliverableVersionProperties " & VersionID ,cn,adOpenForwardOnly
	if not rs.EOF then
		if trim(rs("PostRTMStatus")) = "0" then
			getStatus = getStatus + "<Option selected value=0>Default</Option>"
			getStatus = getStatus + "<OPTION value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "1" then
			getStatus = getStatus + "<OPTION selected value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=0>Default</OPTION>"
			getStatus = getStatus + "<OPTION value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "2" then
			getStatus = getStatus + "<OPTION selected value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=0>Default</OPTION>"			getStatus = getStatus + "<OPTION value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "3" then
			getStatus = getStatus + "<OPTION selected value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=0>Default</OPTION>"
			getStatus = getStatus + "<OPTION value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "4" then
			getStatus = getStatus + "<OPTION selected value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=0>Default</OPTION>"
			getStatus = getStatus + "<OPTION value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "5" then
			getStatus = getStatus + "<OPTION selected value=5>Hold (Issue Pending)</OPTION>"
			getStatus = getStatus + "<OPTION value=0>Default</OPTION>"
			getStatus = getStatus + "<OPTION value=1>New release</OPTION>"
			getStatus = getStatus + "<OPTION value=2>In Test</OPTION>"
			getStatus = getStatus + "<OPTION value=3>Passed</OPTION>"
			getStatus = getStatus + "<OPTION value=4>Failed</OPTION>"
			getStatus = getStatus + "<OPTION value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=8>Passed - Approved</OPTION>"
			getStatus = getStatus + "<OPTION value=9>Failed - disapproved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "6" then
			getStatus = getStatus + "<OPTION selected value=6>Text/CVA Only</OPTION>"
			getStatus = getStatus + "<OPTION value=7>Text/CVA Only - Approved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "7" then
			getStatus = getStatus + "<OPTION selected value=7>Text/CVA Only - Approved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "8" then
			getStatus = getStatus + "<OPTION selected value=8>Passed - Approved</OPTION>"
		elseif trim(rs("PostRTMStatus")) = "9" then
			getStatus = getStatus + "<OPTION selected value=9>Failed - disapproved</OPTION>"
		end if
	end if

	rs.Close
	getStatus = getStatus & "</SELECT>"

	set rs = nothing
	cn.Close
	set cn=nothing
	
end function 

</script>
