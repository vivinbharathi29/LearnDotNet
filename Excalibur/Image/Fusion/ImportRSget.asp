<!--#include file="../../_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "ImageTable" ) %>



<script runat="server" language="vbscript">

function ImageTable(ID) 

on error resume next 


dim cn 
dim rs 
dim cm
dim p
dim strResult
strResult = ""
	
	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	set rs2 = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.open
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListImageDefinitionsFusion"
		

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = ID
	cm.Parameters.Append p
	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

'	rs.open "spListImageDefinitionsByProduct " & ID ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then 	
		strResult =  "<Table ID=""ImageTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>SKU&nbsp;Number</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Software&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Type&nbsp;&nbsp;</b></font></TD></tr>"
		strResult = strResult &  "<TR><TD colspan=4><font size=1 face=verdana>No images found for selected product.</font></TD></TR></table>"
	else
		strResult =  "<Table ID=""ImageTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" checked id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Product&nbsp;Drop</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Comments&nbsp;&nbsp;</b></font></TD></tr>"
		do while not rs.EOF 
				strResult = strResult & "<TR valign=top bgcolor=Ivory><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("ID") & """ checked style=""width:16;height:16;"" type=""checkbox"" id=chkSelected name=chkSelected></td>"
				strResult = strResult & "<TD nowrap style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("ProductDrop") & "&nbsp;</font></TD>"
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Brand") & "&nbsp;</font></TD>"
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("OS") & "&nbsp;</font></TD>"
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Comments") & "&nbsp;</font></TD>"
			rs.MoveNext
		loop
		rs.Close
		strResult = strResult &   "</table>"
	end if
	set rs = nothing
	set rs2 = nothing
	set cn = nothing

	ImageTable = strResult
end function 

</script> 

