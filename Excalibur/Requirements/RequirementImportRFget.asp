<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<BODY>
<%

dim cn 
dim rs 
	
	set cn = server.createobject("ADODB.Connection") 
	set rs = server.createobject("ADODB.Recordset") 
	set rs2 = server.createobject("ADODB.Recordset") 
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.open
	
	rs.open "spListRequirments4Import " & clng(request("ImportID")) & "," & clng(request("ID")) ,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then 	
		Response.Write  "<Table ID=""ReqTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Requirement</b></font></TD><TD><font size=1 face=verdana><b>Specification&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Deliverables&nbsp;&nbsp;</b></font></TD></tr>"
		Response.Write "<TR><TD colspan=4><font size=1 face=verdana>No Additional Requirement found.</font></TD></TR></table>"
	else
		Response.Write  "<Table ID=""ReqTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" checked id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Requirement</b></font></TD><TD><font size=1 face=verdana><b>Specification&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Deliverables&nbsp;&nbsp;</b></font></TD></tr>"
		do while not rs.EOF 
			'if rs("ID") <> 1960 then
				rs2.Open "splistdeliverablesbyrequirement " & clng(rs("ID")) & "," & clng(request("ImportID")), cn, adOpenForwardOnly
				strDeliverables = ""
				do While Not rs2.EOF
					strDeliverables = strDeliverables & "-" & rs2("Name") &  "<BR>"
					rs2.MoveNext
				Loop
				rs2.Close
		
				if strDeliverables = "" then
					strDeliverables = "&nbsp;"
				end if		
				
				Response.Write  "<TR ID=""Row" & rs("ID") & """ valign=top bgcolor=Ivory><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("ProductID") & """ checked style=""width:16;height:16;"" type=""checkbox"" id=chkSelected name=chkSelected></td>"
				Response.Write "<TD nowrap style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana><b>" & rs("Requirement") & "&nbsp;</b></font></TD>"
				Response.Write "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Spec") & "&nbsp;</font></TD>"
				Response.Write	"<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strDeliverables  & "</font></TD></tr>"
			'end if
			rs.MoveNext
		loop
		rs.Close
		Response.Write "</table>"
	end if
	set rs = nothing
	set rs2 = nothing
	set cn = nothing

%>
</BODY>
</HTML>
