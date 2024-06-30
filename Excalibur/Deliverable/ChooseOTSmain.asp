<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
TD{
	FONT-FAMILY:Verdana;
	FONT-SIZE:xx-small;
}
</STYLE>
</HEAD>
<BODY bgcolor=Ivory>
<form ID=frmMain method=post action=ChooseOTSSave.asp>
<%
	if trim(request("ID")) = "" then
		Response.Write "Not enough information supplied to process this request."
	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		on error resume next
		rs.Open "spListOTS4Root " & clng(request("ID")),cn,adOpenForwardOnly
		if cn.Errors.count > 0 then
			Response.Write "<font face=verdana size=2 color=red><b>OTS is unavailable</b></font>"
		else
			Response.write "<font size=2 face=verdana><b>Open Observations written against this root deliverable</b></font>"
			if rs.EOF and rs.BOF then
				Response.Write "<BR><font size=2 face=verdana>None</font>"
			else
				Response.Write "<TABLE cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan border=1>"
				Response.Write "<TR bgcolor=wheat><TD>&nbsp;</TD><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Version</b></TD><TD><b>Priority</b></TD><TD><b>Status</b></TD><TD><b>Summary</b></TD></TR>"
				do while not rs.EOF
					strVersion = rs("Version")
					if rs("Revision") <> "" then
						strVersion = strVersion & "," & rs("Revision")
					end if
					if rs("Pass") <> "" then
						strVersion = strVersion & "," & rs("Pass")
					end if
					Response.Write "<TR>"
					if instr(request("OldIDList") & ",","," & trim(rs("ObservationID")) & ",") = 0 then
						Response.Write "<TD><INPUT type=""checkbox"" id=lstObservations name=lstObservations value=""" &  rs("ObservationID") & """></TD>"
					else
						Response.Write "<TD><INPUT checked type=""checkbox"" id=lstObservations name=lstObservations value=""" &  rs("ObservationID") & """></TD>"
					end if
					Response.Write "<TD nowrap>" & rs("ObservationID") & "</TD>"
					Response.Write "<TD nowrap>" & rs("Product") & "</TD>"
					Response.Write "<TD>" & strVersion & "</TD>"
					Response.Write "<TD>" & rs("Priority") & "</TD>"
					Response.Write "<TD>" & rs("State") & "</TD>"
					Response.Write "<TD>" & rs("Summary") & "</TD>"
					Response.Write "</TR>"
					rs.MoveNext
				loop
				Response.Write "</table>"
			end if
			rs.Close
		end if


		rs.Open "spListOTSRelatedToVersion " & clng(request("VersionID")) & "," & clng(request("UserID")),cn,adOpenForwardOnly
		if cn.Errors.count > 0 then
			Response.Write "<font face=verdana size=2 color=red><b>OTS is unavailable</b></font>"
		else
			Response.write "<font size=2 face=verdana><b><BR><BR>Other open observations assigned to me, the developer, or the manager of this version.</b></font>"
			if rs.EOF and rs.BOF then
				Response.Write "<BR><font size=2 face=verdana>None</font>"
			else
				Response.Write "<TABLE cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan border=1>"
				Response.Write "<TR bgcolor=wheat><TD>&nbsp;</TD><TD><b>ID</b></TD><TD><b>Deliverable</b></TD><TD><b>Product</b></TD><TD><b>Version</b></TD><TD><b>Priority</b></TD><TD><b>Status</b></TD><TD><b>Summary</b></TD></TR>"
				do while not rs.EOF
					strVersion = rs("Version")
					if rs("Revision") <> "" then
						strVersion = strVersion & "," & rs("Revision")
					end if
					if rs("Pass") <> "" then
						strVersion = strVersion & "," & rs("Pass")
					end if
					Response.Write "<TR>"
					if instr(request("OldIDList") & ",","," & trim(rs("ObservationID")) & ",") = 0 then
						Response.Write "<TD><INPUT type=""checkbox"" id=lstObservations name=lstObservations value=""" &  rs("ObservationID") & """></TD>"
					else
						Response.Write "<TD><INPUT checked type=""checkbox"" id=lstObservations name=lstObservations value=""" &  rs("ObservationID") & """></TD>"
					end if
					Response.Write "<TD nowrap>" & rs("ObservationID") & "</TD>"
					Response.Write "<TD>" & rs("DeliverableName") & "</TD>"
					Response.Write "<TD nowrap>" & rs("Product") & "</TD>"
					Response.Write "<TD>" & strVersion & "</TD>"
					Response.Write "<TD>" & rs("Priority") & "</TD>"
					Response.Write "<TD>" & rs("State") & "</TD>"
					Response.Write "<TD>" & rs("Summary") & "</TD>"
					Response.Write "</TR>"
					rs.MoveNext
				loop
			end if
			rs.Close
		end if
	
		set rs = nothing
		cn.Close
		set cn = nothing
	end if
%>
</FORM>
</BODY>
</HTML>
