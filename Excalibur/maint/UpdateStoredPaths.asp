<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD{
	FONT-Size: xx-small;
	FONT-FAMILY: Verdana;
}

</STYLE>
<BODY>
<%
		dim strText
		dim cn
		dim rs
		dim RowsUpdated
		dim blnError
		dim strNewPath


%>


<form id=frmMain action=UpdateStoredPaths.asp method=post>
<%if request("txtFind") = "" and request("txtReplace") = "" and request("txtSQL") = "" then%>
	<font size=2 face=verdana color=red>Note: This only updates 5000 at a time.</font><BR>
	<TABLE>
		<TR>
			<TD>Find:</TD>
			<TD><INPUT width="80%" id=txtFind name=txtFind style="WIDTH: 509px; HEIGHT: 22px" size=65></TD>
		</TR>
		<TR>
			<TD>Replace:</TD>
			<TD><INPUT width="80%" id=txtReplace name=txtReplace style="WIDTH: 510px; HEIGHT: 22px" size=65></TD>
		</TR>
	</TABLE>
	<INPUT type="submit" value="Next" id=button1 name=button1>
<%elseif request("txtSQL") = "" then

	Response.Write "<INPUT type=""submit"" value=""Next"" id=button1 name=button1>"
	
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
	
		rs.Open "spListVersionsFindPaths '" & request("txtFind") & "'",cn,adOpenStatic
		if not (rs.EOF and rs.BOF) then
			Response.Write "<TABLE border=1>"
		end if
		i=0
		do while not rs.EOF
			i=i+1
			if i<5001 then
'				exit do
'			end if
			strNewPath = replace(lcase(rs("Path")),lcase(request("txtFind")),lcase(request("txtReplace")))
			Response.Write "<TR>"
			Response.Write "<TD>" & rs("ID") & "</TD>"
			Response.Write "<TD>" & rs("FieldName") & "</TD>"
			Response.Write "<TD nowrap><B>OLD: </b>" & rs("Path") & "<BR>"
			Response.Write "<B>NEW: </b>" & strNewPath  & "<BR>"
			if (trim(rs("DeliverableName") & "") = "Server") then
			    Response.Write "<INPUT style=""display:none;width:900"" type=""text"" id=txtSQL name=txtSQL value=""" & "Update Server Set " & rs("FieldName") & " = '" & strNewPath & "' Where ID = " & rs("ID") & """>"
			elseif trim(rs("DeliverableName") & "") <> "" then
			    Response.Write "<INPUT style=""display:none;width:900"" type=""text"" id=txtSQL name=txtSQL value=""" & "Update DeliverableVersion set " & rs("FieldName") & " = '" & strNewPath & "' Where ID = " & rs("ID")  & """>"
			else
			    Response.Write "<INPUT style=""display:none;width:900"" type=""text"" id=txtSQL name=txtSQL value=""" & "Update ProductVersion set " & rs("FieldName") & " = '" & strNewPath & "' Where ID = " & rs("ID")  & """>"
			end if
			Response.Write "</TD>"
			Response.Write "<TD>" & rs("DeliverableName") & "</TD>"
			Response.Write "<TD>" & rs("Version") & "</TD>"
			Response.Write "<TD>" & rs("Revision") & "</TD>"
			Response.Write "<TD>" & rs("Pass") & "</TD>"
			Response.Write "</TR>"
            end if
			rs.MoveNext
		loop
		if not (rs.EOF and rs.BOF) then
			Response.Write "</TABLE>" & "<BR>Total: " & i
		end if
		rs.Close
			
		set rs = nothing
		cn.Close
		set cn = nothing
	
else

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
	
		blnError = false
		'cn.begintrans
		for each strText in request("txtSQL")
			Response.Write strText & "<BR>"
			Response.flush
			cn.execute strText,RowsUpdated
			if RowsUpdated <> 1 then
				Response.Write "<BR><BR>Error Found."
				blnError = true
				exit for
			end if
		next
		
		'if blnError then
		'	cn.rollbacktrans
		'else
		'	cn.committrans
		'end if				

		cn.Close
		set cn = nothing

        Response.Write "<br /><br /><b>Finished!!!!!!!</b><br />"
end if%>
</form>
</BODY>
</HTML>
