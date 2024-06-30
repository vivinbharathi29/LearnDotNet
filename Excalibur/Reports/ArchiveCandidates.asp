<%@ Language=VBScript %>
<%
	if request("Type") = "Excel" then
		Response.ContentType = "application/vnd.ms-excel"
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript><!--function cmdUpdate_onclick(){    frmMain.submit();}function window_onload() {
    if (typeof(Reload) != "undefined")
        if (Reload.value == "1")
            window.location.href = "ArchiveCandidates.asp?TeamID=" + TeamID.value;

}
//--></SCRIPT>
</HEAD>
<STYLE>
    Table
    {
        FONT-Family: verdana;
        FONT-Size: xx-small;	
    }
    A:link
    {
        COLOR: blue
    }
    A:visited
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana><b>Archive Candidates</b></font><br><br>
<%
    dim ArchiveArray
    dim strArchive
    dim cn
    dim cnstring
    dim RowsUpdated
    dim blnFailed
    
    if request("chkArchive") <> "" then 
        blnFailed = false
        ArchiveArray =  split(request("chkArchive"),",")
    	set cn = server.CreateObject("ADODB.Connection")
		cnString =Session("PDPIMS_ConnectionString")
    	cn.ConnectionString = cnString
    	cn.Open
    	cn.begintrans
        for each strArchive in ArchiveArray
            if trim(strArchive) <> "" then
                response.Write "Archiving: " & strArchive & "<BR>"
                if trim(request("TeamID")) = "2" then
                    cn.execute "spArchiveDeliverable " & clng(strArchive) & ",1,4",RowsUpdated                
                elseif trim(request("TeamID")) = "1" then
                    cn.execute "spArchiveDeliverable " & clng(strArchive) & ",1,1",RowsUpdated                
                else
                    RowsUpdated = 0
                end if
                if RowsUpdated <> 1 then
                    blnFailed = true
                    exit for
                end if
            end if
        next
        if blnFailed then
            cn.rollbacktrans
            response.write "<BR>Update Failed"
        else
            cn.committrans
            response.write "<BR>Update Complete.<BR><BR><b>Please Wait.  Reloading List...</b>"
        end if
        cn.close
        set cn = nothing    
       response.Write "<input style=""display:none"" id=""Reload"" name=""Reload"" type=""text"" value=""1"">"
       response.write "<input style=""display:none"" id=""TeamID"" name=""TeamID"" type=""text"" value=""" & request("TeamID") & """>"

    else%>
    
    <input id="cmdUpdate" name="cmdUpdate" type="button" value="Archive Selected Versions" onclick="javascript: cmdUpdate_onclick();"><br><br>
<form id=frmMain method=post action="ArchiveCandidates.asp">
<%
	if trim(request("TeamID")) = "" then
		Response.write "No Release Team specified."
	else
		cnString =Session("PDPIMS_ConnectionString")
	
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = cnString
		cn.Open
	
		set rs = server.CreateObject("ADODB.recordset")
		rs.ActiveConnection = cn
		
		rs.Open "spListArchiveCandidates " & clng(request("TeamID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.write "No deliverables found matching the specified criteria."
		else
			Response.Write "<TABLE border=1 bgcolor=ivory>"
			Response.Write "<TR bgcolor=beige><TD>&nbsp;</TD></TD><TD><b>Deliverable</b></TD></TR>"
			do while not rs.EOF
				strVersion = rs("Version") & ""
				if trim(rs("Revision") & "") <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if trim(rs("Pass") & "") <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				Response.Write "<TR><TD><input id=""chkArchive"" name=""chkArchive"" type=""checkbox"" value=""" & rs("ID") & """></td>"
				Response.Write "<TD>" & rs("ID") & " - " & server.htmlencode(rs("DeliverableName") & "") & " [" & server.htmlencode(strVersion) & "]</TD></tr>"
				if left(rs("Path1"),2) = "\\" then
				    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI><a target=_blank href=""" & server.htmlencode(rs("Path1") & "") & """>" & server.htmlencode(rs("Path1") & "")  & "</a></LI>"
				else
				    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI>" & server.htmlencode(rs("Path1") & "")  & "</LI>"
				end if
				if trim(rs("Path2") & "") <> "" then
	    			if left(rs("Path2"),2) = "\\" then
    				    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI><a target=_blank href=""" & server.htmlencode(rs("Path2") & "") & """>" & server.htmlencode(rs("Path2") & "")  & "</a></LI>"
		    		else
			    	    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI>" & server.htmlencode(rs("Path2") & "")  & "</LI>"
				    end if
				end if
				if trim(rs("Path3") & "") <> "" then
	    			if left(rs("Path3"),2) = "\\" then
    				    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI><a target=_blank href=""" & server.htmlencode(rs("Path3") & "") & """>" & server.htmlencode(rs("Path3") & "")  & "</a></LI>"
		    		else
			    	    Response.Write "<TR><TD>&nbsp;</TD><TD colspan=2><LI>" & server.htmlencode(rs("Path3") & "")  & "</LI>"
				    end if
				end if
				Response.Write "<BR><BR></td></TR>"
				rs.movenext	
			loop
			Response.Write "</TABLE>"
		end if
		
		rs.Close
		cn.Close
		set rs = nothing
		set cn = nothing
	end if  
	 %>
    <input style="display:none" id="TeamID" name="TeamID" type="text" value="<%=request("TeamID")%>">
    <input style="display:none" id="Reload" name="Reload" type="text" value="0">
</form>
<%end if%>
</BODY>
</HTML>
