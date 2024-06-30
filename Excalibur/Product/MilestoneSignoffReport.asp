<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
<TITLE>RTM Report</TITLE>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function window_onload(){

}


//-->
</SCRIPT>
</HEAD>
<style>
.EmbeddedTable TBODY TD
{
   	FONT-FAMILY: Verdana;
}
.EmbeddedTable TBODY TD
{
  	Font-Size: xx-small;
}
A:visited,A:Link
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}    
</style>

<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<%
    dim cn
    dim rs
    dim strproduct
    dim strTitle
    dim strDate
    dim strComments
    dim strBIOSComments
    dim strFWComments
    dim strPatchComments
    dim strRestoreComments
    dim strImageComments
    dim strSubmitter
    dim strCreated
    dim strID
    dim AlertRowArray
    dim strAttachment1
    dim PathParts
    dim rtpStatus
    dim rtpDate
    dim serialNumber
    dim rtpComments
    dim typeId
    dim emailRTMNotifications
    dim productId
    dim rtpCompletedby
    
    strID = request("ID")
    
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	rs.open "spGetProductRTM " & clng(strID),cn
	if rs.eof and rs.bof then
        strproduct = ""
        strTitle = ""
        strDate = ""
        strComments = ""
        strBIOSComments = ""
        strFWComments = ""
        strRestoreComments = ""
        strPatchComments = ""
        strImageComments = ""
        strSubmitter = ""
        strCreated = ""
        strAttachment1 = ""
        rtpStatus =  ""
        rtpDate =  ""
        serialNumber =  ""
        rtpComments  = ""
        typeId = ""
        emailRTMNotifications =""
        productId=""
        rtpCompletedby = ""
	else
        strProduct = rs("Product") 
        strTitle = rs("Title") & ""
        strDate = rs("RTMDate") & ""
        strComments = rs("Comments") & ""
        strPatchComments = rs("PatchComments") & ""
        strBIOSComments = rs("BIOSComments") & ""
        strFWComments = rs("FWComments") & ""
        strRestoreComments = rs("RestoreComments") & ""
        strImageComments = rs("ImageComments") & ""
        strSubmitter = rs("Submitter") & ""
        strCreated = rs("created") & ""
        strAttachment1 = rs("Attachment1") & ""
        PathParts = split(strAttachment1,"\")
        rtpStatus = rs("RTPStatus") & ""
        rtpDate = rs("RTPDate") & ""
        serialNumber =  rs("SerialNumber") & ""
        rtpComments  = rs("RTPComments") & ""
        typeId =  rs("typeid") & ""
        emailRTMNotifications =  rs("Email") & ""
        productId = rs("ID") & ""
        rtpCompletedby = rs("RTPCompletedby") & ""
	end if
	rs.close
	
	if strTitle = "" or strDate = "" then
	    response.Write "<font size=2 face=verdana>Unable to find the requested document.</font>"
	else
%>


<font size=3 face=verdana><b><%=strproduct & " - " & strTitle%></b></font><br><br>
<FONT size=2 face=verdana><B>General RTM Information</B></FONT>
<TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
    <TBODY>
        <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTM Title:</B></TD>
            <TD width="100%"><%=strTitle%>&nbsp;</TD>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTM Date:</B></TD>
            <TD width=120><%=strDate%>&nbsp;</TD>
        </TR>
         <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTP Status:</B></TD>
         <% if Not IsNull(rtpStatus) and not rtpStatus = "" then%>
                <%if cbool(rtpStatus) = true then%>
             <TD width="100%">Approved&nbsp;</TD>
               <%else %>
             <TD width="100%">Rejected&nbsp;</TD>
               <%end if%>
         <%else %>
             <TD width="100%">&nbsp;</TD>
         <%end if%>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTP Date:</B></TD>
          <% if Not IsNull(rtpStatus) and not rtpStatus = "" then%>
                <%if cbool(rtpStatus) = true then%>
              <TD width=120> <%=formatdatetime(rtpDate,vbshortdate)%>  &nbsp;</TD>
               <%else %>
             <TD width=120> <%=formatdatetime(rtpDate,vbshortdate)%>  &nbsp;</TD>
               <%end if%>
         <%else %>
             <TD width="100%">&nbsp;</TD>
         <%end if%>
        </TR>
        <%if trim(strComments) <> "" then %>
            <TR>
                <TD bgColor=gainsboro vAlign=top noWrap><B>RTM Comments:</B></TD>
                <TD colSpan=3><%=replace(strComments,vbcrlf,"<br>")%>&nbsp;</TD>
            </TR>
        <%end if%>
        <%if trim(strAttachment1) <> "" then %>
            <TR>
                <TD bgColor=gainsboro vAlign=top noWrap><B>SCMx File:</B></TD>
                <TD colSpan=3><a target=_blank href="file://<%=strAttachment1%>"><%=PathParts(ubound(PathParts))%></a></TD>
            </TR>
        <%end if%>
        <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTM Initiated By:</B></TD>
            <TD width="100%" colspan=3><%=longname(strSubmitter)%> on <%=formatdatetime(strCreated,vbshortdate)%>&nbsp;</TD>
        </TR>
         <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTP Completed By:</B></TD>
              <% if Not IsNull(rtpCompletedby) and not rtpCompletedby = "" then%>
              <TD width="100%" colspan=3><%=longname(strSubmitter)%> on <%=formatdatetime(rtpCompletedby,vbshortdate)%>&nbsp;</TD>
              <%else %>
             <TD width="100%" colspan=3><%=longname(strSubmitter)%>&nbsp;</TD>
              <%end if%>
           
        </TR>
        <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTP Serial Number:</B></TD>
            <TD width="100%" colspan=3><%=serialNumber%>&nbsp;</TD>
        </TR>
        <TR>
            <TD bgColor=gainsboro vAlign=top noWrap><B>RTP Comments:</B></TD>
            <TD width="100%" colspan=3><%=rtpComments%>&nbsp;</TD>
        </TR>
    </TBODY>
</TABLE>
<%
    rs.open "spListProductRTMDeliverables " & clng(strID) & ",1",cn
    if not (rs.eof and rs.bof) then
%>
        <BR>
        <FONT size=2 face=verdana><B>System BIOS to RTM</B></FONT><BR>
        <%if trim(strBIOSComments) <> "" then%>
            <font size=1 face=verdana><BR><i><%=replace(strBIOSComments,vbcrlf,"<br>")%></i><br><br></font>
        <%end if%>
        <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
        <TBODY>
        <TR bgColor=gainsboro>
        <TD><B>ID</B></TD>
        <TD><B>Name</B></TD>
        <TD><B>Version</B></TD>
        <TD><B>Affectivity</B></TD>
    <%end if%>
    <%
    do while not rs.eof
        strversion = rs("Version")
        if trim(rs("Revision") & "") <> "" then
            strversion = strVersion & "," & rs("Revision")
        end if
        if trim(rs("Pass") & "") <> "" then
            strversion = strVersion & "," & rs("Pass")
        end if
        response.write "<TR>"
        response.write "<TD>" & rs("ID") & "</TD>"
        response.write "<TD>" & rs("Name") & "</TD>"
        response.write "<TD>" & strVersion & "</TD>"
        response.write "<TD>" & rs("Details") & "&nbsp;</TD>"
        response.write "</TR>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
    %>
    </TBODY></TABLE>
<%
    end if
    rs.close

        rs.open "spListProductRTMDeliverables " & clng(strID) & ",4",cn
    if not (rs.eof and rs.bof) then
%>
        <BR>
        <FONT size=2 face=verdana><B>Firmware to RTM</B></FONT><BR>
        <%if trim(strFWComments) <> "" then%>
            <font size=1 face=verdana><BR><i><%=replace(strFWComments,vbcrlf,"<br>")%></i><br><br></font>
        <%end if%>
        <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
        <TBODY>
        <TR bgColor=gainsboro>
        <TD><B>ID</B></TD>
        <TD><B>Name</B></TD>
        <TD><B>Version</B></TD>
        <TD><B>Affectivity</B></TD>
    <%end if%>
    <%
    do while not rs.eof
        strversion = rs("Version")
        if trim(rs("Revision") & "") <> "" then
            strversion = strVersion & "," & rs("Revision")
        end if
        if trim(rs("Pass") & "") <> "" then
            strversion = strVersion & "," & rs("Pass")
        end if
        response.write "<TR>"
        response.write "<TD>" & rs("ID") & "</TD>"
        response.write "<TD>" & rs("Name") & "</TD>"
        response.write "<TD>" & strVersion & "</TD>"
        response.write "<TD>" & rs("Details") & "&nbsp;</TD>"
        response.write "</TR>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
    %>
    </TBODY></TABLE>
<%
    end if
    rs.close

    rs.open "spListProductRTMDeliverables " & clng(strID) & ",3",cn
    if not (rs.eof and rs.bof) then
%>
        <BR>
        <FONT size=2 face=verdana><B>Patches to RTM</B></FONT><BR>
        <%if trim(strPatchComments) <> "" then%>
            <font size=1 face=verdana><BR><i><%=replace(strPatchComments,vbcrlf,"<br>")%></i><br><br></font>
        <%end if%>
        <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
        <TBODY>
        <TR bgColor=gainsboro>
        <TD><B>ID</B></TD>
        <TD><B>Name</B></TD>
        <TD><B>Version</B></TD>
        <TD><B>Patch&nbsp;Contents</B></TD>
        <TD><B>Images</B></TD>
    <%end if%>
    <%
    do while not rs.eof
        strversion = rs("Version")
        if trim(rs("Revision") & "") <> "" then
            strversion = strVersion & "," & rs("Revision")
        end if
        if trim(rs("Pass") & "") <> "" then
            strversion = strVersion & "," & rs("Pass")
        end if
        response.write "<TR>"
        response.write "<TD>" & rs("ID") & "</TD>"
        response.write "<TD>" & rs("Name") & "</TD>"
        response.write "<TD>" & strVersion & "</TD>"
        
        strPatchContents = ""
        set rs2 = server.CreateObject("ADODB.recordset")
        rs2.open "spGetSelectedDepends " & clng(rs("ID")),cn
        do while not rs2.eof
            if strPatchContents <> "" then
                strPatchContents = strPatchContents & "<BR>" 
            end if
            strPatchContents = strPatchContents & rs2("Name") & " [" & rs2("Version")
            if trim(rs2("revision")&"") <> "" then
                strPatchContents = strPatchContents & "," & rs2("revision")
            end if
            if trim(rs2("pass")&"") <> "" then
                strPatchContents = strPatchContents & "," & rs2("pass")
            end if
            rs2.movenext
        loop
        rs2.close   
        set rs2 = nothing

        response.write "<TD>" & strPatchContents & "&nbsp;</TD>"
        response.write "<TD><a target=_blank href=""../Image/PatchImages.asp?ProdID=" & rs("Productversionid") & "&DelID=" & rs("id") & """>View</a>&nbsp;</TD>"
        response.write "</TR>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
    %>
    </TBODY></TABLE>
<%
    end if
    rs.close

    rs.open "spListProductRTMDeliverables " & clng(strID) & ",2",cn
    if not (rs.eof and rs.bof) then
%>
        <BR>
        <FONT size=2 face=verdana><B>Restore Media to RTM</B></FONT><BR>
        <%if trim(strRestoreComments) <> "" then%>
            <font size=1 face=verdana><BR><i><%=replace(strRestoreComments,vbcrlf,"<br>")%></i><br><br></font>
        <%end if%>
        <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
        <TBODY>
        <TR bgColor=gainsboro>
        <TD><B>ID</B></TD>
        <TD><B>Name</B></TD>
        <TD><B>Version</B></TD>
        <TD><B>Part</B></TD>
        <TD><B>PMR&nbsp;Date</B></TD></TR>
    <%end if%>
    <%
    do while not rs.eof
        strversion = rs("Version")
        if trim(rs("Revision") & "") <> "" then
            strversion = strVersion & "," & rs("Revision")
        end if
        if trim(rs("Pass") & "") <> "" then
            strversion = strVersion & "," & rs("Pass")
        end if
        response.write "<TR>"
        response.write "<TD>" & rs("ID") & "</TD>"
        response.write "<TD>" & rs("Name") & "</TD>"
        response.write "<TD>" & strVersion & "</TD>"
        response.write "<TD>" & rs("CDPartNumber") & "&nbsp;</TD>"
        response.write "<TD>" & rs("PMRDate") & "&nbsp;</TD>"
        response.write "</TR>"
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
    %>
    </TBODY></TABLE>
<%
    end if
    rs.close

    
    rs.open "splistImages4RTM " & clng(strID),cn
    if not (rs.eof and rs.bof) then
%>
    <BR><FONT size=2 face=verdana><B>Images to RTM</B></FONT><BR>
        <%if trim(strImageComments) <> "" then%>
            <font size=1 face=verdana><i><%=replace(strImageComments,vbcrlf,"<br>")%></i></font>
            <br>
        <%end if%>
    <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
    <TBODY>
        <TR bgColor=gainsboro>
            <TD><B>SKU</B></TD>
            <TD><B>Region</B></TD>
            <TD><B>Model</B></TD>
            <TD><B>OS</B></TD>
            <TD><B>Apps Bundle</B></TD>
            <TD><B>BTO/CTO</B></TD>
        </TR>
<%
    do while not rs.eof
%>    
        <TR>
            <TD><%=rs("SKUNumber")%>&nbsp;</TD>
            <TD><%=rs("Region") %>&nbsp;</TD>
            <TD><%=rs("Model")%>&nbsp;</TD>
            <TD><%=rs("OS")%>&nbsp;</TD>
            <TD><%=rs("Apps")%>&nbsp;</TD>
            <TD><%=rs("ImageType")%>&nbsp;</TD>
        </TR>
<%
        rs.movenext
    loop
%>        
    </TBODY>
    </TABLE>
    
<%
    else

        if trim(strImageComments) <> "" then%>
        <BR><FONT size=2 face=verdana><B>SCMX to RTM</B></FONT><BR>
            <font size=1 face=verdana><i><%=replace(strImageComments,vbcrlf,"<br>")%></i></font>
            <br>
        <%end if

    end if
    rs.close    

    rs.open "splistProductRTMAlerts " & clng(strID),cn
    if not (rs.eof and rs.bof) then
%>
        <BR><FONT size=2 face=verdana><B>Alerts Reviewed</B></FONT><BR>
        <TABLE class=EmbeddedTable border=1 cellSpacing=0 cellPadding=2 width="100%" bgColor=white>
            <TBODY>
            <TR bgColor=gainsboro>
                <TD><B>Alert</B></TD>
                <TD><B>Count</B></TD>
                <TD width="100%"><B>Comments</B></TD>
            </TR>
            <%
            do while not rs.eof
                AlertRowArray = split(lcase(rs("alertHTML") & ""),"</tr>")
            %>
            <TR>
                <TD noWrap><%=rs("Name")%>&nbsp;</TD>
                <TD noWrap align=center><A href="MilestoneSignoffAlertPreview.asp?ID=<%=rs("ID")%>" target=_blank><%=ubound(AlertRowArray)-1%></A></TD>
                <TD><%=rs("Comments")%>&nbsp;</TD>
            </TR>
            <%
                rs.movenext
            loop
            %>


        </TBODY></TABLE>

<%
    
    end if
    rs.close    
end if
	set rs = nothing
	cn.close
	set cn = nothing
	
	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
		
	end function	

%>
</BODY>
    <input type="hidden" id="email" name="email" value="<%=request("SendMail")%>">
    
</HTML>
<form id="frmEmail" method="post" action="UpdateRTPStatusEmail.asp">
    <input type="hidden" id="txtEmailBody" name="txtEmailBody" value="">
    <input type="hidden" id="txtProductName" name="emailBody" value="<%=strproduct%>">
    <input type="hidden" id="txtRTMName" name="emailBody" value="<%=strTitle%>">
    <input type="hidden" id="txtEmailRTMNotifications" name="emailBody" value="<%=emailRTMNotifications%>">
    <input type="hidden" id="txtProductID" name="txtProductID" value="<%=productId%>">
</form>
<script>
    window.onload=sendEmailCallBack; 
    function sendEmailCallBack() {
        var innerhtml = document.getElementsByTagName("body")[0].innerHTML;
        document.getElementById("txtEmailBody").value = "<html>"+ innerhtml + "</html>";
        if (email.value == "true") {
           frmEmail.submit();
        }
    }
</script>


