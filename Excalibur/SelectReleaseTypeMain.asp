<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
    td
    {
    FONT-Family: verdana;
    FONT-Size: x-small;	
    }
    body
    {
    FONT-Family: verdana;
    FONT-Size: x-small;	
    }
</STYLE>
<BODY bgcolor=ivory>
<%
    dim strDisplayedID
    dim strVersion

    strDisplayedID = ""
    strVersion = ""
    
	if trim(request("ID")) = "" or not isnumeric(request("ID")) then
		Response.Write "Unable to find the requested record."
	else
		dim cn
		dim rs
		dim strCertificationStatus
        
        strCertificationStatus = ""

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
        
        rs.open "spGetDeliverableVersionProperties " & clng(request("ID")),cn
        if not (rs.eof and rs.bof) then
            strDisplayedID = rs("ID") & ""
            strVersion = rs("Version") & ""
            if trim(rs("Revision") & "") <> "" then
                strVersion = strVersion & "," & trim(rs("Revision") & "")
            end if
            if trim(rs("Pass") & "") <> "" then
                strVersion = strVersion & "," & trim(rs("Pass") & "")
            end if
            strCertificationStatus = trim(rs("CertificationStatus") & "")
        end if
        rs.close
   end if
   
   if strDisplayedID = "" then
		Response.Write "Unable to find the requested deliverable."
   else
%>
    <font size=3 face=verdana><b>Add New Version</b><br><br></font>
    <b>Choose Release Type</b><br>
    
    <table width="100%" border=1 cellspacing=0 cellpadding=1 bordercolor=tan bgcolor=cornsilk>
    <tr>
        <td>
         <table width="100%">
        <tr><td valign=top>
        <input checked id="optType1" name="optReleaseType" type="radio"></td>
        <td>
         Normal Release
        <ul style="margin-top:0px;margin-bottom:0px">
            <li>Most fields are automatically populated from the root deliverable.</li>
        </ul>
        </td></tr>
        </table>
        <!--- New deliverable files can be attached
        <br>-->
        
        </td>
    </tr>
<%if trim(strCertificationStatus) = "1" or trim(strCertificationStatus) = "2" or trim(strCertificationStatus) = "4" then%>
    <tr>
<%else%>
    <tr style="display:none">
<%end if%>
        <td>
         <table width="100%">
        <tr><td valign=top>
        <input id="optType2" name="optReleaseType" type="radio"></td>
        <td>
         WHQL Bit-Flip
            <ul style="margin-top:0px;margin-bottom:0px">
                <li>Most fields are automatically populated from version <%=strVersion%>.</li>
            </ul>
        </td></tr>
        </table>
        <!--- New deliverable files can be attached
        <br>-->
        
        </td>
    </tr>
    <tr style="display:none">
        <td valign=top>
         <table width="100%">
        <tr><td valign=top>
        <input id="optType3" name="optReleaseType" type="radio"></td>
         <td>
         Update Preinstall CVA Files (System ID changes only)<br>
         <ul style="margin-top:0px;margin-bottom:0px">
        <li>The Release Team will copy new CVA files to version <%=strVersion%>.</li>
        </ul>
        <br><br><font color=green>NOTE: The Release Team will not accept CVA files with any changes other than System ID changes.</font>
        </td></tr></table>
        </td>
    </tr>
    <tr>
        <td valign=top>
         <table width="100%">
        <tr><td valign=top>
        <input id="optType4" name="optReleaseType" type="radio">
        </td>
        <td>
        Update Preinstall CVA File<br>
<!--        - This option should not be used if only the system ID list is changing.<br>
        - Release Emails and Workflow Today Page sections will indicate that only the CVA files have changed in this release.<br>
        - Only the CVA files can be updated.<br>  
        -->
        <ul style="margin-top:0px;margin-bottom:0px">
            <li>Most fields are automatically populated from version <%=strVersion%>.</li>
            <li>You only need to supply a new CVA File.  The Release Team will use that file and a copy of version <%=strVersion%> to create the new pass of this deliverable.<br /><font color=red>This option only works for releasing new CVA files.<br /><br /><b>ANY CHANGES TO FILES OTHER THAN THE CVA FILE WILL BE IGNORED!!!</b></font></li>
        </ul>
        </td></tr></table>
        </td>
    </tr>
    </table>
    
<%
    end if
 %>
</BODY>
</HTML>
