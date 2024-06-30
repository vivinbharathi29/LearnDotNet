<%@ Language=VBScript %>

<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  

  Dim AppRoot
  AppRoot = Session("ApplicationRoot")
%>
	
<HTML>
<HEAD>
    <TITLE>Cycle Disclaimer Notes</TITLE>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

        function cmdOK_onclick() {
            if (!$("#txtDisclaimerNotes") || $("#txtDisclaimerNotes").val().length > 4000)
            {
                alert("Disclaimer Notes must be less than 4000 characters");
                return;
            }

            if ($("#txtDisclaimerNotes").val() === "" && $("#optActiveNotes").is(":checked")) {
                alert("State must be Inactive if no Disclaimer Notes provided");
                return;
            }

            //var ID = 0;
            //if ($("#txtID").val() != "") {
            //    ID = $("#txtID").val();
            //}
            //var ReleaseID = $("#txtReleaseID").val();
            //var Notes = $("#txtDisclaimerNotes").val();
            //var State = 0;
            //if ($("#optActiveNotes").is(":checked")) {
            //    State = 1;
            //}
            //var User = $("#txtUser").val();

            //document.forms["frmUpdate"].action = "EditDisclaimerNotesSave.asp"; //?" + "ID=" + ID + "&ReleaseID=" + ReleaseID + "&DisclaimerNotes=" + Notes + "&State=" + State + "&User=" + User;
            document.forms["frmUpdate"].submit();
        }

    </SCRIPT>
</HEAD>

<BODY bgcolor=ivory>
<%
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentDomain
    dim CurrentUserFullName
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(CurrentUser,"\") > 0 then
		CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
		CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	
	set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = CurrentUser
	cm.Parameters.Append p

	set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	set rs = cm.Execute 

	set cm=nothing	
	
    CurrentUserFullName = rs("Name")

    ' check for prl.edit?

	rs.Close

    dim BusinessSegment
    dim ReleaseName
    dim DisclaimerNotes
    dim Created
    dim CreatedBy
    dim Updated
    dim UpdatedBy
    		
    set cm = server.CreateObject("ADODB.Command")
	set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "usp_ProductVersion_GetDisclaimerNotes"
		
	set p = cm.CreateParameter("@ReleaseID", 3, &H0001)
	p.Value = request("ReleaseID")
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	set rs = cm.Execute 
    
	if not (rs.EOF and rs.BOF) then
        BusinessSegment = rs("BusinessSegment") + ""
        ReleaseName = rs("ReleaseName") + ""
        DisclaimerNotes = rs("DisclaimerNotes") + ""
        Created = rs("Created")
        CreatedBy = rs("CreatedBy") + ""
        Updated = rs("Updated")
        UpdatedBy = rs("UpdatedBy") + ""
	end if
%>


<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b>Cycle Disclaimer Notes</b></font>

<form  id="frmUpdate" name="frmUpdate" method="post" action="EditDisclaimerNotesSave.asp">    
    <table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	    <tr>
		    <td width="150" nowrap valign=top><b>Business Segment:</b>&nbsp;</td>
            <td width="100%"><label id="lblBusinessSegment"><%=BusinessSegment%></label></td>
	    </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Release:</b>&nbsp;</td>
            <td width="100%"><label id="lblReleaseName"><%=ReleaseName%></label></td>
	    </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Disclaimer Notes:</b><font color="red" size="1">&nbsp;*</font></td>
            <td colspan="3">
                <textarea id="txtDisclaimerNotes" name="txtDisclaimerNotes" style="width: 500px; height: 120px;" language="javascript"><%=DisclaimerNotes%></textarea>
                <div style="margin-top: .5em">
                    <span class="p-Instruction">&nbsp;4000 maximum characters</span>
                </div>
            </td>
        </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>State:</b><font color="red" size="1">&nbsp;*</font></td>
            <td>
                <% if IsNull(rs("Id")) or Not IsNull(rs("Disabled")) then %> 
                    <input id="optActiveNotes" name="optState" type="radio" value="1" /><font face="verdana" size="2">Active</font>
                    <input id="optInactiveNotes" name="optState" type="radio" checked="true" value="0" /><font face="verdana" size="2">Inactive</font>
                <% else %>
                    <input id="optActiveNotes" name="optState" type="radio" checked="true" value="1" /><font face="verdana" size="2">Active</font>
                    <input id="optInactiveNotes" name="optState" type="radio" value="0" /><font face="verdana" size="2">Inactive</font>
                <% end if  %>
            </td>
        </tr>

        <% if Created <> "" then %>
	        <tr>
		        <td width="150" nowrap valign=top><b>CreatedBy:</b>&nbsp;</td>
                <td width="100%"><label id="lblCreatedBy"><%=CreatedBy%></label></td>
	        </tr>
            <tr>
		        <td width="150" nowrap valign=top><b>Created:</b>&nbsp;</td>
                <td width="100%"><label id="lblCreated"><%=Created%></label></td>
	        </tr>
	        <tr>
		        <td width="150" nowrap valign=top><b>UpdatedBy:</b>&nbsp;</td>
                <td width="100%"><label id="lblUpdatedBy"><%=UpdatedBy%></label></td>
	        </tr>
	        <tr>
		        <td width="150" nowrap valign=top><b>Updated:</b>&nbsp;</td>
                <td width="100%"><label id="lblUpdated"><%=Updated%></label></td>
	        </tr>
        <% end if %>
    </table>

    <input type="hidden" id=txtID name=txtID value="<%=rs("Id")%>">
    <input type="hidden" id=txtReleaseID name=txtReleaseID value="<%=rs("ReleaseId")%>">
    <input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserFullName%>">
</form>

<%
    'rs.Close
	set rs = nothing
	set cn = nothing
%>

</BODY>
</HTML>
