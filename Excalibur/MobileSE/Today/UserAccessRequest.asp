<script type="text/javascript">

    //*****************************************************************
    //Function:     ShowApprovalPage();
    //Description:  Approval page
    //Created:      (4/29/1016) - PBI 7833/Task 19948
    //*****************************************************************
    function ShowApprovalPage() {
        var sQueryString = "/ipulsar/Admin/System Admin/UsersAndRoles_Requests.aspx";
        /*OpenUsersRolesRequests(sQueryString);*/
        window.showModalDialog(sQueryString, "", " dialogWidth:500px; dialogHeight:500px; center:Yes; help:No; maximize:no; resizable:no; status:No");
        
    }
</script>
    
<form ID="frmPhWebActionItems_Pulsar" method="post">
<table id="TablePHwEbPulsar" border="0" cellspacing="0" cellpadding="2">		
 <thead>	
  <tr>
    <td colspan="7">
    <%if strImpersonateName <> "" then%>
         <font size="2"><strong><u><br>Users and Roles Requests</u></strong></font>&nbsp;-&nbsp;<strong><font color=red><%=strImpersonateName%></strong></font>
    </td>
	<%else%>
	     <font size="2"><strong><u><br>Users and Roles Requests</u></strong></font>
    </td>
	<%end if%>
  </tr>
  <tr>
    <td nowrap align=left>
    <% 
      if strImpersonateName <> "" then
	      btnUpdate5 = "<input type=""button"" value=""Approval Page"" id=""btnAutoInput"" name=""btnAutoInput"" class=""button2"" style=""width:100px"" onclick=""return ShowApprovalPage()"">"
	      Response.Write(btnUpdate5)
    	
	  else
	      btnUpdate5 = "<input type=""button"" value=""Approval Page"" id=""btnAutoInput"" name=""btnAutoInput"" class=""button2"" style=""width:100px"" onclick=""return ShowApprovalPage()"">"
	      Response.Write(btnUpdate5)
    	
	  end if
	%>
	</td>
  </tr>
</thead>
</table>

 <table id="PhWebAvActionItemsPulsar" border="0" width="100%" cellspacing="0" cellpadding="2">		
     <thead>
     <tr bgcolor="beige">
        <td  style="width: 400px;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Team Name</b></font></td>
	    <td  style="width: 150px;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">User Name</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Date Requested</b></font></td>	
     </tr>
     </thead>

     <%
 	set rs = server.CreateObject("ADODB.recordset")
	If strImpersonateID = "" Then
	    rs.Open "usp_ADMIN_UserAccessRequest_TodayPage " & clng(CurrentUserID),cn, adOpenStatic
    Else
	    rs.Open "usp_ADMIN_UserAccessRequest_TodayPage " & clng(strImpersonateID),cn, adOpenStatic

    End If
     
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD><font size=1 face=verdana><b>none</b></font></TD></TR>"
    else
         do while not rs.EOF    
    %>

    <tr bgcolor="ivory" id='Tr1'">
		<td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("Name")%></font> 
		</td>
        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("CreatedBy")%></font> 
		</td>
        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("Created")%></font> 
		</td>		
    </tr>

	<%
	rs.MoveNext
	loop

	end if

	rs.Close
	%>
</table>
</form>
