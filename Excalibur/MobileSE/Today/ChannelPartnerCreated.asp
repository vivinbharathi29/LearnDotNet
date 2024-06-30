<script type="text/javascript">
    var IOIDclicked = 0;
    function displayIOProperty(IOID) 
    {
       // window.open("/ipulsar/Admin/areas/InstallOption_Edit.aspx?Mode=update&IOID=" + IOID, "_blank", "","Width=750,Height=650,menubar=no,toolbar=no,scrollbars=Yes,resizable=Yes,status=No");

        IOIDclicked = IOID;
        var QueryString = "/ipulsar/Admin/areas/InstallOption_Edit.aspx?Mode=update&IOID=" + IOID;
        $("#divOpenIOPopUp").dialog({ width: 800, height: 600, modal: true });
        $("#ifdivOpenIOPopUp").attr("width", "98%");
        $("#ifdivOpenIOPopUp").attr("height", "98%");
        $("#ifdivOpenIOPopUp").attr("src", QueryString);
        $("#divOpenIOPopUp").dialog("option", "title", "Install Option");
        $("#divOpenIOPopUp").dialog("open");

    }
    function CloseEditPopup(refresh) {
        if (refresh) {
            $('#TrChannelpartner' + IOIDclicked).closest("tr").remove();

        }
        ClosePopup("divOpenIOPopUp");
        return false;
    }
    
</script>
    
<%
    
 	set rs = server.CreateObject("ADODB.recordset")
	If strImpersonateID = "" Then
	    rs.Open "usp_Admin_ViewChannelPartnersWithEmptyIODMIString " & clng(CurrentUserID),cn, adOpenStatic
    Else
	    rs.Open "usp_Admin_ViewChannelPartnersWithEmptyIODMIString " & clng(strImpersonateID),cn, adOpenStatic
    End If
    
    
    
    
    if Not rs.EOF or Not rs.BOF then %>
    <form ID="frmChannelpartner_Pulsar" method="post">
    <table id="TableChannelpartner" border="0" cellspacing="0" cellpadding="2">		
     <thead>	
      <tr>
        <td colspan="7">
        <%if strImpersonateName <> "" then%>
             <font size="2"><strong><u><br>Channel Partner Created</u></strong></font>&nbsp;-&nbsp;<strong><font color=red><%=strImpersonateName%></strong></font>
        </td>
	    <%else%>
	         <font size="2"><strong><u><br>Channel Partner Created</u></strong></font>
        </td>
	    <%end if%>
      </tr>
  
    </thead>
    </table>

     <table id="Channelpartnercreated" border="0" width="100%" cellspacing="0" cellpadding="2">		
         <thead>
           <tr bgcolor="beige">
            <td style="width:300px;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Channel Partner Name</b></font></td>
	        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Created By </b></font></td>	
               <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Created </b></font></td>	
         </tr>
         </thead>

         <%
 	
	      do while not rs.EOF
	        
     %>
        <tr bgcolor="ivory" id='TrChannelpartner<%=rs("IOID")%>'   onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()">
            
		    <td  style="width:300px; BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="displayIOProperty(<%=rs("IOID")%>);">
                <font class="text" size="1"><%= rs("PartnerName")%> </font> 
		    </td>
            <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="displayIOProperty(<%=rs("IOID")%>);">
                <font class="text" size="1"><%= rs("CreatedBy") %></font> 
		    </td>	
            <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="displayIOProperty(<%=rs("IOID")%>);">
                <font class="text" size="1"><%= rs("TimeCreated") %></font> 
		    </td>	
        </tr>

	    <%
	    rs.MoveNext
	    loop


	    rs.Close
	    %>
    </table>
        <div id="divOpenIOPopUp" title="Coolbeans" style="display: none;">
        <iframe frameborder="0" name="ifdivOpenIOPopUp" id="ifdivOpenIOPopUp" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
    </form>
<%end if %>