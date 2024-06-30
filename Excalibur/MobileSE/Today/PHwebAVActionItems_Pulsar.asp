<script type="text/javascript">
    function chkPHwebActionItemAll_onclick() {
        var i;
        var checkBoxes = document.getElementsByTagName("input");
        var chkCreateAVsAll, chkBoxName;
        chkPHwebActionItemAll = document.getElementById("chkPHwebActionItemsPulsarAll");
        chkBoxName = "chkPHwebActionItemsPulsar";
     
               
        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].name == chkBoxName) {
                checkBoxes[i].checked = chkPHwebActionItemAll.checked;
            }
        }
    }

    
    function ShowLifecycleDataManagement(UserID) {

        var i;
        var elemChecked;
        var elemsIDsChecked;

        var checkBoxes = document.getElementsByTagName("input");
        var chkCreateAVsAll, chkBoxName;

        chkBoxName = "chkPHwebActionItemsPulsar";

        elemChecked = false;
        elemsIDsChecked = "";


        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].name == chkBoxName) {
                if (checkBoxes[i].checked == true) {
                    elemChecked = true;
        
                    if (elemsIDsChecked == "")
                       elemsIDsChecked=checkBoxes[i].value;
                    else
                       elemsIDsChecked = elemsIDsChecked + ',' + checkBoxes[i].value;
                     
                }
            }
        }


        if (elemChecked == false)     // No elements checked
        {
            window.open("/ipulsar/SCM/LifeCycleDataMgmt.aspx", "_blank", "", "Width=750,Height=650,menubar=no,toolbar=no,scrollbars=Yes,resizable=Yes,status=No");
        }
        else
        {
            // elems checked, we pass the ID's for all the elements ID's checked 
            window.open("/ipulsar/SCM/LifeCycleDataMgmt.aspx?ProductBrandID=" + elemsIDsChecked, "_blank", "", "Width=750,Height=650,menubar=no,toolbar=no,scrollbars=Yes,resizable=Yes,status=No");
        }
    }

        function ShowLifecycleDataManagementAVaction(ProductBrandID) 
        {
        window.open("/ipulsar/SCM/LifeCycleDataMgmt.aspx?ProductBrandID=" + ProductBrandID, "_blank", "","Width=750,Height=650,menubar=no,toolbar=no,scrollbars=Yes,resizable=Yes,status=No");
    }

    
</script>
    
<%if intPDMUser > 0 then%>
<form ID="frmPhWebActionItems_Pulsar" method="post">
<table id="TablePHwEbPulsar" border="0" cellspacing="0" cellpadding="2">		
 <thead>	
  <tr>
    <td colspan="7">
    <%if strImpersonateName <> "" then%>
         <font size="2"><strong><u><br>PHweb AV Action Items (Pulsar Product)</u></strong></font>&nbsp;-&nbsp;<strong><font color=red><%=strImpersonateName%></strong></font>
    </td>
	<%else%>
	     <font size="2"><strong><u><br>PHweb AV Action Items (Pulsar Product)</u></strong></font>
    </td>
	<%end if%>
  </tr>
  <tr>
    <td nowrap align=left>
    <% 
      if strImpersonateName <> "" then
	      btnUpdate5 = "<input type=""button"" value=""Manage Lifecycle Data"" id=""btnAutoInput"" name=""btnAutoInput"" class=""button2"" style=""width:150px"" onclick=""return ShowLifecycleDataManagement(" & "'" & clng(strImpersonateID) & "'" & ")"">"
	      Response.Write(btnUpdate5)
    	
	  else
	      btnUpdate5 = "<input type=""button"" value=""Manage Lifecycle Data"" id=""btnAutoInput"" name=""btnAutoInput"" class=""button2"" style=""width:150px"" onclick=""return ShowLifecycleDataManagement(" & "'" & clng(CurrentUserID) & "'" & ")"">"
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
        <td valign="top" style="width:30px;" >
            <input id="chkPHwebActionItemsPulsarAll" type="checkbox" style="height:16px;width:16px" onclick="javascript: chkPHwebActionItemAll_onclick();">
        </td>
        <td style="width:300px;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product Name</b></font></td>
	    <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Days Until: </b></font></td>	
     </tr>
     </thead>

     <%
 	set rs = server.CreateObject("ADODB.recordset")
	If strImpersonateID = "" Then
	    rs.Open "usp_SelectPhWebAvActionItems2_Pulsar " & clng(CurrentUserID),cn, adOpenStatic
    Else
	    rs.Open "usp_SelectPhWebAvActionItems2_Pulsar " & clng(strImpersonateID),cn, adOpenStatic
    End If
     
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD><font size=1 face=verdana><b>none</b></font></TD></TR>"
	else	 
      BID = ""
	  do while not rs.EOF
	    if BID <> trim(rs("BID")) then
	        set rs2 = server.CreateObject("ADODB.recordset")
	        rs2.Open "usp_SelectPhWebDaysUntil " & rs("PVID") & "," & rs("BID"),cn, adOpenStatic
	        if not (rs2.EOF and rs2.BOF) then
	            DaysUntil = rs2("DaysUntil")
	        end if
	        if DaysUntil & "" = "" then
	            DaysUntil = "N/A"
	        end if 
	        rs2.Close 
        BID = trim(rs("BID"))
       end if
 %>
    <tr bgcolor="ivory" id='Tr1'   onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()">
        <td valign="top" style="width:30px; BORDER-TOP: <%=strRowBorderColor%> thin solid">
            <input id="chkPHwebActionItemsPulsar" name="chkPHwebActionItemsPulsar" type="checkbox" style="height:16px;width:16px" value="<%=rs("BID")%>">
        </td>  
		<td  style="width:300px; BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="ShowLifecycleDataManagementAVaction(<%=rs("BID")%>);">
            <font class="text" size="1"><%= rs("ProductName")%>&nbsp;/&nbsp;<%=rs("BrandName")%> </font> 
		</td>
        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="ShowLifecycleDataManagementAVaction(<%=rs("BID")%>);">
            <font class="text" size="1"><%= DaysUntil%></font> 
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
<%end if %>