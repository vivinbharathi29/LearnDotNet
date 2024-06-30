<script type="text/javascript">
    function UpdateGPGSAName(ActionItemID) {
        var strID;
      
       
        var url = "/ipulsar/Product/EditBaseUnitMissingSA.aspx?ActionItemID=" + ActionItemID;
        
        ShowPopupUpdateGPGSAName(url, "Update GPG (40c SA) ", 800, 500);

    }

    function ShowPopupUpdateGPGSAName(QueryString, Title, DlgWidth, DlgHeight) {
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        $("#iframeDialog").dialog({
            width: DlgWidth,
            height: DlgHeight,
            modal: true
        });
        $("#modalDialog").attr("width", "98%");
        $("#modalDialog").attr("height", "98%");
        $("#modalDialog").attr("src", QueryString);
        $("#iframeDialog").dialog("option", "title", Title);
        $("#iframeDialog").dialog("open");
    }

    function CloseUpdateGPGSAName(Refresh,ActionItemID, GPG_40c_SA) {
        $("#modalDialog").attr("src", "");
        $("#iframeDialog").dialog("close");
        if (Refresh == true) {
            $('#BaseUnitGPGSA' + ActionItemID).html(GPG_40c_SA);
            $('#BaseUnitGPGSA' + ActionItemID).css("color", "black");
            var checkBoxes = document.getElementsByTagName("input");
            for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].value == ActionItemID) {
               checkBoxes[i].disabled=false;
            }
        }
        }
    }
    function btnExportRPN_onclick() {
                                  
             
        var elemsIDChecked ="";
        var strCheckedIDs ="";  

        if ($(":checkbox[name='chkBUMissingSAItems']").is(':checked')) {            
            $('input:checkbox[name="chkBUMissingSAItems"]:checked').each(function () {
                elemsIDChecked = $(this).val();
                strCheckedIDs = strCheckedIDs == "" ? elemsIDChecked : strCheckedIDs + "," + elemsIDChecked;
                
            });            
        }
        else {
            alert("Please select base unit features to export RPN");
            return;
        }

        var url = "/ipulsar/Product/BaseUnitMissingSA_ExportRPN.aspx?ActionItemIDs=" + strCheckedIDs;

        ShowPopupUpdateGPGSAName(url, " Exporting RPN ", 250, 80);


        // code to close popup after report is generated 
        var millisecondsToWait = 2000; //Run every 2 seconds to check for cookie
        var intrvl = setInterval(function () {
            if (getCookie("ExportRPNfromTodayPage") != "") // if cookie exists
            {
                document.cookie = "ExportRPNfromTodayPage=; expires=Thu, 01 Jan 1970 00:00:00 UTC" + "; path=/"; // delete cookie
                $("#modalDialog").attr("src", "");
                $("#iframeDialog").dialog("close"); // close popup        
                clearInterval(intrvl); //Clear timer 
            }
        }, millisecondsToWait);


        //hide the rows exported:
        var i = 0;
        var t = document.getElementById('BaseUnitMissingSATable');
        var rowid = "";
        var GPGSA = "";

        
         $('input:checkbox[name="chkBUMissingSAItems"]:checked').each(function () {
                elemsIDChecked = $(this).val();
              
                   $('#BaseUnitMissingSARow' + elemsIDChecked).closest("tr").remove();
            });            


    }

    function getCookie(cname) {
        var name = cname + "=";
        var ca = document.cookie.split(';');
        for (var i = 0; i < ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) == ' ') {
                c = c.substring(1);
            }
            if (c.indexOf(name) == 0) {
                return c.substring(name.length, c.length);
            }
        }
        return "";
    }
    
    function chkBUMissingSAAll_onclick()
    {
        var i;
        var checkBoxes = document.getElementsByTagName("input");
        var chkCreateAVsAll, chkBoxName;
        chkRejectedAvsItemsAll = document.getElementById("chkBUMissingSAAll");
        chkBoxName = "chkBUMissingSAItems";
        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].name == chkBoxName &&  !checkBoxes[i].disabled) {
                checkBoxes[i].checked = chkRejectedAvsItemsAll.checked;
            }
        }
      
    }
</script>
<%  
      
  		rs.Open "usp_Product_ViewBaseUnitMissingSAActionItems " & Currentuserid ,cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			
		    if 	blnEngCoordinatorHeaderDisplayed = false and intEngCoordinator > 0 then 
			%>
				<table cellSpacing="1" cellPadding="1" width="100%" bgcolor="<%=strBGColor%>" border="0">
				  
				<tr>
				 <td><strong><font face="verdana" color="<%=strForeColor%>" size="2">Engineering Coordinator Alerts</font>  
				   </strong></td></tr></table><br>
			<%
				blnEngCoordinatorHeaderDisplayed = true
			end if
	  %>
	
	        <table ID="BaseUnitMissingSATable" cellspacing="0" border="0" width="100%">
	        <thead>
	        <tr>
		        <td colspan="3">
		        <p>
		        <font size="2" face="verdana"><strong><u>Missing Base Unit Subassembly Numbers (Pulsar Product)</u></strong><br></font></p></td>
		        <td><font size="2"></font></td></tr>
             <tr>
		        <td colspan="3">
		            <input type="button" value="Export RPN" id="btnExportRPN" name="btnExportRPN" class="button2" style="width:90px" onclick="return btnExportRPN_onclick()" >

		        </td></tr>
                
	        <tr bgcolor="#dadcbc">
                <td valign="top" style="width:30px;" >
                    <input id="chkBUMissingSAAll" type="checkbox" style="height:16px;width:16px" onclick="javascript: chkBUMissingSAAll_onclick();">
                </td>
	            <td onclick="SortTable( 'BaseUnitMissingSATable', 0 ,1,2);"><font face="verdana" size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Base Unit Feature ID</b></font></td>
	            <td onclick="SortTable( 'BaseUnitMissingSATable', 1 ,0,2);" nowrap><font face="verdana" size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product</font></td>
	            <td onclick="SortTable( 'BaseUnitMissingSATable', 2 ,0,2);"><font face="verdana" size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">GPG (40c SA)</b></font></td>
	         </tr>
	        </thead>
	<%
		    do while not rs.EOF
			    if trim(rs("ProductName") & "") = "" then
				     strProduct = "&nbsp;"
			    else
				    strProduct = rs("ProductName") 
			    end if
			    if rs("id") = 0 then
				    strSubID = "Feature" & trim(rs("id"))
			    else
				    strSubID = rs("id")
			    end if
			
			%>

		        <tr ID="BaseUnitMissingSARow<%=strSubID%>" bgcolor="ivory"  onmouseover="return Commodity_onmouseover()" onmouseout="return Commodity_onmouseout()">
		            <td valign="top" style="width:30px;BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL ">
                        <input id="chkBUMissingSAItems" name="chkBUMissingSAItems" type="checkbox" style="height:16px;width:16px" value="<%=rs("id")%>"  <% if len(rs("FeatureName")) > 40 then response.Write " disabled"%>>
                    </td>      
                    <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap  onclick="return UpdateGPGSAName(<%=rs("id")%>)"><font face="verdana" class="text" size="1"><%=rs("FeatureID")%>&nbsp;</font></td>
		            <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap  onclick="return UpdateGPGSAName(<%=rs("id")%>)"><font face="verdana" class="text" size="1"><%=strProduct%></font></td>
		            <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap  onclick="return UpdateGPGSAName(<%=rs("id")%>)"><font face="verdana" class="text" size="1"  <% if len(rs("FeatureName")) > 40 then response.Write " color='red'"%> id="BaseUnitGPGSA<%=rs("id")%>"><%=rs("FeatureName")%>&nbsp;</font></td>
                </tr>
		<%	
			rs.MoveNext
		loop
	
	%>
	</table><br>
	<%
		end if 'close the if not eof if statement
		rs.Close
	
%>
