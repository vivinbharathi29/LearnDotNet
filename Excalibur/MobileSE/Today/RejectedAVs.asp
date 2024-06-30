<script type="text/javascript">
    function ShowRejectedAVDetail(ProductVersionID, AvDetailID, ProductBrandID,userID)
    {
        var objReturn;      
        ShowPropertiesDialog("<%=AppRoot %>/SupplyChain/avFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "&UserID=" + userID + "&FromTodayPage=1", "SCM AV Details", 980, 800);
     
    }
    function ClosePopUpViewFromAvDetail_Features(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform) {
        if (document.getElementById('modalDialog').contentWindow != null)
            document.getElementById('modalDialog').contentWindow.ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform);
        $("#divOpenFeaturePopUp").dialog("close");
    }
    function ShowFeatureSelectDialog(QueryString, Title, DlgWidth, DlgHeight) {
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        QueryString = "../" + QueryString;
        $("#divOpenFeaturePopUp").dialog({ width: DlgWidth, height: DlgHeight, modal: true });
        $("#ifOpenFeaturePopUp").attr("width", "98%");
        $("#ifOpenFeaturePopUp").attr("height", "98%");
        $("#ifOpenFeaturePopUp").attr("src", QueryString);
        $("#divOpenFeaturePopUp").dialog("option", "title", Title);
        $("#divOpenFeaturePopUp").dialog("open");
    }    
    function ReloadAVData(avid, AvNo, GpgDescription, MarketingDesc, MarketingDescPMG, ConfigRules, RulesSyntax, TextAvId, Group1, Group2, Group3, Group4, Group5, Group6, Group7, Ids_Skus, Ids_Cto, Rcto_Skus, Rcto_Cto, Weight, GSEndDt, ProductLine, PBID, RTP, PAAD, SA, GA, EOM, BSAMB, Releases) {
       document.getElementById("tdMarketingDesc"+PBID + avid).innerHTML = MarketingDescPMG;
    }
    function ReloadAVDataFromMkt(AvID, GPGDescription, MarketingDesc, RTPDate, RASDisDate, PAADDate, SADate, GADate) {
        //document.getElementById("tdMarketingDesc"+PBID + avid).innerHTML = MarketingDescPMG;
    }
    
    function cmdRejectedAvsComplete_onclick(CurrentUser) {
        
        var i;       
        var elemsIDChecked;

        elemsIDChecked = "";      

        if ($(":checkbox[name='chkRejectedAVsItems']").is(':checked')) {            
            $('input:checkbox[name="chkRejectedAVsItems"]:checked').each(function () {
                elemsIDChecked = $(this).val();
                //strCheckedIDs = strCheckedIDs == "" ? elemsIDChecked : strCheckedIDs + "," + elemsIDChecked;
                var parameters = "AVPHwebRejectionItemID=" + elemsIDChecked + "&CurrentUserName=" + CurrentUser;
                var request = null;
                //Initialize the AJAX variable.
                if (window.XMLHttpRequest) {// Are we working with mozilla
                    request = new XMLHttpRequest(); //Yes -- this is mozilla.
                } else { //Not Mozilla, must be IE
                    request = new ActiveXObject("Microsoft.XMLHTTP");
                } //End setup Ajax.
                request.open("POST", "RejectedAvsActions.asp", false);
                request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                request.send(parameters);
                if (request.readyState == 4) {
                    $(this).prop("checked", false);
                    $('#RejectedAVsPulsarTr' + elemsIDChecked).closest("tr").remove();
                }
            });            
        }
        else {
            alert("Please select AVs to set to complete");
        }
    }

    function chkRejectedAvsAll_onclick()
    {
        var i;
        var checkBoxes = document.getElementsByTagName("input");
        var chkCreateAVsAll, chkBoxName;
        chkRejectedAvsItemsAll = document.getElementById("chkRejectedAVsAll");
        chkBoxName = "chkRejectedAVsItems";


        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].name == chkBoxName) {
                checkBoxes[i].checked = chkRejectedAvsItemsAll.checked;
            }
        }
    }

</script>

<table id="TableRejectedAvs" border="0" cellspacing="0" cellpadding="2">		
    <tr>
        <td colspan="7">
        <%if strImpersonateName <> "" then%>
            <font size="2"><strong><u><br>Rejected AVs (Pulsar Product)</u></strong></font>&nbsp;-&nbsp;<strong><font color=red><%=strImpersonateName%></strong></font>
        </td>
	    <%else%>
	        <font size="2"><strong><u><br>Rejected AVs (Pulsar Product)</u></strong></font>
        </td>
	    <%end if%>	    
    </tr>
    <tr>
        <td nowrap align="left">
             <% 
      if strImpersonateName <> "" then
          btnUpdate = "<input type=""button"" value=""Action Item Complete"" id=""btnUpdateRejectedAVs"" name=""btnUpdateRejectedAVs"" class=""button2"" style=""width:145px"" onclick=""return cmdRejectedAvsComplete_onclick(" & "'" & strImpersonateName & "'" & ")"">"
	      Response.Write(btnUpdate)    
	  else
          btnUpdate = "<input type=""button"" value=""Action Item Complete"" id=""btnUpdateRejectedAVs"" name=""btnUpdateRejectedAVs"" class=""button2"" style=""width:145px"" onclick=""return cmdRejectedAvsComplete_onclick(" & "'" & CurrentUserName & "'" & ")"">"
	      Response.Write(btnUpdate)
      end if
	%>
        </td>
    </tr>
</table>

 <table id="RejectedAvsItemsPulsar" border="0" width="100%" cellspacing="0" cellpadding="2">
     <thead>
       <tr bgcolor="beige">
        <td valign="top" style="width:30px;" >
            <input id="chkRejectedAVsAll" type="checkbox" style="height:16px;width:16px" onclick="javascript: chkRejectedAvsAll_onclick();">
        </td>
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product</b></font></td>
	    <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">AV Part No.</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Marketing Description(100 Char)</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">PHweb Action</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Reason Code</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Sub-Reason Code</b></font></td>	
        <td ><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Comments</b></font></td>	
     </tr>
     </thead>
     <%
         Dim m_IsMarketingUser
         m_IsMarketingUser = securityObj.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = securityObj.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = securityObj.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If

         dim strRowBorderColor:strRowBorderColor="Gainsboro"
    set rs = server.CreateObject("ADODB.recordset")
	Dim dw, cmd
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_GetRejectedAvs")	
    If strImpersonateID = "" Then
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(CurrentUserID)
    Else
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(strImpersonateID)
    End If
    Set rs = dw.ExecuteCommAndReturnRS(cmd)
     
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD><font size=1 face=verdana><b>none</b></font></TD></TR>"
	else	 
      do while not rs.EOF

    %>
    <tr bgcolor="ivory" onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()" id="RejectedAVsPulsarTr<%=rs("AVPHwebRejectionItemID")%>">
         <td valign="top" style="width:30px;BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL ">
            <input id="chkRejectedAVsItems" name="chkRejectedAVsItems" type="checkbox" style="height:16px;width:16px" value="<%=rs("AVPHwebRejectionItemID")%>">
        </td>  
        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
            <font class="text" size="1"><%= rs("Product")%></font> 
		</td>
        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
            <font class="text" size="1"><%= rs("Avno")%></font> 
		</td>
         <td id="tdMarketingDesc<%=rs("ProductBrandID")%><%=rs("AvDetailID")%>"  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
            <font class="text" size="1"><%= rs("marketingdescription")%></font> 
		</td>
        <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
            <font class="text" size="1"><%= rs("PHwebAction")%></font> 
        </td>
         <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
            <font class="text" size="1"><%= rs("Reasoncode")%></font> 
		</td>
        <%If rs("Sub_ReasonCode") <> "" Then %>
             <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
                <font class="text" size="1"><%= rs("Sub_ReasonCode")%></font> 
		    </td>
        <%Else%>
	        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
               &nbsp
		    </td>
        <%End If%> 
        <%If rs("comments") <> "" Then %>
           <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
                <font class="text" size="1"><%= rs("comments")%></font> 
		    </td>
        <%Else%>
	        <td  style="BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK:KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowRejectedAVDetail(<%=rs("ProductVersionID")%>,<%=rs("AvDetailID")%>,<%=rs("ProductBrandID")%>,<%=clng(CurrentUserID)%>);">
               &nbsp
		    </td>
        <%End If%> 
         
    </tr>
   <%
          rs.MoveNext
	loop

     

	end if
	rs.Close
     %>
 </table> 
<div id="divOpenFeaturePopUp" title="Coolbeans" style="display: none;">
        <iframe frameborder="0" name="ifOpenFeaturePopUp" id="ifOpenFeaturePopUp" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>