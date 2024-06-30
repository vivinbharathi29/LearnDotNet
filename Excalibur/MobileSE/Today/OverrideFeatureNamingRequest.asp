<%
    strRowBorderColor="Gainsboro"
    set rs = server.CreateObject("ADODB.recordset")
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cmd = dw.CreateCommAndSP(cn, "usp_Today_FeatureNamingOverrideRequestedList")	
    If strImpersonateID = "" Then
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(CurrentUserID)
    Else
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(strImpersonateID)
    End If
    Set rs = dw.ExecuteCommAndReturnRS(cmd)
%>
<script type="text/javascript">
    function ShowOverrideFeatureDetail(FeatureID, CurrentUser) {
        document.getElementById("txtCurrentUser").value = CurrentUser;
        ShowFeaturePropertiesDialog("/IPulsar/Features/FeatureProperties.aspx?FromModule=1&FeatureID=" + FeatureID + "&AltNamingTodayPage=1", "Feature Properties", 1200, 800);
    }
    function CloseFeaturePropertiesPopUp(refresh, FeatureID) {      
        ClosePropertiesDialog();
    }
    function CompleteAction(FeatureID, FeatureName) {
        ClosePropertiesDialog();
        var CurrentUser = "";
        CurrentUser = document.getElementById("txtCurrentUser").value;
        var parameters = "Function=UpdateFeature&FeatureIDs=" + FeatureID + "&CurrentUserName=" + CurrentUser + "&ActionType=0";
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {// Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "OverrideFeatureActions.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);        
        
        if (request.readyState == 4) {
            //send email
            var email = "";
            email = document.getElementById("tdEmail" + FeatureID).innerHTML;
            parameters = "Function=SendEmail&FeatureIDs=" + FeatureID + "&FeatureName=" + FeatureName.split(" ").join("%20") + "&CurrentUserName=" + CurrentUser + "&ActionType=0" + "&Email=" + email;
            request = null;
            //Initialize the AJAX variable.
            if (window.XMLHttpRequest) {// Are we working with mozilla
                request = new XMLHttpRequest(); //Yes -- this is mozilla.
            } else { //Not Mozilla, must be IE
                request = new ActiveXObject("Microsoft.XMLHTTP");
            } //End setup Ajax.
            request.open("POST", "OverrideFeatureActions.asp", false);
            request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
            request.send(parameters);
            //remove from the grid
            $(this).prop("checked", false);
            $('#trOverrideFeatures' + FeatureID).closest("tr").remove();
        }
        
    }
    function GetNotActionableComment() {
        var strComment;
        strComment = window.showModalDialog("OverrideFeatureFeedback.asp?", "", "dialogWidth:595px;dialogHeight:165px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No");
        return strComment;
    }
    function SetNotActionableAction(CurrentUser) {        
        var i, strComment="";
        var elemsIDChecked;
        elemsIDChecked = "";
      
        if ($(":checkbox[name='chkOverrideFeatures']").is(':checked')) {
            strComment = GetNotActionableComment();
            if (strComment == "Cancel")
                return;
            //gets all the features id and update the database and remove from the today page
            $('input:checkbox[name="chkOverrideFeatures"]:checked').each(function () {
                elemsIDChecked = $(this).val();
                var parameters = "Function=UpdateFeature&FeatureIDs=" + elemsIDChecked + "&CurrentUserName=" + CurrentUser + "&ActionType=1";
                var request = null;
                //Initialize the AJAX variable.
                if (window.XMLHttpRequest) {// Are we working with mozilla
                    request = new XMLHttpRequest(); //Yes -- this is mozilla.
                } else { //Not Mozilla, must be IE
                    request = new ActiveXObject("Microsoft.XMLHTTP");
                } //End setup Ajax.
                request.open("POST", "OverrideFeatureActions.asp", false);
                request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                request.send(parameters);
                if (request.readyState == 4) {
                    var email = "";
                    email = document.getElementById("tdEmail" + elemsIDChecked).innerHTML;
                    parameters = "Function=SendEmail&FeatureIDs=" + elemsIDChecked + "&Comment=" + encodeURIComponent(strComment) + "&CurrentUserName=" + CurrentUser + "&ActionType=1" + "&Email=" + email;
                    
                    request = null;
                    //Initialize the AJAX variable.
                    if (window.XMLHttpRequest) {// Are we working with mozilla
                        request = new XMLHttpRequest(); //Yes -- this is mozilla.
                    } else { //Not Mozilla, must be IE
                        request = new ActiveXObject("Microsoft.XMLHTTP");
                    } //End setup Ajax.
                    request.open("POST", "OverrideFeatureActions.asp", false);
                    request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                    request.send(parameters);

                    $(this).prop("checked", false);
                    $('#trOverrideFeatures' + elemsIDChecked).closest("tr").remove();
                }
            });        
        }
        else {
            alert("Please select Features to set to 'Not Actionable'");
        }
    }

    function chkOverrideFeaturesAll_onclick() {
        var i;
        var checkBoxes = document.getElementsByTagName("input");
        var chkCreateAVsAll, chkBoxName;
        chkRejectedAvsItemsAll = document.getElementById("chkOverrideFeatureAll");
        chkBoxName = "chkOverrideFeatures";
        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].name == chkBoxName) {
                checkBoxes[i].checked = chkRejectedAvsItemsAll.checked;
            }
        }
    }

</script>
<% If Not (rs.Bof And rs.Eof) Then
' Sort by Date Requested in ascending order
rs.Sort="DateRequested"    
%>

<table id="tblOVHeaders" border="0" cellspacing="0" cellpadding="2">
    <tr>
        <td colspan="7">
            <%if strImpersonateName <> "" then%>
            <font size="2"><strong><u><br>Feature Naming Override Request</u></strong></font>&nbsp;-&nbsp;<strong><font color="red"><%=strImpersonateName%></strong></font>
        </td>
        <%else%>
        <font size="2"><strong><u><br>Feature Naming Override Request</u></strong></font>
        </td>
	    <%end if%>
    </tr>
    <tr>
        <td nowrap align="left">
            <% 
      if strImpersonateName <> "" then
          btnUpdate = "<input type=""button"" value=""Not Actionable"" id=""btnNotActionable"" name=""btnNotActionable"" class=""button2"" style=""width:100px"" onclick=""return SetNotActionableAction(" & "'" & strImpersonateName & "'" & ")"">"
	      Response.Write(btnUpdate)    
	  else
          btnUpdate = "<input type=""button"" value=""Not Actionable"" id=""btnNotActionable"" name=""btnNotActionable"" class=""button2"" style=""width:100px"" onclick=""return SetNotActionableAction(" & "'" & CurrentUserName & "'" & ")"">"
	      Response.Write(btnUpdate)
      end if
            %>
            <input type="text" id="txtCurrentUser" value="" style="display: none" />
        </td>
    </tr>
</table>

<table id="tblOverrideFeatures" border="0" width="100%" cellspacing="0" cellpadding="2">
    <thead>
        <tr bgcolor="beige">
            <td valign="top" style="width: 2%;">
                <input id="chkOverrideFeatureAll" type="checkbox" style="height: 16px; width: 16px" onclick="javascript: chkOverrideFeaturesAll_onclick();">
            </td>
            <td style="width: 8%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Feature ID</b></font></td>
            <td style="width: 67%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Reason for Override Request</b></font></td>
            <td style="width: 13%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Requested By</b></font></td>
            <td style="width: 10%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Date Requested</b></font></td>
        </tr>
    </thead>
    <%
	do while not rs.EOF  
    %>
    <tr bgcolor="ivory" onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()" id="trOverrideFeatures<%=rs("FeatureID")%>">
        <td valign="top" style="width: 2%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL">
            <input id="chkOverrideFeatures" name="chkOverrideFeatures" type="checkbox" style="height: 16px; width: 16px" value="<%=rs("FeatureID")%>">
        </td>
        <td style="width: 8%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowOverrideFeatureDetail(<%=rs("FeatureID")%>,'<%=CurrentUserName%>');">
            <font class="text" size="1"><%= rs("FeatureID")%></font>
        </td>
        <%If rs("OverrideReason") <> "" Then %>
        <td style="width: 67%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowOverrideFeatureDetail(<%=rs("FeatureID")%>,'<%=CurrentUserName%>');">
            <font class="text" size="1"><%= rs("OverrideReason")%></font>
        </td>
        <%Else%>
        <td style="width: 67%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <td style="width: 13%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("RequestedBy")%></font>
        </td>
        <td style="width: 10%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("DateRequested")%></font>
        </td>
        <td id="tdEmail<%=rs("FeatureID")%>" style="width: 0%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL; display:none" class="cell" valign="top" nowrap>
            <%= rs("Email")%>
        </td>
    </tr>
    <%
    rs.MoveNext
	loop	
    %>
</table>
<% 
End If 
rs.Close
%>
