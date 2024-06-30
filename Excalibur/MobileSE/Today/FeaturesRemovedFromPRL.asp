<script type="text/javascript">
    function CreateFeatureRemovedMenu(FeatureActionItemID, ProductVersionID, CurrentUser) {
        document.getElementById("txtCurrentUser").value = CurrentUser;
         // The variables "lefter" and "topper" store the X and Y coordinates
        // to use as parameter values for the following show method. In this
        // way, the popup displays near the location the user clicks. 
        var lefter = event.clientX + document.documentElement.scrollLeft;
        var topper = event.clientY + document.documentElement.scrollTop;
        var popupBody;

        if (window.event.srcElement.className == "text") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement.parentElement)
                    SelectedRow.style.color = "black";
            SelectedRow = window.event.srcElement.parentElement.parentElement;
            SelectedRow.style.color = "red";

        }
        else if (window.event.srcElement.className == "cell") {
            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != window.event.srcElement.parentElement)
                    SelectedRow.style.color = "black";

            SelectedRow = window.event.srcElement.parentElement;
            SelectedRow.style.color = "red";
        }

        popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

        //Obsolete Localized AV        
        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:FeatureChange_oncontextmenu(" + FeatureActionItemID + "," + 1 + ")'\" >&nbsp;&nbsp;&nbsp;Obsolete&nbsp;AV</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";            

        //Not Actionable
        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:FeatureChange_oncontextmenu(" + FeatureActionItemID + "," + 2 + ")'\" >&nbsp;&nbsp;&nbsp;Not&nbsp;Actionable</SPAN></FONT></DIV>";

        popupBody = popupBody + "</DIV>";
         
        
        oPopup.document.body.innerHTML = popupBody;

        oPopup.show(lefter, topper, 130, 85, document.body);

        //Adjust window size
        var NewHeight;
        var NewWidth;

        NewHeight = oPopup.document.body.scrollHeight;
        NewWidth = oPopup.document.body.scrollWidth;
        oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
    }
    function FeatureChange_oncontextmenu(FeatureActionItemID, ActionType) {
        oPopup.hide();
        var CurrentUser = "";
        CurrentUser = document.getElementById("txtCurrentUser").value;
        
        SetFeatureChangeToItem(FeatureActionItemID, CurrentUser, ActionType);

    }    
    function RemoveRow(FeatureActionItemID)
    {
        $('#trFeatureRemoved' + FeatureActionItemID).closest("tr").remove();
    }
    function SetFeatureChangeToItem(FeatureActionItemID, CurrentUser, ActionType) {
        var CurrentUser = "";
        CurrentUser = document.getElementById("txtCurrentUser").value;
        var parameters = "Function=UpdateFeatureAction&FeatureActionItemID=" + FeatureActionItemID + "&ActionType=" + ActionType + "&CurrentUserName=" + CurrentUser;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {// Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "FeaturesRemovedFromPRLActions.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
        RemoveRow(FeatureActionItemID);       
    }  

</script>
<%
If strImpersonateID = "" Then
    rs.Open "usp_Today_GetFeatureRemovedFromPRL " & CLng(CurrentUserID), cn, adOpenStatic
Else
    rs.Open "usp_Today_GetFeatureRemovedFromPRL " & CLng(strImpersonateID), cn, adOpenStatic
End If

If Not (rs.Bof And rs.Eof) Then
%>
<form ID=frmImageTabChange method=post>
<table id="tblOVHeaders" border="0" cellspacing="0" cellpadding="2">
    <tr>
        <td colspan="7">
            <input type="text" id="txtCurrentUser" value="" style="display: none" />
            <%if strImpersonateName <> "" then%>
                <font size="2"><strong><u><br>Feature removed from POST PRL Lock</u></strong></font>&nbsp;-&nbsp;<strong><font color="red"><%=strImpersonateName%></strong></font>
            <%else%>
                <font size="2"><strong><u><br>Feature removed from POST PRL Lock</u></strong></font>
        </td>
	    <%end if%>        
    </tr>
    <!--<tr>
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
            
        </td>
    </tr>-->
</table>

<table id="tblImageChanges" border="0" width="100%" cellspacing="0" cellpadding="2">
    <thead>
        <tr bgcolor="beige">
            <td style="width: 10%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product Name</b></font></td>
            <td style="width: 20%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">SCM Name</b></font></td>
            <td style="width: 10%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Feature ID</b></font></td>
            <td style="width: 70%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Feature Full Name</b></font></td>
        </tr>
    </thead>
    <%
	do while not rs.EOF  
    %>
    <tr bgcolor="ivory" onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()" id="trFeatureRemoved<%=rs("FeatureActionItemID")%>" onclick="CreateFeatureRemovedMenu(<%=rs("FeatureActionItemID")%>,<%=rs("ProductVersionID")%>,'<%=CurrentUserName%>');" >
        <td style="width: 10%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("ProductName")%></font>
        </td>
        <td style="width: 20%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("SCMName")%></font>
        </td>
        <td style="width: 10%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("FeatureID")%></font>
        </td>
        <td style="width: 70%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
            <font class="text" size="1"><%= rs("FeatureName")%></font>
        </td>
    </tr>
    <%
    rs.MoveNext
	loop	
    %>
</table>
    </form>
<% 
End If 
rs.Close
%>
