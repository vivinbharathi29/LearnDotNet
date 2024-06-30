<script type="text/javascript">
    
    function CreateAvMenuPulsar(AvCreateID, CurrentUserId, FeatureId, FeatureName, ProductBrandId, SCMCategoryID, ProductVersionID) {
        if(SCMCategoryID == "")
            SCMCategoryID = 0;
        
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

        //popupBody = popupBody + "<DIV>";
        //popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:CreateAVsPulsar_oncontextmenu(" + AvCreateID + "," + CurrentUserId + "," + 0 + ")'\" >&nbsp;&nbsp;&nbsp;Create&nbsp;AV...</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:CreateAVsPulsar_oncontextmenu(" + AvCreateID + "," + CurrentUserId + "," + 1 + ")'\" >&nbsp;&nbsp;&nbsp;Not&nbsp;Actionable...</SPAN></FONT></DIV>";

        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddExistingAVNoToFeature(" + AvCreateID + "," + CurrentUserId + "," + FeatureId + ",&quot;" + FeatureName + "&quot;," + ProductBrandId + "," + ProductVersionID + ")'\" >&nbsp;&nbsp;&nbsp;Enter&nbsp;Existing&nbsp;AV&nbsp;No.&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";
        // popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:AddExistingAVNoToFeature(" + AvCreateID + "," + CurrentUserId +  "," + FeatureName + "," + ProductBrandId + ")'\" >&nbsp;&nbsp;&nbsp;Enter&nbsp;Existing&nbsp;AV&nbsp;No.&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";
        popupBody = popupBody + "<DIV>";
        popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

        popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
        popupBody = popupBody + "<FONT face=Arial size=2>";
       // popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeCategory_oncontextmenu(" + AvCreateID + "," + SCMCategoryID + ")'\" >&nbsp;&nbsp;&nbsp;Change&nbsp;Category&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";
        popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeCategory_oncontextmenu(" + AvCreateID + "," + CurrentUserId + "," + FeatureId + ",&quot;" + FeatureName + "&quot;," + ProductBrandId + ",&quot;" +  SCMCategoryID + "&quot;," + ProductVersionID +  ")'\" >&nbsp;&nbsp;&nbsp;Change&nbsp;Category&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

        //popupBody = popupBody + "<DIV>";
        //popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

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

    //function ChangeCategory_oncontextmenu(AvCreateID, CategoryID) {
    function ChangeCategory_oncontextmenu(AvCreateID, CurrentUserId, FeatureId, FeatureName, ProductBrandId, CategoryID, ProductVersionID) {

        var Height = ($(window).height() * 15) / 100;
        var Width = ($(window).width() * 25) / 100;
        var result = window.showModalDialog("/Pulsar/scm/ListCategory?SelectedCategoryID=" + CategoryID, "SCM Category", "dialogWidth:" + Width + "px;dialogHeight:" + Height + "px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No; scroll: No");

        if (typeof (result) != "undefined") {
            var parameters = "function=ChangeCategory&AvCreateID=" + AvCreateID + "&CategoryID=" + result.CategoryID;
            var request = null;
            //Initialize the AJAX variable.
            if (window.XMLHttpRequest) {// Are we working with mozilla
                request = new XMLHttpRequest(); //Yes -- this is mozilla.
            } else { //Not Mozilla, must be IE
                request = new ActiveXObject("Microsoft.XMLHTTP");
            } //End setup Ajax.
            request.open("POST", "CreateAv.asp", false);
            request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
            request.send(parameters);
            //frmSimpleAVsPulsar.submit();
            if (request.readyState == 4) {
                var trinput = $("#SimpleAVsPulsarTr" + AvCreateID);
                if (trinput) {
                    trinput.attr("catID", result.CategoryID);
                    $(':checkbox', trinput).val(AvCreateID + "-" + result.CategoryID);
                    trinput.closest("tr").children("td").each(function () {
                        var sfunction = $(this).attr("onclick");
                        $(this).removeAttr("onclick");
                        if (sfunction) {
                            sfunction = sfunction.replace(",'" + CategoryID + "',", ",'" + result.CategoryID + "',");
                            $(this).on("click", function (ev) {
                                var newfn = sfunction;
                                var NewCatID = $(this).parent('tr').attr("catID");
                                var OldCateID = result.CategoryID;
                                if (NewCatID != OldCateID)
                                    newfn = newfn.replace(",'" + OldCateID + "',", ",'" + NewCatID + "',");

                                var func = new Function(newfn);
                                func();
                            });
                        }
                    });
                }
                var tdinput = document.getElementById("CSAP" + AvCreateID);
                if (tdinput) {
                    tdinput.innerHTML = "<font class='text' size='1'>" + result.CategoryName + "&nbsp;</font>";
                }
            }
        }
    }

    function AddExistingAVNoToFeature(AvCreateID, CurrentUserId, FeatureId, FeatureName, ProductBrandId, ProductVersionID) {
        var strID;
        oPopup.hide();
        var trinput = document.getElementById("SimpleAVsPulsarTr" + AvCreateID);
        var SCMcat = trinput.CatID; // this property is refreshed with the change category context menu.

        var url = "<%=AppRoot %>/MobileSE/Today/AddAVNoToFeatureFrame.asp?AvCreateID=" + AvCreateID + "&CurrentUserId=" + CurrentUserId + "&FeatureId=" + FeatureId + "&FeatureName=" + FeatureName + "&ProductBrandId=" + ProductBrandId + "&SCMCategoryId=" + SCMcat + "&ProductVersionID=" + ProductVersionID;
        //strID = window.parent.showModalDialog(url, "", "dialogWidth:423px;dialogHeight:180px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        //frmSimpleAVsPulsar.submit();

        ShowPopupExistingAVNoToFeature(url, "Enter existing AV " , 500, 220);

    }

    function ShowPopupExistingAVNoToFeature(QueryString, Title, DlgWidth, DlgHeight) {
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

    function CloseExistingAVNoToFeature(Refresh,AvDetailID) {
        $("#modalDialog").attr("src", "");
        $("#iframeDialog").dialog("close");

        if (Refresh == 1)
        {
            $('#SimpleAVsPulsarTr' + AvDetailID).closest("tr").remove();
        }
    }

    function btnChangeCategory_onclick() {
        var Height = ($(window).height() * 15) / 100;
        var Width = ($(window).width() * 25) / 100;
		
        if ($(":checkbox[name='chkCreateAVsPulsar']").is(':checked'))
        {
            var SelectedCategoryID = $('input:checkbox[name="chkCreateAVsPulsar"]:checked:visible:first').val().split('-')[1];
            var result = "";            
            result = window.showModalDialog("/Pulsar/scm/ListCategory?SelectedCategoryID=" + SelectedCategoryID, "SCM Category", "dialogWidth:" + Width + "px;dialogHeight:" + Height + "px;edge: raised ;center:Yes; help: No;resizable: Yes;status: No; scroll: No");
            if (typeof (result) != "undefined") {
                $('input:checkbox[name="chkCreateAVsPulsar"]:checked').each(function () {         	
                    var ids = $(this).val().split('-');
					var parameters = "function=ChangeCategory&AvCreateID=" + ids[0] + "&CategoryID=" + result.CategoryID;
                    var request = null;
                    //Initialize the AJAX variable.
                    if (window.XMLHttpRequest) {// Are we working with mozilla
                        request = new XMLHttpRequest(); //Yes -- this is mozilla.
                    } else { //Not Mozilla, must be IE
                        request = new ActiveXObject("Microsoft.XMLHTTP");
                    } //End setup Ajax.
                    request.open("POST", "CreateAv.asp", false);
                    request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                    request.send(parameters);
                    if (request.readyState == 4) {
                        $(this).val(ids[0] + "-" + result.CategoryID);
                        var trinput = $("#SimpleAVsPulsarTr" + ids[0]);
                        if (trinput) {
                            trinput.attr("catID", result.CategoryID);
                            trinput.closest("tr").children("td").each(function () {
                                var sfunction = $(this).attr("onclick");
                                $(this).removeAttr("onclick");
                                if (sfunction) {
                                    sfunction = sfunction.replace(",'" + SelectedCategoryID + "',", ",'" + result.CategoryID + "',");
                                    $(this).on("click", function (ev) {
                                        var newfn = sfunction;
                                        var NewCatID = $(this).parent('tr').attr("catID");
                                        var OldCateID = result.CategoryID;
                                        if(NewCatID != OldCateID)
                                            newfn = newfn.replace(",'" + OldCateID + "',", ",'" + NewCatID + "',");

                                        var func = new Function(newfn);
                                        func();
                                    });
                                }
                            });
                            var tdinput = document.getElementById("CSAP" + ids[0]);
                            if (tdinput) {                                
                                tdinput.innerHTML = "<font class='text' size='1'>" + result.CategoryName + "&nbsp;</font>";
                                $(this).prop("checked", false);
                            }
                        }
                    }                    
                });
            }
        }
        else {
            alert("Please select row(s) to change category");
        }
    }

    function btnCreateAVsPulsar_onclick(CurrentUserId, NotActionable) {
        
        var message = "";
        if ($(":checkbox[name='chkCreateAVsPulsar']").is(':checked')) {
            $('input:checkbox[name="chkCreateAVsPulsar"]:checked').each(function () {
                var ids = $(this).val().split('-');
                var parameters = "function=CreateSimpleAv&AvCreateID=" + ids[0] + "&NotActionable=" + NotActionable + "&UserID=" + CurrentUserId + "&Pulsar=1";
                var request = null;
                //Initialize the AJAX variable.
                if (window.XMLHttpRequest) {// Are we working with mozilla
                    request = new XMLHttpRequest(); //Yes -- this is mozilla.
                } else { //Not Mozilla, must be IE
                    request = new ActiveXObject("Microsoft.XMLHTTP");
                } //End setup Ajax.
                request.open("POST", "CreateAv.asp", false);
                request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                request.send(parameters);
                // get the feature names if GPG description is empty from the stored procedure
                var result = request.responseText;
                if (request.readyState == 4) {
                    if (result == "") { // GPG description not empty 
                        $(this).prop("checked", false);
                        $('#SimpleAVsPulsarTr' + ids[0]).closest("tr").remove();
                    }
                    else // GPG description empty 
                    {
                        $(this).prop("checked", false);
                        message = message + ', ' + result;
                    }
                }                

            });

            if (message != "")
            {
                message = message.substring(2, message.length);
                alert("The following AV(s) are/were not created because the GPG description is blank. Please complete the GPG description for the missing Features and start this process again : " + message)
            }
            
        } else {
                        
            if(NotActionable == 1)
                alert("Please select row(s) to set Not Actionable");
            else 
                alert("Please select row(s) to Create Avs");
        }        
    }

    function CreateAVsPulsar_oncontextmenu(AvCreateID, CurrentUserId, NotActionable) {
        oPopup.hide();
        var parameters = "function=CreateSimpleAv&AvCreateID=" + AvCreateID + "&NotActionable=" + NotActionable + "&UserID=" + CurrentUserId + "&Pulsar=1";
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {// Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "CreateAv.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
        // get the feature names if GPG description is empty from the stored procedure
        var result = request.responseText;
        if (request.readyState == 4) {
            if (result == "") { // GPG description not empty 
                $('#SimpleAVsPulsarTr' + AvCreateID).closest("tr").remove();
            }
            else // GPG description empty 
            {
                alert("The following AV was not created because the GPG description is blank. Please complete the GPG description for the missing Features and start this process again : \n" + result);
            }
        }
        //frmSimpleAVsPulsar.submit();
    }

    function chkCreateAVsAllPulsar_onclick() {
        var CheckAll = $("#chkCreateAVsAllPulsar").is(':checked');
        $('input:checkbox[name="chkCreateAVsPulsar"]').each(function () {
            $(this).prop("checked", CheckAll);
        });
    }
</script>
<%
 
If strImpersonateID = "" Then
    rs.Open "usp_SelectSimpleAVCreateList_Pulsar " & CLng(CurrentUserID), cn, adOpenStatic
Else
    rs.Open "usp_SelectSimpleAVCreateList_Pulsar " & CLng(strImpersonateID), cn, adOpenStatic
End If

If Not (rs.Bof And rs.Eof) Then 
%>
<form ID=frmSimpleAVsPulsar method=post>
<table ID="SimpleAVsPulsar" border=0 width="100%" cellspacing="0" cellpadding="2">
<thead>
  <tr>
    <td colspan="8">
    <%if strImpersonateName <> "" then%>
         <font size="2"><strong><u><br>Create Simple AVs (Pulsar Product)</u></strong></font>&nbsp;-&nbsp;<strong><font color=red><%=strImpersonateName%></strong></font>
	<%else%>
	     <font size="2"><strong><u><br>Create Simple AVs (Pulsar Product)</u></strong></font>
	<%end if 
	     btnCreateAVs = "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" value=""Create AVs"" id=""btnCreateAVsPulsar"" name=""btnCreateAVsPulsar"" class=""button2"" style=""width:100px"" onclick=""return btnCreateAVsPulsar_onclick(" & CurrentUserId & ",0)"">"
	     Response.Write(btnCreateAVs)
	     btnNonActionable = "&nbsp;&nbsp;<input type=""button"" value=""Not Actionable"" id=""btnNonActionablePulsar"" name=""btnNonActionablePulsar"" class=""button2"" style=""width:120px"" onclick=""return btnCreateAVsPulsar_onclick(" & CurrentUserId & ",1)"">"
	     Response.Write(btnNonActionable)
         btnChangeCategory = "&nbsp;&nbsp;<input type=""button"" value=""Change Category"" id=""btnChangeCategory"" name=""btnChangeCategory"" class=""button2"" style=""width:140px"" onclick=""return btnChangeCategory_onclick()""></td>"
	     Response.Write(btnChangeCategory)
	%>
		
  </tr>
  <tr bgcolor="beige">
    <td><input id="chkCreateAVsAllPulsar" type="checkbox" style="height:16px;width:16px" onclick="javascript: chkCreateAVsAllPulsar_onclick();"></td>
	<td onclick="SortTable( 'SimpleAVsPulsar', 1 ,1,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">FeatureID</b></font></td>
	<td onclick="SortTable( 'SimpleAVsPulsar', 2 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Product</b></font></td>
	<td onclick="SortTable( 'SimpleAVsPulsar', 3 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Brand</b></font></td>
	<td onclick="SortTable( 'SimpleAVsPulsar', 4 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Feature</b></font></td>
    <td onclick="SortTable( 'SimpleAVsPulsar', 5 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Component Linkage</b></font></td>
	<td onclick="SortTable( 'SimpleAVsPulsar', 6 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">SCM Category</b></font></td>   
    <td onclick="SortTable( 'SimpleAVsPulsar', 7 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Requested From PRL/Platform</b></font></td>  
    <td onclick="SortTable( 'SimpleAVsPulsar', 8 ,0,2);"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Release</b></font></td>  
  </tr>
  </thead>
<%
	do while not rs.EOF
	  
%>
  <tr bgcolor="ivory" onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()" id="SimpleAVsPulsarTr<%=rs("AvCreateID")%>" CatID="<%=rs("SCMCategoryID")%>">
        <td valign=top style="BORDER-TOP: <%=strRowBorderColor%> thin solid"><input id="chkCreateAVsPulsar" name="chkCreateAVsPulsar" type="checkbox" style="height:16px;width:16px" value="<%=rs("AvCreateID") & "-" & rs("SCMCategoryID")%>"</td>  
		<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=rs("FeatureID")%> </font> </td>
		<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=Server.HTMLEncode(rs("DOTSName"))%> </font></td>
		<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=Server.HTMLEncode(rs("BrandName"))%> </font></td>		
		<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=Server.HTMLEncode(rs("FeatureName"))%></font></td>
      	<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=rs("ComponentLinkage")%></font></td>
		<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" id='CSAP<%=rs("AvCreateID")%>'><font class="text" size="1"><%=rs("Category")%>&nbsp;</font></td>
        <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=rs("RequestFrom")%>&nbsp;</font></td>
      <td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap onclick="CreateAvMenuPulsar(<%=rs("AvCreateID")%>,<%=CurrentUserId%>,<%=rs("FeatureID")%>,'<%=rs("FeatureName")%>',<%=rs("ProductBrandID")%>,'<%=rs("SCMCategoryID")%>',<%=rs("ProductVersionID")%>);" ><font class="text" size="1"><%=rs("Release")%>&nbsp;</font></td>
  </tr>
<%
	rs.MoveNext
	loop
    response.Write "</table></form><BR>"
End If 
rs.Close
%>