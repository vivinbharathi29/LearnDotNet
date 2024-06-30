<%
    strRowBorderColor="Gainsboro"
    set rs = server.CreateObject("ADODB.recordset")
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cmd = dw.CreateCommAndSP(cn, "usp_TodayPage_GetSharedAVActionItems")	
    If strImpersonateID = "" Then
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(CurrentUserID)
    Else
         dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 8, clng(strImpersonateID)
    End If
    Set rs = dw.ExecuteCommAndReturnRS(cmd)
%>
<script type="text/javascript">
    function ShowSharedAVDetail(AVDetailID, CurrentUser, ReasonType) {
        window.open("/IPulsar/Admin/SCM/SharedAV_Main.aspx?FromTodayPage=1&AVDetailID=" + AVDetailID + "&CurrentUserID=" + CurrentUser + "&ReasonType=" + ReasonType, "_parent");
    }
</script>
<% If Not (rs.Bof And rs.Eof) Then %>

<table id="tblSharedAVActionsHeader" border="0" cellspacing="0" cellpadding="2">
    <tr>
        <td colspan="7">
            <%if strImpersonateName <> "" then%>
                <font size="2"><strong><u><br>Shared AV Actions</u></strong></font>&nbsp;-&nbsp;<strong><font color="red"><%=strImpersonateName%></strong></font>
            <%else%>
                <font size="2"><strong><u><br>Shared AV Actions</u></strong></font>
            <%end if%>
        </td>	    
    </tr>    
</table>

<table id="tblSharedAVActions" border="0" width="100%" cellspacing="0" cellpadding="2">
    <thead>
        <tr bgcolor="beige">            
            <td style="width: 8%; "><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">AV No.</b></font></td>
            <td style="width: 20%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">AV Marketing Long Description <br /> (100 Char)</b></font></td>
            <td style="width: 5%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Missing Load to PHweb</b></font></td>
            <td style="width: 20%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">SCM(s) Where Used</b></font></td>
           <!-- remove FCS from all areas - task 20243-->
            <td style="width: 7%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">RTP/MR Date</b></font></td>
            <td style="width: 7%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">PA:AD <br />(Intro) Date</b></font></td>
            <td style="width: 7%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Select Availability <br />(SA) Date</b></font></td>
            <td style="width: 7%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">General Availability <br />(GA) Date</b></font></td>
            <td style="width: 7%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">End of Manufacturing <br />(EM) Date</b></font></td>
            <td style="width: 12%;"><font size="1"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Reason</b></font></td>
        </tr>
    </thead>
    <%
	do while not rs.EOF  
    %>
    <tr bgcolor="ivory" onmouseover="return openrows_onmouseover()" onmouseout="return openrows_onmouseout()" id="trSharedAvAction<%=rs("AVDetailID")%>">        
        <%If rs("AvNo") <> "" Then %>
        <td style="width: 8%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("AvNo")%></font>
        </td>
        <%Else%>
        <td style="width: 8%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <%If rs("MarketingDescriptionPMG") <> "" Then %>
        <td style="width: 20%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("MarketingDescriptionPMG")%></font>
        </td>
        <%Else%>
        <td style="width: 20%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <%If rs("MissingLoadToPHwebYN") <> "" Then %>
        <td style="width: 5%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" align="center" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1" color="red"><%= rs("MissingLoadToPHwebYN")%></font>
        </td>
        <%Else%>
        <td style="width: 5%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <%If rs("SCMsWhereUsed") <> "" Then %>
        <td style="width: 20%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("SCMsWhereUsed")%></font>
        </td>
        <%Else%>
        <td style="width: 20%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>

        <%If rs("RTPDate") <> "" Then %>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("RTPDate")%></font>
        </td>
        <%Else%>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>

        <%If rs("PhwebDate") <> "" Then %>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("PhwebDate")%></font>
        </td>
        <%Else%>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>

        <%If rs("CPLBlindDt") <> "" Then %>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("CPLBlindDt")%></font>
        </td>
        <%Else%>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <%If rs("GeneralAvailDt") <> "" Then %>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("GeneralAvailDt")%></font>
        </td>
        <%Else%>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <%If rs("RASDiscontinueDt") <> "" Then %>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("RASDiscontinueDt")%></font>
        </td>
        <%Else%>
        <td style="width: 7%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>&nbsp</td>
        <%End If%>
        <td style="width: 12%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" onclick="ShowSharedAVDetail(<%=rs("AvDetailID")%>,<%=clng(CurrentUserID)%>,<%=rs("ReasonID")%>);">
            <font class="text" size="1"><%= rs("Reason")%></font>
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
