<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    frmUpdate.txtRev.focus();
}


function cmdImageButton_onmouseover() {
	window.event.srcElement.style.cursor = "default";
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
}

function cmdImageButton_onmouseout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
}

function cmdImageButton_onmousedown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";

}

function cmdImageButton_onmouseup() {
    window.event.srcElement.style.borderColor = "gold";
    window.event.srcElement.style.borderStyle = "solid";
    window.event.srcElement.style.backgroundColor = "gainsboro";
    ImageButton_Pressed(window.event.srcElement.name);
}

function cmdImageButton2_onmouseup() {
    window.event.srcElement.style.borderColor = "gold";
    window.event.srcElement.style.borderStyle = "solid";
    window.event.srcElement.style.backgroundColor = "gainsboro";
    ImageButton2_Pressed(window.event.srcElement.name);
}

function cmdImageButton_onkeydown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";
}

function cmdImageButton_onkeyup() {
    window.event.srcElement.style.borderColor = "";
    window.event.srcElement.style.borderStyle = "solid";
    window.event.srcElement.style.backgroundColor = "gainsboro";
    if (window.event.keyCode != 9)
        ImageButton_Pressed(window.event.srcElement.name);
}

function cmdImageButton2_onkeyup() {
    window.event.srcElement.style.borderColor = "";
    window.event.srcElement.style.borderStyle = "solid";
    window.event.srcElement.style.backgroundColor = "gainsboro";
    if (window.event.keyCode != 9)
        ImageButton2_Pressed(window.event.srcElement.name);
}

function ImageButton_Pressed(ID) {
    if (frmUpdate.txtRev.value == "")
        frmUpdate.txtRev.value = "2";
    else if (isNumeric(frmUpdate.txtRev.value))
        frmUpdate.txtRev.value = parseInt(frmUpdate.txtRev.value) + 1;
    else
        alert("Internal Rev must be an integer.");


    ResetStatus(1);
    frmUpdate.txtRev.focus();
}
function ImageButton2_Pressed(ID) {
    if (frmUpdate.txtPNRev.value == "")
        frmUpdate.txtPNRev.value = "2";
    else if (isNumeric(frmUpdate.txtPNRev.value))
        frmUpdate.txtPNRev.value = parseInt(frmUpdate.txtPNRev.value) + 1;
    else
        alert("Part Number Rev must be an integer.");

    ResetStatus(1);
    frmUpdate.txtPNRev.focus();
}

function cmdImageButton_onfocusout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
	window.event.srcElement.style.backgroundColor = "gainsboro";

}
function isNumeric(sText)
{
   var ValidChars = "-0123456789";
   var IsNumber=true;
   var Char;

 
   for (i = 0; i < sText.length && IsNumber == true; i++) 
      { 
      Char = sText.charAt(i); 
      if (ValidChars.indexOf(Char) == -1) 
         {
         IsNumber = false;
         }
      }
   return IsNumber;
   
   }
function cmdOK_onclick() {
    if (frmUpdate.txtRev.value == "")
		{
//	    frmUpdate.txtOSCode.value = txtOSCodeDisplay.value;
	//    frmUpdate.txtPrepStatus.value =  cboPrepStatus.options[cboPrepStatus.options.selectedIndex].value;
	  //  frmUpdate.txtRev.value = txtRev.value;
		frmUpdate.submit();
		}
else if (!isNumeric(frmUpdate.txtRev.value))
		{
		alert("Internal Rev number be an integer.");
		frmUpdate.txtRev.focus();
		}
else if (parseInt(frmUpdate.txtRev.value) < 1)
		{
		alert("Internal Rev number be a positive integer.");
		frmUpdate.txtRev.focus();
		}
else {
        //frmUpdate.txtOSCode.value = txtOSCodeDisplay.value;
        //frmUpdate.txtPrepStatus.value = cboPrepStatus.options[cboPrepStatus.options.selectedIndex].value;
        //frmUpdate.txtRev.value = txtRev.value;
		frmUpdate.submit();
		}	
}

function cmdCancel_onclick() {
	window.parent.close();
}

function PopulatePN() {
    if (PNCVR.innerText.indexOf("|R") == -1)
        frmUpdate.txtPartNumber.value = PNCVR.innerText;
    else 
        {
        var PNArray = PNCVR.innerText.split("|R")
        frmUpdate.txtPartNumber.value = PNArray[0];
        frmUpdate.txtPNRev.value = PNArray[1];
        }
}

function PopulatePN2(strID) {
    if (document.all("PNCVR" + strID).innerText.indexOf("|R") == -1)
        document.all("txtPN" + strID).value = document.all("PNCVR" + strID).innerText;
    else
        {
        var PNArray = document.all("PNCVR" + strID).innerText.split("|R");
        document.all("txtPN" + strID).value = PNArray[0];
        frmUpdate.txtPNRev.value = PNArray[1];
        }
}

function ResetStatus(TypeID) {
    var txtField = event.srcElement.id;
    var tagField = txtField + "tag";

    if (txtField == "")
        return;

    if (frmUpdate.cboPrepStatus.value == "1" && TypeID == 1) {
        lblStatusWarning.style.display = "";
        frmUpdate.cboPrepStatus.selectedIndex = 0;
    }
    else if (frmUpdate.cboPrepStatus.value == "1" && TypeID == 2 && document.getElementById(txtField).value != document.getElementById(tagField).value) {
        lblStatusWarning.style.display = "";
        frmUpdate.cboPrepStatus.selectedIndex = 0;
    }
}

function TurnOffWarning() {
    lblStatusWarning.style.display = "none";
    frmUpdate.txtOSCodetag.value = frmUpdate.txtOSCode.value;
    frmUpdate.txtRevtag.value = frmUpdate.txtRev.value;
    frmUpdate.txtPNRevtag.value = frmUpdate.txtPNRev.value;
    if (typeof(frmUpdate.txtPartNumbertag) != "undefined")
        frmUpdate.txtPartNumbertag.value = frmUpdate.txtPartNumber.value;

   // alert(typeof(document.all(1)));
}
//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<LINK href="../style/wizard%20style.css" type=text/css rel=stylesheet>

<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">


<%
	dim cn
	dim rs
	dim strRev
    dim strPNRev
    dim CurrentUserID
    dim CurrentUser
    dim CurrentUserPINGroup
    dim strGroupName
    dim strOSCode
    dim strPreinstallPrepStatus
    dim strDeliverable
    dim blnMultilanguage
    dim strPartNumber
    dim strDelName
    dim strDelVersion
    dim strDelRevision
    dim strDelPass
    dim blnLocalizationFound
    dim blnGlobalFound
    	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


    'Get User
	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))
	
	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
		
	
	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p
	
	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
	set cm=nothing
	
    CurrentUserPINGroup = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
        if rs("workgroupid")= 15 then
            CurrentUserPINGroup = 1
            strGroupName = "Houston "
        elseif rs("workgroupid")= 22 then
            CurrentUserPINGroup = 2
            strGroupName = "Taiwan "
        else
            CurrentUserPINGroup = 1
            strGroupName = "Houston "
	    end if
	end if
	rs.Close

    rs.open "spGetPreinstallDeliverableProperties " & clng(request("ID")),cn,adOpenStatic
    if rs.eof and rs.bof then
        strOSCode = "0"
        strPreinstallPrepStatus = ""
        strDeliverable = "&nbsp;"
        strDelName = ""
        blnMultilanguage = 0
        strPartNumber = ""
        strDelRevision = ""
        strDelPass = ""
    else
        strOSCode = trim(rs("OSCode") & "")
        strPreinstallPrepStatus= trim(rs("PreinstallPrepStatus") & "")
        strDelName = trim(rs("Name") & "")
        strDelVersion = trim(rs("version")  & "")
        strDelRevision = trim(rs("revision")  & "")
        strDelPass = trim(rs("pass")  & "")
        strDeliverable = rs("Name") & " [" & rs("version") 
        if trim(rs("Revision") & "") <> "" then
            strDeliverable = strDeliverable & "," & rs("Revision")
        end if
        if trim(rs("Pass") & "") <> "" then
            strDeliverable = strDeliverable & "," & rs("Pass")
        end if
        strDeliverable = strDeliverable & "]"
        blnMultilanguage = rs("Multilanguage")
        strPartNumber = rs("PartNumber")
    end if	
    rs.Close

	rs.Open "spGetInternalRev " & clng(request("ID")) & "," & clng(CurrentUserPINGroup),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strRev = ""
        strPNRev = ""
	else
		strRev = trim(rs("PreinstallInternalRev") & "")
	    strPNRev = trim(rs("PreinstallPNRev") & "")
    end if
    if trim(strRev) = "" then
        strRev = "1"
    end if
    if trim(strPNRev) = "" then
        strPNRev = "1"
    end if
    rs.close 


    'Check for Global Deliverable in Conveyor
    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
    cm.CommandType = 4
	cm.CommandText = "spGetConveyorPartNumber"
	
    Set p = cm.CreateParameter("@Delivname", 200, &H0001, 240)
	p.Value = left(strDelName,240)
	cm.Parameters.Append p
    	
	Set p = cm.CreateParameter("@DelivVersion", 200, &H0001, 120)
	p.Value = left(strDelVersion,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DelivOSCode", 200, &H0001, 120)
	p.Value = left(strOSCode,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DelivLangCode", 200, &H0001, 120)
	p.Value = "XX"
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DelivRevision", 200, &H0001, 120)
	p.Value = left(strDelRevision,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DelivExtPass", 200, &H0001, 120)
	p.Value = left(strDelPass,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DelivIntPass", 200, &H0001, 120)
	if strRev = "" then
        p.Value = null
    else
        p.Value = left(strRev,120)
    end if
	cm.Parameters.Append p
	
	set rs = server.CreateObject("ADODB.Recordset")
    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
            	
	set cm=nothing

    if rs.eof and rs.bof then
        blnGlobalFound = false
    else
        blnGlobalFound = true
    end if

    rs.close

    rs.open "spGetSelectedLanguages "  & clng(request("ID")),cn,adOpenForwardOnly
	strPartRow=""
    strLangIDList = ""
    blnLocalizationFound = false
    do while not rs.EOF
    	if rs("ID") & "" <> "58" then
            strLangIDList = strLangIDList & "," & rs("id")
    		strPartRow = strPartRow & "<TR ID=PNRow" & trim(rs("ID")) & "><TD nowrap>" & rs("cvrcode") & " - " & rs("Name") & "</TD><TD><INPUT style=""width:100%"" type=""text"" onchange=""javascript:ResetStatus(1);"" onmouseup=""javascript:ResetStatus(2);"" onkeyup=""javascript:ResetStatus(2);"" id=""txtPN" & rs("ID") & """ name=""txtPN" & rs("ID") & """ value=""" & rs("PartNumber") & """><input id=""txtPN" & rs("ID") & "tag"" type=""hidden"" value=""" & rs("PartNumber") & """/></TD>"

	        set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
            cm.CommandType = 4
	        cm.CommandText = "spGetConveyorPartNumber"
	
            Set p = cm.CreateParameter("@Delivname", 200, &H0001, 240)
	        p.Value = left(strDelName,240)
	        cm.Parameters.Append p
    	
	        Set p = cm.CreateParameter("@DelivVersion", 200, &H0001, 120)
	        p.Value = left(strDelVersion,120)
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DelivOSCode", 200, &H0001, 120)
	        p.Value = left(strOSCode,120)
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DelivLangCode", 200, &H0001, 120)
	        p.Value = left(rs("cvrcode"),120)
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DelivRevision", 200, &H0001, 120)
	        p.Value = left(strDelRevision,120)
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DelivExtPass", 200, &H0001, 120)
	        p.Value = left(strDelPass,120)
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@DelivIntPass", 200, &H0001, 120)
	        if strRev = "" then
                p.Value = null
            else
                p.Value = left(strRev,120)
            end if
	        cm.Parameters.Append p
	
	        set rs2 = server.CreateObject("ADODB.Recordset")
    	    rs2.CursorType = adOpenForwardOnly
	        rs2.LockType=AdLockReadOnly
	        Set rs2 = cm.Execute 
            	
	        set cm=nothing

            if not (rs2.eof and rs2.bof) then
               ' if trim(rs("PartNumber") & "") = "" then
                    strPartRow = strPartRow & "<td>" & "<a ID=PNCVR" & trim(rs("ID")) & " href=""javascript: PopulatePN2(" & rs("ID") & ");"">" & rs2("PartNumber") & "|R" & rs2("PNRevision") & "</a></td>"
               ' else
               '     strPartRow = strPartRow & "<td>" & rs2("PartNumber") & "|R" & rs2("PNRevision") & "</td>"
               ' end if
               if trim(rs2("PartNumber") & "") <> "" then
                    blnLocalizationFound = true
               end if
            end if
            rs2.Close
            
            strPartRow = strPartRow & "</tr>"
        end if    		
		rs.MoveNext
	loop
	rs.Close

    if strLangIDList <> "" then
        strLangIDList = mid(strLangIDList,2)
    end if


%>
    <font size=3 face=verdana><b>Preinstall Properties</b></font><BR>
<%
   
    if blnLocalizationFound and blnMultilanguage = 1 then
        response.write "<BR><font face=verdana color=red>Warning: Deliverable is marked as global(XX) in Excalibur and localized in Conveyor.</font><BR><BR>"
    elseif blnGlobalFound and (blnMultilanguage <> 1) then
        response.write "<BR><font face=verdana color=red>Warning: Deliverable is marked as localized in Excalibur and global(XX) in Conveyor.</font><BR><BR>"
    end if

%>
    <form id="frmUpdate" method="post" action="PreinstallPropertiesSave.asp">
<!--<input id="Submit1" type="submit" value="submit" />-->
	<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr>
            <td  width=120 valign=top><b>Deliverable:</b></td>
            <td><%=strDeliverable%></td>
        </tr>
     <TR><TD width=120><b>Prep&nbsp;Status:&nbsp;&nbsp;</b></TD>
        <TD width="100%">
            <select style="width:120px" id="cboPrepStatus" name="cboPrepStatus" onchange="javascript:TurnOffWarning()">
                <%if strPreinstallPrepStatus="0" then%>
                    <option selected value=0></option>
                <%else%>
                    <option value=0></option>
                <%end if %>
                <%if strPreinstallPrepStatus="1" then%>
                    <option value=1 selected>Complete</option>
                <%else%>
                    <option value=1>Complete</option>
                <%end if %>
                <%if strPreinstallPrepStatus="2" then%>
                    <option value=2 selected>Not Required</option>
                <%else%>
                    <option value=2>Not Required</option>
                <%end if %>
            </select>
            <label id=lblStatusWarning style="color:Red;display:none">Prep Status reset.</label>
        </TD>
     </TR>
     <TR><TD width=120><b>OS&nbsp;Code:&nbsp;</b></TD>
        <TD width="100%"><input id="txtOSCode" name="txtOSCode" onchange="javascript:ResetStatus(1);" onmouseup="javascript:ResetStatus(2);" onkeyup="javascript:ResetStatus(2);" style="width:120px" type="text" maxlength=5 value="<%=strOSCode%>"><input id="txtOSCodetag" type="hidden" value="<%=strOSCode%>" /></TD>
     </TR>

        <TR><TD><b>Internal&nbsp;Rev:&nbsp;<font size=1 color=red>*</font>&nbsp;&nbsp;</b></TD><TD width="100%"><table cellpadding=0 cellspacing=0 border=0><tr><td><INPUT id=txtRev name=txtRev style="MARGIN-TOP: -10px; VERTICAL-ALIGN: middle; WIDTH: 58px; HEIGHT: 22px" onchange="javascript:ResetStatus(1);" onmouseup="javascript:ResetStatus(2);" onkeyup="javascript:ResetStatus(2);"
      size=7 value="<%=strRev%>">&nbsp;<input type="image" src="../images/PLUS2.GIF" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; MARGIN-TOP: 1px; BORDER-LEFT: thin outset; WIDTH: 18px; BORDER-BOTTOM: thin outset; TOP: 0px; HEIGHT: 18px; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()"             
      id="cmdRev" name ="cmdRev" title ="Increment Internal Rev" LANGUAGE="javascript" 
      onmouseover="return cmdImageButton_onmouseover()" onclick="return false;"
      onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" 
      onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" 
      onkeyup="return cmdImageButton_onkeyup()" 
      WIDTH="18" HEIGHT="18"></TD><td valign=middle nowrap>&nbsp;-&nbsp;<%=strGroupName%></TD></tr></table>
            <input id="txtRevtag" type="hidden" value="<%=strRev%>"/>
      </TD></TR>
        <TR><TD><b>Part&nbsp;Number&nbsp;Rev:&nbsp;<font size=1 color=red>*</font>&nbsp;&nbsp;</b></TD><TD width="100%"><table cellpadding=0 cellspacing=0  border=0><tr><td><INPUT id=txtPNRev name=txtPNRev style="MARGIN-TOP: -10px; VERTICAL-ALIGN: middle; WIDTH: 58px; HEIGHT: 22px"  onchange="javascript:ResetStatus(1);" onmouseup="javascript:ResetStatus(2);" onkeyup="javascript:ResetStatus(2);"
      size=7 value="<%=strPNRev%>">&nbsp;<input type="image" src="../images/PLUS2.GIF" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; MARGIN-TOP: 1px; BORDER-LEFT: thin outset; WIDTH: 18px; BORDER-BOTTOM: thin outset; TOP: 0px; HEIGHT: 18px; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()"             
      id="cmdPNRev" name ="cmdPNRev" title ="Increment Part Number Rev" LANGUAGE="javascript" 
      onmouseover="return cmdImageButton_onmouseover()" onclick="return false;"
      onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton2_onmouseup()" 
      onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" 
      onkeyup="return cmdImageButton2_onkeyup()" 
      WIDTH="18" HEIGHT="18"></TD><td valign=middle nowrap>&nbsp;-&nbsp;<%=strGroupName%></TD></tr></table>
      <input id="txtPNRevtag" type="hidden" value="<%=strPNRev%>"/>
      </TD></TR>

     <TR><TD width=120 valign=top><b>Part&nbsp;Number:&nbsp;</b></TD>
        <TD width="100%">
        <%if blnMultilanguage=1 then %>
            <input id="txtPartNumber" name="txtPartNumber" style="width:200px" type="text" onchange="javascript:ResetStatus(1);" onmouseup="javascript:ResetStatus(2);" onkeyup="javascript:ResetStatus(2);" value="<%=strPartNUmber%>">
            <input id="txtPartNumbertag" type="hidden" value="<%=strPartNUmber%>" />
            <%
                if trim(strPartNUmber) = "" and trim(strDelName) <> "" then
	                set cm = server.CreateObject("ADODB.Command")
	                Set cm.ActiveConnection = cn
                	cm.CommandType = 4
	                cm.CommandText = "spGetConveyorPartNumber"
	
                	Set p = cm.CreateParameter("@Delivname", 200, &H0001, 240)
	                p.Value = left(strDelName,240)
	                cm.Parameters.Append p
    	
	                Set p = cm.CreateParameter("@DelivVersion", 200, &H0001, 120)
	                p.Value = left(strDelVersion,120)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@DelivOSCode", 200, &H0001, 120)
                    p.Value = left(strOSCode,120)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@DelivLangCode", 200, &H0001, 120)
	                p.Value = left("XX",120)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@DelivRevision", 200, &H0001, 120)
	                p.Value = left(strDelRevision,120)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@DelivExtPass", 200, &H0001, 120)
	                p.Value = left(strDelPass,120)
	                cm.Parameters.Append p

	                Set p = cm.CreateParameter("@DelivIntPass", 200, &H0001, 120)
	                if strRev = "" then
                        p.Value = null
                    else
                        p.Value = left(strRev,120)
                    end if
	                cm.Parameters.Append p
	
    	            rs.CursorType = adOpenForwardOnly
	                rs.LockType=AdLockReadOnly
	                Set rs = cm.Execute 
            	
	                set cm=nothing

                    if not (rs.eof and rs.bof) then
                        response.write "<a ID=PNCVR href=""javascript: PopulatePN();"">" & rs("PartNumber") & "|R" & rs("PNRevision") & "</a>"
                    end if
                    rs.Close
                end if            
            %>
        <%else%>
            <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll; border-left: steelblue 1px solid; width: 100%; border-bottom: steelblue 1px solid;height: 200px; background-color: white" id="DIV1">
                <table id="TablePart" width="100%">
                    <thead>
                        <tr style="position: relative; top: expression(document.getElementById('DIV1').scrollTop-2);">
                            <td bgcolor="lightsteelblue" nowrap style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset">Language&nbsp;&nbsp;</td>
                            <td bgcolor="lightsteelblue" style="width:100%; border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset">Excalibur&nbsp;Part&nbsp;Number&nbsp;&nbsp;</td>
                            <td bgcolor="lightsteelblue" nowrap style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset">Conveyor&nbsp;Part&nbsp;Number&nbsp;</td>
                        </tr>
                    </thead>
                    <%=strPartRow%>
                </table>
            </div>
        <%end if%>
        
        </TD>
     </TR>
	</table>
<%
	set rs = nothing
	cn.Close
	set cn = nothing
%>
<INPUT style="display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT style="display:none" type="text" id=txtTeam name=txtTeam value="<%=CurrentUserPINGroup%>">
<input style="display:none" id="txtMultiLanguage" name="txtMultiLanguage" type="text" value="<%=blnMultilanguage%>">
<input style="display:none" id="txtLangIDList" name="txtLangIDList" type="text" value="<%=strLangIDList%>">

</form>

</BODY>
</HTML>
