<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<HEAD>
<TITLE>Preinstall Properties</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.close();
        }
    }
}

function cmdOK_onclick() {
	frmProperties.submit();
}



function window_onload() {
    if (txtWorkgroupID.value =="22")
	    frmProperties.txtPartNumber.focus();
}

//-->
</SCRIPT>
</HEAD>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">

<%
	if request("ProductID") = "" or request("VersionID") = "" then
		Response.Write "Not enough information supplied to complete this action."
	else
		dim cn 
		dim rs
		dim cm
		dim p
        dim strGroupID
		dim blnSuccess
		dim strProdName
		dim strDelName
        dim strDelRev
        dim strProdRev
        dim strSkipRev
        dim strSkipChecked
        dim strDevCenter

        strDelRev = ""
        strProdRev = ""
        strSkipRev = ""
        strSkipChecked = ""
        strDevCenter = ""


		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
	
		set rs = server.CreateObject("ADODB.recordset")
  
  
	  	'Get User
		dim CurrentDomain
		dim CurrentUserPartner
		CurrentUser = lcase(Session("LoggedInUser"))
	
		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		set rs = server.CreateObject("ADODB.recordset")
	
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

	  	if (rs.EOF and rs.BOF) then
			set rs = nothing
        	set cn=nothing
        	Response.Redirect "../NoAccess.asp?Level=1"
        else
                CurrentUserPartner = rs("PartnerID")
                CurrentWorkgroupID = trim(rs("WorkgroupID") & "")
        end if 
        rs.Close


  
		blnSuccess = true
		rs.open "spGetProductVersionName " & request("ProductID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "Not enough information supplied to complete this action."
			blnSuccess = false
		else
			strProdname = rs("Name") & ""
            strDevCenter = trim(rs("DevCenter") & "")
			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					set rs = nothing
					set cn=nothing
				
					Response.Redirect "../NoAccess.asp?Level=1"
				end if
			end if
			
		end if
		rs.Close
		
		if blnSuccess then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetDeliverableVersionProperties"
			
	
			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p
		
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
		
			'rs.open "spGetDeliverableVersionProperties " & request("VersionID"),cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				Response.Write "Not enough information supplied to complete this action."
				blnSuccess = false
			else
				strDelname = rs("Name") & " " & rs("Version")
				if rs("Revision") <> "" then
					strDelName = strDelName & "," & rs("Revision")
				end if
				if rs("Pass") <> "" then
					strDelName = strDelName & "," & rs("Pass")
				end if

                if strDevCenter = "2" then
                    strDelRev = trim(rs("PreinstallInternalRevTDC") & "")
                else
                    strDelRev = trim(rs("PreinstallInternalRev") & "")
                end if
			end if
			rs.Close
		end if
			
		if blnSuccess then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetPartNumber"
			

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			p.Value = request("ProductID")
			cm.Parameters.Append p
		

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
		
			'rs.open "spGetPartNumber " & request("VersionID") & "," & request("ProductID"),cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				Response.Write "Not enough information supplied to complete this action."
				blnSuccess = false
			else
				strPartNumber = rs("PartNumber") & ""
				strInImage = rs("InImage") & ""
				strInImage = replace(replace(strInImage,"True","checked"),"False","")
				strInPINImage = rs("InPINImage") & ""
				strInPINImage = replace(replace(strInPINImage,"True","checked"),"False","")
			    strProdRev = trim(rs("PreinstallInternalRev") & "")
			    strSkipRev = trim(rs("PreinstallInternalRevSkipped") & "")
            end if
			rs.Close
		end if			

        if strDelRev = "" then
            strDelRev = "1"
        end if
        if strProdRev = "" then
            strProdRev = "1"
        end if

        if strDelRev = strSkipRev and strProdRev <> strDelRev then
            strSkipChecked = "checked"
        end if
		
		if blnSuccess then

%>
<form ID=frmProperties action="PIPropertiesSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<font size=4 face=verdana><b>Update Preinstall Properties</b></font><BR>
<table border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr>
	<td valign=top width=20><FONT face=verdana size=2><STRONG>Deliverable:</STRONG></FONT>&nbsp;</td>
	<td><font size=2 face=verdana><%=strDelName%></font></td>	
  </tr>
  <tr>
	<td width=20><FONT face=verdana size=2><STRONG>Product:</STRONG></FONT>&nbsp;</td>
	<td><font size=2 face=verdana><%=strProdName%></font></td>	
  </tr>
  <%if CurrentWorkgroupID = "22" then%>
    <tr>
  <%else%>
    <tr style="display:none">
  <%end if%>
	<td width=20><FONT face=verdana size=2><STRONG>Part&nbsp;Number:</STRONG></FONT>&nbsp;</td>
	<td><INPUT style="WIDTH: 100%" id=txtPartNumber name=txtPartNumber maxlength=50 value="<%=strPartNumber%>"></td>	
  </tr>
  <tr>
	<td width=20><FONT face=verdana size=2><STRONG>Status:</STRONG></FONT>&nbsp;</td>
	<td><INPUT <%=strInImage%> type="checkbox" id=chkInImage name=chkInImage disabled><font face=verdana size=2>In Image</font><INPUT <%=strInImage%> style="Display:none" type="checkbox" id=chkInImageTag name=chkInImageTag>&nbsp&nbsp
	</td>	
  </tr>
  <tr>
	<td valign=top width=20><FONT face=verdana size=2><STRONG>Internal&nbsp;Rev:</STRONG></FONT>&nbsp;</td>
	<%if strProdRev <> strDelRev then%>
        <td><b>Deliverable:</b>&nbsp;<%=strDelRev%>&nbsp;&nbsp;&nbsp;<b>Product:</b>&nbsp;<%=strProdRev%><br>
            <input <%=strSkipChecked%> id="chkSkip" type="checkbox" name="chkSkip" value="<%=strDelRev%>"> Skip this Rev on this product
        </td>	
    <%else%>
        <td><%=strProdRev%><input style="display:none" id="chkSkip" type="checkbox" name="chkSkip" value=""></td>	
    <%end if%>
  </tr>
</table>
<table width="400" border=0>
  <tr><TD align=right>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')">
  </TD></tr>
</table>
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
</form>
<INPUT type="hidden" id=txtWorkgroupID value="<%=CurrentWorkgroupID%>">

</body>

<%	
		end if
		set rs= nothing
		set cn = nothing
	end if
%>

</HTML>
