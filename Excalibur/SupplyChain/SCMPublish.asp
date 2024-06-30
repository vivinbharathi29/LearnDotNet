    <%@ Language=VBScript %>
    <% OPTION EXPLICIT %>
    <!-- #include file = "../includes/Security.asp" -->
    <!-- #include file="../includes/DataWrapper.asp" -->
    <!-- #include file="../includes/no-cache.asp" -->
    <!-- #include file="../includes/lib_debug.inc" -->
    <%
    dim strError, strTabName
    dim strModuleID, strPALCategoryID
    dim oErr, oSvr, nProductBrandID, Ors, nProductVersionID
    dim sSCMPublishDate
    dim sRevision, sPrevRevision, sprevSCMName   
    dim intTotalRevisions, intProductID, bdisableProductList
    dim nstandardversion, nNonStandarversion, nQualificationversion
    dim bFirstPjPublish
    dim bCycleAVlinked, bfirstTimeCycleAVPublish
    dim intXForSR, intXForNSR, intY, snewRev, strYZ, bnewformat
    dim strSCMName 
    Dim cn, dw, cmd, rs, cnString 
    dim bsnapshot, intSnapshotValue
    Dim Security, m_UserFullName, returnValue, newSCMID
    Dim sStandardpublishurl
    Dim bDesktop
    Dim AppRoot
    Dim sLatestSCMRevision,sPrevSCMRevision
    AppRoot = Session("ApplicationRoot")
	sStandardpublishurl = ""
	
	Set Security = New ExcaliburSecurity
	
	
	m_UserFullName = Security.CurrentUserFullName()

    strError = ""

    if Request.QueryString("ProductBrandID") <> "" then
	    nProductBrandID = clng(Request.QueryString("ProductBrandID"))
    end if
    if Request.QueryString("PVID") <> "" then
        nProductVersionID = clng(Request.QueryString("PVID"))
    end if
    
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    if len(Request.form("bsnapshot")) > 0 then
		intSnapshotValue = cint(Request.form("bsnapshot"))
		bsnapshot = true
	end if


        
	if bsnapshot then
		
        returnValue = 0
		select case intSnapshotValue
			case 1
				'Regular scheduled publish
				    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_Publish")
                    dw.CreateParameter cmd, "@ProductBrandID", adInteger, adParamInput, 8, nProductBrandID
                    dw.CreateParameter cmd, "@p_chrSCMName", adVarchar, adParamInput, 260, Request.Form("txtSCMName")
                    dw.CreateParameter cmd, "@p_chrUserName", adVarchar, adParamInput, 120, m_UserFullName
                    dw.CreateParameter cmd, "@p_chrRevision", adVarchar, adParamInput, 10,  Request.Form("txtStandarversion")
                    dw.CreateParameter cmd, "@p_intStandardRelease", adInteger, adParamInput, 8, 1
                    dw.CreateParameter cmd, "@p_chr_Reason", adVarchar, adParamInput, 256, Request.Form("txtReason")
                    dw.CreateParameter cmd, "@p_intSCMID", adInteger, adParamOutput, 8, ""
				    returnValue = dw.ExecuteNonQuery(cmd)
		
				    newSCMID = cmd("@p_intSCMID")
            
    
			case 2
				'non-standard publish is redirected to ipulsar page 		

		end select     
            if newSCMID = 0 then 'MODIFIED THE SP SO it will return 0 if erorr
                Response.Write("SCM Publish Failed.")
                Response.End
            else
		        sStandardpublishurl =  "/ipulsar/Reports/SCM/SCM_Report_DT.aspx?SCMID=" & newSCMID & "&BID=" & nProductBrandID & "&PVID=" & nProductVersionID & "&Publish=True"
	        end if
	end if
	'end of snapshot



    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_GetDefaultSCMName")
    dw.CreateParameter cmd, "@ProductBrandID", adInteger, adParamInput, 8, nProductBrandID

    Set rs = dw.ExecuteCommandReturnRS(cmd)

	
	if not (rs.EOF and rs.BOF) then
		strSCMName = rs("SCMName") 
        sRevision = rs("Revision")
        sPrevRevision = sRevision
        sSCMPublishDate =rs("SCMPublishDate")
        sprevSCMName = rs("PrevSCMName")
        if rs("bDesktop")="True" then
             bDesktop ="1"
        else 
            bDesktop =""
        end if
        sLatestSCMRevision = rs("LatestSCMRevision")
        sPrevSCMRevision = sLatestSCMRevision
	end if
	rs.Close
    bFirstPjPublish = false
    if sLatestSCMRevision = "" then
        bFirstPjPublish = true 
    end if
  
    'construct the new revision
    'detetct if the revision is in new format
				' for the legacy stuff, need to continuet the prv number
				if not instr(sRevision, ".") > 0 then
					sRevision = sRevision & ".0"
				end if				
				sRevision = sRevision & ".0"	
        
                if not instr(sLatestSCMRevision, ".") > 0 then
					sLatestSCMRevision = sLatestSCMRevision & ".0"
				end if				
				sLatestSCMRevision = sLatestSCMRevision & ".0"				
				
				'if sLatestSCMRevision = "" or isnull(sLatestSCMRevision) then
				'	bnewformat =false
				'else
				if instr(sLatestSCMRevision, ".") > 0 then
					strYZ = mid(sLatestSCMRevision, instr(sLatestSCMRevision, ".") + 1 )
					if instr(strYZ, ".") > 0 then
						bnewformat =true			
					end if 
				end if	
				'end if 		
				
				if not bnewformat then
					intXForSR = 0
                    intXForNSR = 0
					intY = 0				
				else
					if left(sRevision,instr(sRevision, ".") - 1) = "" then
						intXForSR =0
					else
						intXForSR=left(sRevision,instr(sRevision, ".")-1)
					end if

                    if left(sLatestSCMRevision,instr(sLatestSCMRevision, ".") - 1) = "" then
						intXForNSR =0
					else
						intXForNSR=left(sLatestSCMRevision,instr(sLatestSCMRevision, ".")-1)
					end if
					
					intY =left(strYZ,instr(strYZ, ".")-1)			
					
					
				end if 
				if isnumeric(intXForSR) then						
					nstandardversion = cstr(formatNumber(cint(intXForSR) + 1, 0)) & "." & cstr(0) 
				end if 
				
				if isnumeric(intY) then						
					nNonStandarversion = cstr(intXForNSR) & "." & cstr(formatNumber(cint(intY) + 1, 0)) 
				end if 
				
        strSCMName = strSCMName & " " & sSCMPublishDate & " " & nstandardversion
        strSCMName = Replace(strSCMName,"/","_")

    
%>
<html>
<head>
<title>Enter new SCM Publish Date</title>
<style>
  html, body, button, div, input, select, td, fieldset { font-family: MS Shell Dlg; font-size: 8pt; };
</style>

<script type="text/javascript" LANGUAGE="javascript">
function rdoStndard1_onclick() {
    thisform.rdoStndard2.checked = false;
    var strSCMName = document.getElementById("txtSCMName").value;
    document.getElementById("txtRevision").value = document.getElementById("txtStandarversion").value;
    document.getElementById("txtSCMName").value = strSCMName.replace(document.getElementById("txtNonStandarversion").value, document.getElementById("txtStandarversion").value)
    thisform.txtReason.value ="Standard/Scheduled Release";
}
function rdoStndard2_onclick() {
  thisform.rdoStndard1.checked=false;
  var strSCMName = document.getElementById("txtSCMName").value;
  document.getElementById("txtRevision").value = document.getElementById("txtNonStandarversion").value
  document.getElementById("txtSCMName").value = strSCMName.replace(document.getElementById("txtStandarversion").value, document.getElementById("txtNonStandarversion").value)
  thisform.txtReason.value = "Non-Scheduled/Out of Cycle Release";
  
}



function btnOKClick() {
	
    //disable the OK button once it is clicked
    document.getElementById("btnOK").disabled = true;

	if (document.getElementById("txtReason").value == "")	{	
		alert("Please enter a reason");
		return false;	
	}
	if (document.getElementById("txtReason").value.length > 256) {
	    alert("The reason field is limited to 256 characters");
	    return false;
	}
	if (document.getElementById("txtSCMName").value == "") {
	    alert("Please enter SCM Name");
	    return false;
	}	
    //the textbox is alreadys et to be amx 260, just add code here to make sure
	if (document.getElementById("txtSCMName").value.length > 260) {
	    alert("The SCM Name is limited to 260 characters");
	    return false;
	}
    //Do not allow invalid characters in the file name:   \ / ? : * " > < |
	var strSCMName = document.getElementById("txtSCMName").value;
	if (strSCMName.indexOf("\"") > -1
        || strSCMName.indexOf(">") > -1
        || strSCMName.indexOf("*") > -1
        || strSCMName.indexOf("\\") > -1
        || strSCMName.indexOf("|") > -1
        || strSCMName.indexOf("/") > -1
        || strSCMName.indexOf("<") > -1
        || strSCMName.indexOf("?") > -1
        || strSCMName.indexOf(":") > -1
        ) {
	    alert("SCM Name does not allow these characters:  \\ / ? : * \" > < |");
	    return false;
	}
	
	
	if (thisform.rdoStndard1.checked) {
		document.thisform.txtStandardRelease.value = "1";
		document.thisform.action = 'SCMPublish.asp?ProductBrandID=<%=cstr(nProductBrandID)%>&PVID=<%=cstr(nproductversionID)%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>';
		document.thisform.bsnapshot.value = "1";
		document.thisform.submit();
	} else {  //non-selected publish
		document.thisform.txtStandardRelease.value = "0";			
		document.thisform.bsnapshot.value = "2";
		var isDesktop = document.getElementById("txtBDesktop").value;
		var strSCMName = document.thisform.txtSCMName.value;
		var strNonStandarversion = document.thisform.txtNonStandarversion.value;
		var strReason = document.thisform.txtReason.value
		var pulsarplusDivId = document.getElementById('hdnTabName');
		if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
		    //parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);
		    // For Closing current popup
		    parent.window.parent.closeExternalPopup();
		    parent.window.parent.document.getElementById("txtNonstandardpublishName").value = strSCMName;
		    parent.window.parent.document.getElementById("txtNonStandarversion").value = strNonStandarversion;
		    parent.window.parent.document.getElementById("txtReason").value = strReason;
		    var productBrandID = "<%=cstr(nProductBrandID)%>";
		    var pvid = "<%=cstr(nproductversionID)%>";
		    parent.window.parent.ShowAVhistoryDialog(productBrandID, pvid, isDesktop);
		}
		else {
		    window.parent.ClosePropertiesDialog();
		    window.parent.document.getElementById("txtNonstandardpublishName").value = strSCMName;
		    window.parent.document.getElementById("txtNonStandarversion").value = strNonStandarversion;
		    window.parent.document.getElementById("txtReason").value = strReason;
		    window.parent.ShowAVhistoryDialog('/ipulsar/SCM/SCM_AVHistoriesforNonStandardPublish.aspx?ProductBrandID=<%=cstr(nProductBrandID)%>&PVID=<%=cstr(nproductversionID)%>&BDesktop="' + isDesktop +'"&pulsarplusDivId=<%=Request("pulsarplusDivId")%>', "Publish SCM - Non-Scheduled Publish");

		}	
			
	}

}


function Cancel_onclick() {
    var pulsarplusDivId = document.getElementById('hdnTabName');
    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    }
    else {
        window.parent.ClosePropertiesDialog();
    }
    
}

function body_onload()
{
    var strStandardpublishurl = document.getElementById("txtStandardpublishurl").value;
    if (strStandardpublishurl != "")
    {
        var pulsarplusDivId = document.getElementById('hdnTabName');
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Closing current popup
            //parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);
            parent.window.parent.ShowRollbackScmPublish();
            parent.window.parent.closeExternalPopup();
           window.open(strStandardpublishurl,"_blank");
        }
        else {
            window.parent.ClosePropertiesDialog();
            window.location.href = strStandardpublishurl;
        }
        

    }
    
    
}
</SCRIPT>
</HEAD>
<BODY id=bdy onload="return body_onload();" style="background: threedface; color: windowtext; margin: 10px; BORDER-STYLE: none" scroll=no>
<form name=thisform method=post>
<%
if strError <> "" then
	Response.Write strError
else
	%>
	<table border=0 cellspacing=0 cellpadding=0>
        <tr>
			<TD style="width:18%;vertical-align:top;">SCM Name<font color=red>*</font><br></TD>
			<TD style="width:72%">		
				<input type="text" id="txtSCMName" name="txtSCMName" style="width:98%" value="<%=strSCMName%>"  maxlength="260">
				<br><br>
			</TD>
            <td style="width:10%"><font color="orange">(260 max. characters)</font></td>
		</tr> 
		 <tr>
			<TD style="width:15%;vertical-align:top;">Previously Published SCM Name<br></TD>
			<TD style="width:75%">		
				<%=sPrevSCMName%>
				<br><br>
			</TD>
            <td style="width:10%"></td>
		</tr> 
		<tr>
			<TD><br>Previous Revision<br><br></TD>
			<TD><br /><%=sPrevRevision%> &nbsp&nbsp <font color="orange">(This number also includes the Program Matrix publishes)</font><br><br></TD>
             <td ></td>
		</tr>
		<tr>
			<TD><br>Previous SCM Revision<br><br></TD>
			<TD><br /><%=sPrevSCMRevision%><br><br></TD>
             <td ></td>
		</tr>
		<tr>
			<TD>New Revision<br></TD>
			<TD>		
				<input id="txtRevision" name="txtRevision" disabled="true"
                <%
				response.write " value=""" & nstandardversion & """ "
				%>
				size=10 maxlength=10  >
			</TD>
             <td style="width:100px;"></td>
		</tr>
		
		<tr>
			<TD colspan = 2><input type= radio ID = rdoStndard1 Name = rdoStndard1 value = 1 checked  onclick="return rdoStndard1_onclick();">&nbsp;Standard/Scheduled Release&nbsp;&nbsp;&nbsp;&nbsp;
				<input type= radio ID = rdoStndard2 Name = rdoStndard2 value = 0 <%if bFirstPjPublish then Response.Write " disabled "  %>onclick="return rdoStndard2_onclick();" >&nbsp;Non-Scheduled/Out of Cycle Release
				
				<br><br></TD>
             <td ></td>
		</tr> 
		<tr>
			<TD>Reason<font color=red>*</font><br></TD>
			<TD >		
				<textarea cols=60 rows=6 name=txtReason id=txtReason style='font-family:Arial;font-size:12px;'>Standard/Scheduled Release</textarea>
				<br><br>
			</TD>
             <td ><font color="orange">(256 max. characters)</font></td>
		</tr> 
	</table>
	
	<BUTTON style="width: 7em; height: 2.2em;" type=button id="btnOK" name="btnOK" onclick="btnOKClick()">OK</BUTTON> &nbsp; <BUTTON 
	style="width: 7em; height: 2.2em;" type=button onClick="Cancel_onclick();">Cancel</BUTTON>
	<%
end if
%>

<input type=hidden Name=txtOriginalPublishdate ID=txtOriginalPublishdate value="<%= sSCMPublishDate %>">
<input type=hidden Name=txtStandarversion ID=txtStandarversion value="<%=nstandardversion%>">
<input type=hidden Name=txtNonStandarversion ID=txtNonStandarversion value="<%=nNonStandarversion%>">
<input type=hidden Name=txtQualificationversion ID=txtQualificationversion value="<%=nQualificationversion%>">
<INPUT type="hidden" id=bsnapshot name=bsnapshot value="">
<INPUT type="hidden" id=txtBDesktop name=txtBDesktop value="<%=bDesktop%>" >
<INPUT type="hidden" id=txtStandardRelease name=txtStandardRelease value="">
<INPUT type="hidden" id=txtStandardpublishurl name=txtStandardpublishurl value="<%=sStandardpublishurl%>">
        <input type="hidden" id="hdnTabName" name="hdnTabName" value="<%= Request("pulsarplusDivId")%>" />
<p><font color=red>*</font>Required field
</form>
</BODY>
</HTML>
<%
'Set Publish Date to the last Wed. of the month unless the last Wed. is today
'or later. Then use the last Wed. of next month. 
function SetSCMPublishDate()
	dim strWeek, strDay, strNow
	strNow = Now
	strWeek = DateSerial(Year(strNow), Month(strNow)+1, 1)  'One month ahead
	strDay = strWeek - 1 - (WeekDay(strWeek)+2) mod 7 'Back a bit, answer as a Date
	'make sure the last Wed. isn't today or later in the month
	if strNow >= strDay then	'go to next month
		strWeek = DateSerial(Year(strNow), Month(strNow)+2, 1)  'Two months ahead
		strDay = strWeek - 1 - (WeekDay(strWeek)+2) mod 7 'Back a bit, answer as a Date
	end if
	
	SetSCMPublishDate = FormatDateTime(strDay, vbShortDate)
end function
%>
