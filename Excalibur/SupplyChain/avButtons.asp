<%@ Language=VBScript %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/lib_debug.inc" --> 
<%
Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsPC : m_IsPC = Request("IsPC")
Dim m_IsMarketingUser
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_IsPulsarSystemAdmin

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity

	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If

	m_IsPulsarSystemAdmin = Security.IsPulsarSystemAdmin()

	Set Security = Nothing
'##############################################################################	
   
    Dim MarketingProductCount : MarketingProductCount = 0
    Dim CurrentUser : CurrentUser = lcase(Session("LoggedInUser"))
    Dim CurrentDomain
    If instr(CurrentUser,"\") > 0 Then
        CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
        CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
    End If
    Dim rsMarketingRole, dwMarketingRole, cnMarketingRole, cmdMarketingRole
    Set rsMarketingRole = Server.CreateObject("ADODB.RecordSet")
    Set dwMarketingRole = New DataWrapper
    Set cnMarketingRole = dwMarketingRole.CreateConnection("PDPIMS_ConnectionString")

    Set cmdMarketingRole = dwMarketingRole.CreateCommAndSP(cnMarketingRole, "spGetUserInfo")
    dwMarketingRole.CreateParameter cmdMarketingRole, "@UserName", adVarchar, adParamInput, 80, CurrentUser
    dwMarketingRole.CreateParameter cmdMarketingRole, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
    Set rsMarketingRole = dwMarketingRole.ExecuteCommandReturnRS(cmdMarketingRole)

    If not (rsMarketingRole.EOF And rsMarketingRole.BOF) Then
	    'add the permission from the Users and Roles to the Pulsar products
        If Not m_IsMarketingUser Then
		    MarketingProductCount = rsMarketingRole("MarketingProductCount")
            if MarketingProductCount > 0 then
                m_IsMarketingUser = True
            end if
	    End If
    End If
    rsMarketingRole.Close
    Set rsMarketingRole = Nothing
	Set cmdMarketingRole = Nothing

If Request("AVID") <> Request("hidAVID") And Request("hidAVID") <> "" Then
	Response.Redirect "avButtons.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request.Form("hidAVID")
End If

Dim rs, dw, cn, cmd
Dim iPrev, iCur, iNext
Dim sStatus 

If LCase(Request("MODE")) <> "add" Then
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail_Pulsar")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvCategoryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
	dw.CreateParameter cmd, "@p_GpgDescription", adVarchar, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_UPC", adChar, adParamInput, 12, ""
	dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
	dw.CreateParameter cmd, "@p_KMAT", adChar, adParamInput, 6, ""
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	If Trim(Request("AVID")) <> "" Then
		Do Until rs.EOF Or iCur = CLng(Trim(Request("AVID")))
			iPrev = iCur
			iCur = rs("AvDetailID")            
			rs.MoveNext
		Loop

		If Not rs.EOF Then
			iNext = rs("AvDetailID")
		End If
	End If
	
	rs.Close
End If	

%>
<html>
<head>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<script type="text/javascript" src="../includes/Date.asp"></script>
<script src="../includes/client/jquery.min.js" type="text/javascript"></script>
<script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
<% If m_IsMarketingUser and not m_IsPC and not m_IsPulsarSystemAdmin and LCase(Request("MODE")) <> "add" Then %>
<SCRIPT type="text/javascript">
<!--
    function VerifyStatus() {
        if (window.parent.frames["UpperWindow"].frmMain.txtIsMarketingScreen.value == "1") {
            // remove FCS from all areas - task 20243
            with (window.parent.frames["UpperWindow"].frmMain) {
                if (txtRTPDate.value != "") {
                    if (!(isDate_mmddyyyy(txtRTPDate.value))) {
                        alert("RTP/MR Date must be in mm/dd/yyyy format.");
                        return false;
                    }                    
                }

                if (txtMarketingDiscDate.value != "") {
                    if (!(isDate_mmddyyyy(txtMarketingDiscDate.value))) {
                        alert("End of Manufacturing (EM) Date must be in mm/dd/yyyy format.");
                        return false;
                    }                    
                }
            }
        }
         if (typeof (window.parent.frames["UpperWindow"].frmMain.chkRelease) != "undefined") {
                var strRelease = "";
                var strReleaseFormatted = "";
                if (typeof (window.parent.frames["UpperWindow"].frmMain.chkRelease.length) == "undefined") {
                    if (window.parent.frames["UpperWindow"].frmMain.chkRelease.checked == true)
                        strRelease = strRelease + ", " + window.parent.frames["UpperWindow"].frmMain.chkRelease.ReleaseID;
                }
                else {
                    for (i = 0; i < window.parent.frames["UpperWindow"].frmMain.chkRelease.length; i++)
                        if (window.parent.frames["UpperWindow"].frmMain.chkRelease[i].checked)
                            strRelease = strRelease + ", " + window.parent.frames["UpperWindow"].frmMain.chkRelease[i].value;
                }
        
                strReleaseFormatted = strRelease.substring(2);
                window.parent.frames["UpperWindow"].frmMain.txtReleaseList.value = strReleaseFormatted;
          }

        return true;
    }
//-->
</SCRIPT>
<% Else %>
<SCRIPT type="text/javascript">
<!--
function VerifyStatus()
{// remove FCS from all areas - task 20243
    if (window.parent.frames["UpperWindow"].frmMain.txtIsMarketingScreen.value == "1") {
        with (window.parent.frames["UpperWindow"].frmMain) {
            if (txtRTPDate.value != "") {
                if (!(isDate_mmddyyyy(txtRTPDate.value))) {
                    alert("RTP/MR Date must be in mm/dd/yyyy format.");
                    return false;
                }
            }

            if (txtMarketingDiscDate.value != "") {
                if (!(isDate_mmddyyyy(txtMarketingDiscDate.value))) {
                    alert("End of Manufacturing (EM) Date must be in mm/dd/yyyy format.");
                    return false;
                }
            }
        }
        return true;
    }
    with (window.parent.frames["UpperWindow"].frmMain) {
        //check for date formate
        if (txtRTPDate.value != "") {
            if (!(isDate_mmddyyyy(txtRTPDate.value))) {
                alert("RTP/MR Date must be in mm/dd/yyyy format.");
                return false;
            }
        }

        if (txtMarketingDiscDate.value != "") {
            if (!(isDate_mmddyyyy(txtMarketingDiscDate.value))) {
                alert("End of Manufacturing (EM) Date must be in mm/dd/yyyy format.");
                return false;
            }
       }
         // Check for releases if it is Supply Chain screen - task 16504
        var strRelease = "";
        var strReleaseFormatted = "";
        if (typeof (window.parent.frames["UpperWindow"].frmMain.chkRelease.length) == "undefined") {
            if (window.parent.frames["UpperWindow"].frmMain.chkRelease.checked == true)
                strRelease = strRelease + ", " + window.parent.frames["UpperWindow"].frmMain.chkRelease.ReleaseID;
        }
        else {
            for (i = 0; i < window.parent.frames["UpperWindow"].frmMain.chkRelease.length; i++)
                if (window.parent.frames["UpperWindow"].frmMain.chkRelease[i].checked)
                    strRelease = strRelease + ", " + window.parent.frames["UpperWindow"].frmMain.chkRelease[i].value;
        }
        
        if (strRelease == "")
        {
            window.alert("Release is required.");
            return false;
        }

        strReleaseFormatted = strRelease.substring(2);
        window.parent.frames["UpperWindow"].frmMain.txtReleaseList.value = strReleaseFormatted;
                
        if (txtFeatureID.value == "") {
            alert("Please select a Feature to create an AV");
            return false;
        }

        // only do check if it is Add New AV
        if (hidAVID.value == "") {
            var NumChecked = 0;
            if (typeof (window.parent.frames["UpperWindow"].frmMain.chkBrand.length) == "undefined") {
                if (window.parent.frames["UpperWindow"].frmMain.chkBrand.checked == true)
                    NumChecked += 1;
            }
            else {
                for (i = 0; i < window.parent.frames["UpperWindow"].frmMain.chkBrand.length; i++)
                    if (window.parent.frames["UpperWindow"].frmMain.chkBrand[i].checked)
                        NumChecked += 1;
            }            
            if (NumChecked == 0) {
                alert("Please select at least one Brand");
                return false;
            }
        }

        if (cboCategory.value == 0) {
        	alert('Please choose a valid SCM Category');
        	return false;
        }

        if (cboProductLine.value == "") {
        	alert("Product Line is required");
        	return false;
        }

        if (AvNo.value == "" && txtAvGpgDescription.value == "")
        {
            alert("AV Number or GPG Description is required");
            return false;
        }
	
        if ((hidParentID.value == 0 && AvNo.value != "") || (hidAVID.value == "" && AvNo.value != "")) {
            //check if NON-Localized AV contains a # char
            var sAvNo = AvNo.value;
            var index = sAvNo.indexOf("#");

            if (index > -1 && hidAVID.value == "") {
                alert("You are creating new NON-Localized AV.  Please remove the # character and any characters after the #.");
                return false;
            }
            else if (index > -1) {
                alert("This AV is a NON-Localized AV.  Please remove the # character and any characters after the #.");
                return false;
            }
        }
        else if (AvNo.value != "" && hidParentID.value > 0 && hidBaseAvNo.value != "") {
            var sAvNo = AvNo.value;
            var index = sAvNo.indexOf("#"); //right(AvNo, AvNo.length - AvNo.instr("#"))
            if (index != -1) {
                var BaseAvNo = sAvNo.substring(0, index);
                if (BaseAvNo.toLowerCase() != hidBaseAvNo.value.toLowerCase()) {
                    if (!confirm("The base part '" + BaseAvNo + "' of this localized AV is different than the parent AV '" + hidBaseAvNo.value + "'.  Ex: if localized AV# is XXXXX#YYY then the Base is XXXXX.  Click OK to change the base part for this and all other localizations to this new base part number. Click Cancel to return to AV Details."))
                        return false;
                    else 
                    {
                        $.ajax({
                            url: "GetBaseAvNo?AvDetailID=" + hidParentID.value,
                            method: "GET",
                            success: function (returnData) {
                                alert(returnData);
                            },
                            cache: false
                        });
                    }
                }
            }
            else {
                alert("The Av# of this localized AV is not in the correct format Ex: AV# is XXXXX#YYY");
                return false;
            }
        }
    
        if (AvNo.value != "" && (AvNo.value == hidAvNo.value) && hidMode.value == "clone") //check for part number instead of the GPG description since same feature can be added multiple times to the SCM
		{
			alert("AvNo must be changed");
			return false;
	    }
        if (AvNo.value != "") {
            var strAVNo = TrimBlankSpace(AvNo.value);
            var firstChar = strAVNo.charAt(0);
            if (firstChar.match(/^[a-zA-Z0-9]+/))
                return true
            else {
                alert("The first characther of the AV part number is not an alphanumeric characther.\rReplace the first character of the AV part number with an alphanumeric characther.")
                return false;
            }
        }
		if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
		    if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "False") {
		        if (!validateTextAreaSize(txtMarketingDesc, 40, 'Marketing Description')) { return false; }
		        if (!validateTextAreaSize(txtConfigRules, 1600, 'Config Rules')) { return false; }
		        if (!validateTextAreaSize(txtManufacturingNotes, 1600, 'Manufacturing Notes')) { return false; }
		    }
		}
		
		if (!validateDateInput(txtGSEndDt, "GS End Date")){ return false; }
    }

	return true;
}
function TrimBlankSpace(x) {
    return x.replace(/^\s+|\s+$/gm, '');
}

//-->
</SCRIPT>
<% End If %>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        } else {
            parent.window.parent.ClosePropertiesDialog();
        }
}

function cmdOK_onclick() 
{
	var blnAll = true;
	var i;
    var sReturnValue;

	if (VerifyStatus()) {
	    //if (IsFromPulsarPlus()) {
	    //    if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
	    //        if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "True") {
	    //            window.parent.frames["UpperWindow"].SaveNameElements();
	    //        } else {
	    //            window.parent.frames["UpperWindow"].frmMain.strNameElements.value = "";
	    //        }
	    //    }
	    //    window.frmButtons.cmdOK.disabled = true;
	    //    window.frmButtons.cmdCancel.disabled = true;

	    //    window.parent.frames["UpperWindow"].frmMain.hidFunction.value = "save";
	    //    window.parent.frames["UpperWindow"].frmMain.submit();
	    //    ClosePulsarPlusPopup();
	    //} else {
	        if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted != null) {
	            if (window.parent.frames["UpperWindow"].frmMain.IsNameFormatted.value == "True") {
	                window.parent.frames["UpperWindow"].SaveNameElements();
	            } else {
	                window.parent.frames["UpperWindow"].frmMain.strNameElements.value = "";
	            }
	        }
	        window.frmButtons.cmdOK.disabled = true;
	        window.frmButtons.cmdCancel.disabled = true;

	        window.parent.frames["UpperWindow"].frmMain.hidFunction.value = "save";
	        window.parent.frames["UpperWindow"].frmMain.submit();

	    }
	//}
	return;
}

function cmdPrev_onclick()
{
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidPrev.value;
	frmButtons.hidAVID.value = frmButtons.hidPrev.value;
	cmdOK_onclick();
	frmButtons.submit();
}

function cmdNext_onclick()
{
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidNext.value;
	frmButtons.hidAVID.value = frmButtons.hidNext.value;
	cmdOK_onclick();
	frmButtons.submit();
}

function cmdClone_onclick()
{
	var strID;
	var productVersionID = frmButtons.PVID.value;
	var avDetailID = frmButtons.AVID.value;
	var productBrandID = frmButtons.BID.value;
	parent.window.parent.ClosePropertiesDialog();	
	parent.window.parent.ShowAvDetailsForClone(productVersionID, avDetailID, productBrandID);
	
}
function document_OnLoad()
{
	window.frmButtons.cmdOK.disabled = true;
	if (typeof(window.parent.frames["UpperWindow"].document.all["hidMode"]) == 'object')
	{
		if (window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'add' || window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'clone')
		{	
			window.frmButtons.cmdOK.disabled = false;
		}
		else if ((window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'edit') || (window.parent.frames["UpperWindow"].document.all["hidMode"].value.toLowerCase() == 'editdates'))
		{
			if(window.parent.frames["UpperWindow"].document.all["hidStatus"].value != "O" && window.parent.frames["UpperWindow"].document.all["hidStatus"].value != "D")
			{
				window.frmButtons.cmdOK.disabled = false;
			}
		}
	}
	
	if (typeof(window.frmButtons.cmdNext) == 'object')
	{
		if (window.frmButtons.hidNext.value == '')
			window.frmButtons.cmdNext.disabled = true;
	}	
	if (typeof(window.frmButtons.cmdPrev) == 'object')
	{
		if (window.frmButtons.hidPrev.value == '')
			window.frmButtons.cmdPrev.disabled = true;
	}	
}
//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmButtons"  action=avButtons.asp method=post>
<INPUT type="HIDDEN" id=PVID name=PVID value="<%= Request("PVID")%>">
<INPUT type="HIDDEN" id=AVID name=AVID value="<%= Request("AVID")%>">
<INPUT type="HIDDEN" id=BID name=BID value="<%= Request("BID")%>">
<INPUT type="HIDDEN" id=Mode name=Mode value="<%= Request("Mode")%>">
<INPUT type="HIDDEN" id=hidPrev name=hidPrev value="<%= iPrev%>">
<INPUT type="HIDDEN" id=hidNext name=hidNext value="<%= iNext%>">
<INPUT type="HIDDEN" id=hidAVID name=hidAVID value="<%= Request("AVID")%>">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" WIDTH=100%>
	<tr><TD width=33% align=left>
<% If LCase(Request("MODE")) <> "clone" And  m_EditModeOn AND Request("FromTodayPage") <> "1" AND NOT m_IsMarketingUser Then %>			
			<INPUT type="button" value="Clone" id=cmdClone name=cmdClone onclick="return cmdClone_onclick()"></TD>
<% Else %>
			&nbsp;
<% End If %>
		<TD width=33% align=center>
<% If LCase(Request("MODE")) <> "add" And LCase(Request("MODE")) <> "clone" And LCase(Request("MODE")) <> "editdates" AND Request("FromTodayPage") <> "1" Then %>			
			<INPUT type="button" accesskey="P" value="Prev" id=cmdPrev name=cmdPrev onclick="return cmdPrev_onclick()">
			<INPUT type="button" accesskey="N" value="Next" id=cmdNext name=cmdNext onclick="return cmdNext_onclick()">
<% Else %>
			&nbsp;
<% End If %>
		</TD>
		<TD width=33% align=right><INPUT accesskey="S" type="button" value="Save" id=cmdOK name=cmdOK onclick="return cmdOK_onclick()">
			<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>


</FORM>
</body>
</html>
