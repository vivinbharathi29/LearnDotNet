<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/lib_debug.inc" --> 
<%
  Dim AppRoot
  AppRoot = Session("ApplicationRoot")

  Dim SuccessCount
  Dim ErrorCount
  Dim sFunction			: sFunction = Request.Form("hidFunction")
  Dim BusinessID
  '##############################################################################	
  '
  ' Create Security Object to get User Info
  '
	
	'm_EditModeOn = False
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	strUserID = Security.CurrentUserID()
	strUserName = Security.CurrentUserFullName()

	Set Security = Nothing
  '##############################################################################	 
  
  function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i

	    badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
  end function 

If LCase(sFunction) = "save" Then
    SuccessCount = 0
    ErrorCount = 0
    Dim BaseIDArray
    Dim BaseGPG
    Dim ImageArray
    Dim txtSuccess
    For Each item in Request.Form
        If instr(item, "Base") > 0 Then
            BaseIDArray  = split(Request.Form(item),",")            
            For i = lbound(BaseIDArray) To ubound(BaseIDArray)
                GetBaseGPG trim(BaseIDArray(i))
             Next	
        End If
    Next
    
    If ErrorCount > 0 Then
        Response.Write(ErrorCount & " Records Failed...")
    Else
        txtSuccess = "YES"
        Response.Write(SuccessCount & " Records Created Successfully...")
    End If
   
End If

Sub GetBaseGPG(BaseID)
    Dim BaseGPG    
    For Each item in Request.Form
        If instr(item, "txtGPGDescription" & BaseID) > 0 Then
            BaseGPG = trim(Request.Form(item))            
            GetSelectedRows BaseID, BaseGPG
        End If
    Next
End Sub

Sub GetSelectedRows(BaseID, BaseGPG)
    For Each item in Request.Form        
        If instr(item, "chkRow" & BaseID) > 0 Then
            arSelectedRow = split(Request.Form(item),",")
            Dim FeatureID, AvParentID, NewAVParentID
            FeatureID = 0
            AvDetailID = 0
            NewAVParentID = 0
            For j = lbound(arSelectedRow) To ubound(arSelectedRow)
                SelectedRow = trim(arSelectedRow(j))                   
                arData = split(SelectedRow,"|")               
                set cn = server.CreateObject("ADODB.connection")
                set cmd = Server.CreateObject("ADODB.Command")
                On Error Resume Next                       
	            cn.ConnectionString = Session("PDPIMS_ConnectionString")
	            cn.Open
                cmd.CommandText = "usp_InsertLocalized_Pulsar"
                cmd.CommandType = adCmdStoredProc
                set cmd.ActiveConnection = cn
                    
                If FeatureID = clng(arData(0)) and clng(arData(3)) = 0 then
                    AvParentID = NewAVParentID
                Else
                    AvParentID = clng(arData(3))
                End If

                cmd.Parameters.Append cmd.CreateParameter("@PVID", adInteger, adParamInput, , clng(request("PVID")))
                cmd.Parameters.Append cmd.CreateParameter("@BID", adInteger, adParamInput, , clng(request("BID")))
                cmd.Parameters.Append cmd.CreateParameter("@FeatureID", adInteger, adParamInput, , clng(ScrubSQL(arData(0))))
                cmd.Parameters.Append cmd.CreateParameter("@UserName", adVarChar, adParamInput, 100, ScrubSQL(strUserName))                   
                cmd.Parameters.Append cmd.CreateParameter("@ConfigCode", adVarChar, adParamInput, 5, ScrubSQL(arData(2)))
                cmd.Parameters.Append cmd.CreateParameter("@CountryCode", adVarChar, adParamInput, 10, arData(1))
                cmd.Parameters.Append cmd.CreateParameter("@AVParentID", adInteger, adParamInput, , clng(AvParentID))
                cmd.Parameters.Append cmd.CreateParameter("@GeoID", adInteger, adParamInput, , clng(arData(4)))
                cmd.Parameters.Append cmd.CreateParameter("@NewAVParentID", adInteger, adParamOutput)
                cmd.Parameters.Append cmd.CreateParameter("@ShareAV", adInteger, adParamInput, 0)
                cmd.Parameters.Append cmd.CreateParameter("@ReleaseIDs", adVarChar, adParamInput, 250, ScrubSQL(request("Releases")))
                cmd.Parameters.Append cmd.CreateParameter("@RTPDate", adVarChar, adParamInput, 25, ScrubSQL(request("RTPDate")))
                cmd.Parameters.Append cmd.CreateParameter("@EMDate", adVarChar, adParamInput, 25, ScrubSQL(request("EMDate")))
                
	            cn.BeginTrans()                                       
                cmd.Execute

                NewAVParentID = cmd.Parameters("@NewAVParentID")	                
                FeatureID = clng(arData(0))

                If Err.number <> 0 then
                    cn.RollbackTrans()
	                ErrorCount = ErrorCount + 1
                else 
                    cn.CommitTrans()
	                SuccessCount = SuccessCount + 1	
                end if
                set cmd=nothing
                set cn=nothing                              
            Next
        End If
    Next
End Sub
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    function body_onload(pulsarplusDivId) {
        var txtSuccess = document.getElementById("txtSuccess");

        if (txtSuccess.value == "YES") {
            window.returnValue = "YES";
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.closeExternalPopup();
            }
            else {
                window.close();
            }
        } else {
            var checkBoxes = document.getElementsByTagName("input");
            for (i = 0; i < checkBoxes.length; i++) {
                if (checkBoxes[i].className == "chkBase" && !checkBoxes[i].disabled) {
                    var ids = checkBoxes[i].id.split('-');
                    SetUpParentCheckBox(ids[0], ids[1]);
                }
            }
        }
    }

    function chkAll_onclick() {       
        var i, n;
        var checkBoxes = document.getElementsByTagName("input");
        var chkAll = document.getElementById("chkAll");

        if (chkAll.indeterminate) {
            chkAll.indeterminate = false;
            chkAll.checked = false;
        }

        for (i = 0; i < checkBoxes.length; i++) {
            if (!checkBoxes[i].disabled) {
                if (checkBoxes[i].indeterminate)
                    checkBoxes[i].indeterminate = false;

                checkBoxes[i].checked = chkAll.checked;
            }
        }

        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].className == "chkBase" && !checkBoxes[i].disabled) {
                var ids = checkBoxes[i].id.split('-');
                SetUpParentCheckBox(ids[0], ids[1]);                
            }
        }        
        var Count = 0;
        var CheckedCount = 0;
        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].className == "chkBase") {
                Count++;
                if (checkBoxes[i].checked)
                    CheckedCount++;
            }
        }

        if (Count > CheckedCount && CheckedCount > 0)
            chkAll.indeterminate = true;

    }

    function SetUpParentCheckBox(FeatureID, AvDetailID) {
        var i = document.getElementsByName("chkRow" + FeatureID + '-' + AvDetailID).length;
        var n = $('input:checkbox[id^="chkRow' + FeatureID + '-' + AvDetailID + '"]:checked').length;

        var chkBase = document.getElementById(FeatureID + '-' + AvDetailID);

        if (i != n && n > 0) {
            chkBase.indeterminate = true;
        }
        else if (i == n && n > 0) {
            chkBase.checked = true;
        }
    }

    function chkBase_onclick(FeatureID, AvDetailID) {
        var i;
        var chkBase = document.getElementById(FeatureID + '-' + AvDetailID);
        if (chkBase.indeterminate) {
            chkBase.indeterminate = false;
            chkBase.checked = false;
        }

        var checkBoxes = document.getElementsByName("chkRow" + FeatureID + '-' + AvDetailID);
        for (i = 0; i < checkBoxes.length; i++) {
            if (!checkBoxes[i].disabled) {
                if (checkBoxes[i].indeterminate)
                    checkBoxes[i].indeterminate = false;

                checkBoxes[i].checked = chkBase.checked;
            }
        }

        SetUpParentCheckBox(FeatureID, AvDetailID);       
    }

    function UpdateBase(chkClicked) {
        var i;
        var blnAllSame = true;       

        var checkedVals = $('.' + chkClicked.className + ':checkbox:not(:checked)').map(function () {
            return this.value;
        }).get();

        if (checkedVals.join(",").length > 0)
            blnAllSame = false;

        var base = document.getElementById(chkClicked.className)
        if (blnAllSame) {
            base.indeterminate = false;
            base.checked = chkClicked.checked;
        }
        else {
            base.indeterminate = true;
            base.checked = true;
        }
    }

    function chkRow_onclick(FeatureID, AvDetailID) {
        UpdateBase(window.event.srcElement);
    }    

    function BaseRow_onmouseover(strID, avdetailid) {
        if (window.event.srcElement.id == strID + "-" + avdetailid) {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("BaseRow" + strID + "-" + avdetailid).bgColor = "lightsteelblue";
    }

    function BaseRow_onmouseout(strID, avdetailid) {
        document.all("BaseRow" + strID + "-" + avdetailid).bgColor = "cornsilk";
    }

    function BaseRow_onclick(strID, AvDetailID) {
        if (window.event.srcElement.tagName == "INPUT")
            return;

        if (document.all("ChildRow" + strID + "-" + AvDetailID).style.display == "")
            document.all("ChildRow" + strID + "-" + AvDetailID).style.display = "none";
        else
            document.all("ChildRow" + strID + "-" + AvDetailID).style.display = "";
    }


    function ChooseLanguages(strID, strAll) {
        var strResults;
        var i;

        strResults = window.showModalDialog("ChangeImageLanguage.asp?SelectedLangs=" + document.all("Lang" + strID).innerText + "&AllLangs=" + strAll, "", "dialogWidth:300px;dialogHeight:240px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
        if (typeof (strResults) != "undefined") {
            if (strResults != "") {
                if (strAll == strResults) {
                    document.all("Row" + strID).bgColor = "ivory";
                    document.all("chkRow" + strID).indeterminate = 0;
                    document.all("chkRow" + strID).checked = true;
                    UpdateBase(document.all("chkRow" + strID));
                }
                else {
                    document.all("Row" + strID).bgColor = "mistyrose";
                    document.all("chkRow" + strID).indeterminate = -1;
                    document.all(document.all("chkRow" + strID).className).indeterminate = -1;
                }
                document.all("Lang" + strID).innerText = strResults;
            }
        }
    }

    function ChangeThis_onmouseover() {
        window.event.srcElement.style.cursor = "hand";
    }

    function ChangeDefault_onmouseover() {
        window.event.srcElement.style.cursor = "hand";

    }

    function GpgDescription_onchange(definitionId) {
        var baseDescription = document.getElementById("txtGpgDescription" + definitionId);
        var elements = document.getElementsByTagName("span");
        for (i = 0; i < elements.length; i++) {
            if (elements[i].className == "gpgDescription" + definitionId)
                elements[i].innerHTML = baseDescription.value;
        }

    }


    //-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<STYLE>
TH
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}

.HWKitTable TBODY TD{
	BORDER-TOP: gray thin solid;
}

.ChildRows TBODY TD{
	BORDER-TOP: none;
}

.ChildRows THEAD TD{
	BORDER-TOP: none;
}

A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}


</STYLE>
<body bgcolor="Ivory" onload="body_onload('<%Request("pulsarplusDivId")%>')">


<%
	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strAllRows
	dim LastFeatureID
    dim LastAvDetailID
	dim strLanguageList
	dim strRows
	dim strBase
	dim YesCount
	dim NoCount
	dim MixedCount
	dim strSavedLanguages
	dim strLoadedLanguages
	dim strCellColor
	dim strAllChecked
	dim TotalImageDefsChecked
	dim CategoryName
    dim EnableImaging
    dim sqlCustomeError

    dim LastSharedAV
    dim LastIsUsedIOSCM
    dim LastAvId
    dim rowDisabled
    dim rowStyle

	TotalImageDefsChecked = 0
	strAllChecked = ""
	AnyImagesChecked=false
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	
    
    rs.open "usp_GetFeatureCategoryById " & request("CategoryID"), cn, adOpenForwardOnly
    do while not rs.EOF
        CategoryName = rs("Name")
        EnableImaging = rs("EnableImaging")
        rs.MoveNext
    loop
    rs.Close

	strAllRows = ""
    ON ERROR RESUME NEXT
	rs.open "usp_ListLocalization_Pulsar " & request("PVID") & "," & request("BID") & "," & request("CategoryID") & "," & request("ShowAllLocs"), cn, adOpenForwardOnly
        
    if ERR.number <> 0 then
        sqlCustomeError = ERR.Description
    end if

	LastFeatureID = ""
    LastAvDetailID = ""
	
    if rs.EOF and rs.BOF then
        if sqlCustomeError = "" then
            strAllRows = "<tr><td colspan=10><FONT size=1 face=verdana><p>No items found by selected category.  Please make sure there are countries added to the product in Localization tab.</p></font></td></tr>"
        else 
            strAllRows = "<tr><td colspan=10><FONT size=1 face=verdana><P>" + sqlCustomeError + "</P></font></td></tr>"
        end if
    else
		dim imagecount
		imagecount=0
		
		BusinessID = rs("BusinessID")
		dim whereUseMessages

		do while not rs.EOF

				imagecount = imagecount + 1
				if ((LastAvDetailID <> rs("AvDetailID") and LastFeatureID <> rs("FeatureID")) or (LastAvDetailID <> rs("AvDetailID") and LastFeatureID = rs("FeatureID")) or (rs("AvDetailID") = 0 and LastFeatureID <> rs("FeatureID"))) and LastFeatureID <> "" and LastAvDetailID <> "" then
					strAllRows = strAllRows & strBase                   
                    
                    whereUseMessages = ""
                    if (LastSharedAV = 1 and LastIsUsedIOSCM > 0 and IsNull(LastAvId)) then
                        whereUseMessages = "<br><font size=1 face='verdana' color='red'>This Base AV is used in other SCM, therefore you cannot localize it.</font></br>"
                    end if

                    if EnableImaging = false then 
					    strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7>" & whereUseMessages & "<br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td><td><b>Share AV</b></td><td><b>Status</b></td></thead>" & strRows & "</table><br></td></tr>"
					else 
                        strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=8>" & whereUseMessages & "<br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td><td><b>Share AV</b></td><td><b>Status</b></td><td><b>In Image</b></td></thead>" & strRows & "</table><br></td></tr>"
                    end if                         
    
                    strRows = ""                    

					YesCount = 0
					NoCount = 0
					MixedCount=0
				end if
				LastFeatureID = rs("FeatureID")
				LastAvDetailID = rs("AvDetailID")
                LastSharedAV = rs("SharedAV")
                LastIsUsedIOSCM = rs("IsUsedIOSCM")
                LastAvId = rs("AvId")

				strCellColor = "ivory"
			    	
                rowDisabled = ""
                rowStyle = ""
                if (rs("GeoID") = 5) then
                    rowDisabled = "Disabled"
                    rowStyle = "style=""color:grey"""
                end if

				if (rs("AvId") = "" Or IsNull(rs("AvId"))) and rs("Status") <> "O" and (rs("IsUsedIOSCM") = 0 or IsNull(rs("IsUsedIOSCM"))) then                    
                    if rs("InImage") = 0 then 
                        ' Task 16448 ----pass featureID and AvdetaildID to--- chkRow_onclick function 
				        strRows = strRows & "<TR bgcolor=ivory "& rowStyle &"><TD><input "& rowDisabled &" class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" id=chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "|" & rs("CountryCode") & "|" & rs("OptionConfig") & "|" & rs("AvParentID") & "|" & rs("GeoID") & "|" & rs("SharedAV") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>" 
					    NoCount = NoCount + 1
                    else 
                        strRows = strRows & "<TR bgcolor=ivory "& rowStyle &"><TD><input "& rowDisabled &" class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" checked id=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "|" & rs("CountryCode") & "|" & rs("OptionConfig") & "|" & rs("AvParentID") & "|" & rs("GeoID") & "|" & rs("SharedAV") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>"
                        YesCount = YesCount + 1    
                    end if                    
				elseif (rs("AvId") = "" Or IsNull(rs("AvId"))) and (rs("Status") = "O" Or rs("IsUsedIOSCM") > 0) then
                    strRows = strRows & "<TR bgcolor=ivory style=""color:grey""><TD><input Disabled class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" id=chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "|" & rs("CountryCode") & "|" & rs("OptionConfig") & "|" & rs("AvParentID") & "|" & rs("GeoID") & "|" & rs("SharedAV") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>"
					NoCount = NoCount + 1
                else
					strCellColor = "ivory"
					YesCount = YesCount + 1                    
					strRows = strRows & "<TR bgcolor=ivory style=""color:grey""><TD><input Disabled class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" checked id=chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onclick=""return chkRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" style=""WIDTH:16;HEIGHT:16""></TD>"
				end if
				'Task 16448 ---- give id for ---- Base checkboxes--------------

                if ((rs("SharedAV") = 1 and rs("IsUsedIOSCM") > 0) Or rs("Status") = "O") and IsNull(rs("AvId")) then
                    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox""  class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>"
                else
			        if YesCount = 0 then' and MixedCount=0 
				        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox""  class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>" 
				    elseif NoCount=0 then'and MixedCount=0  
					    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox""  class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>" 
				        TotalImageDefsChecked = TotalImageDefsChecked + 1
				    else
				        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") &  "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" indeterminate=-1></TD>"
				        TotalImageDefsChecked = TotalImageDefsChecked + 1
				    end if
                end if
                
                'Task 16448 ---- give id for Description td to get the Feature Description
				strBase = strBase &  "<TD width=400 id=""BaseDesc" & rs("FeatureID") & "-" & rs("AvDetailID") & """>" & rs("Name") & "</TD>"                
                strBase = strBase &  "<TD>" & trim(rs("FeatureID")) & "</TD>"
				strBase = strBase &  "<TD width=100>" & rs("AvBaseNo") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD width=250 style='display:none'><input style=""width:250; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDescription" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=txtGPGDescription" & rs("FeatureID") & "-" & rs("AvDetailID") & " value=""" & rs("BaseGPGDescription") & """ onchange=""GpgDescription_onchange(" & rs("FeatureID") & "-" & rs("AvDetailID") & ")""></TD>"
               
    
                'Task 16448----BASE--add id for td--BASE --(used to check GPG description)--------------------------
                strBase = strBase &  "<TD width=250 id=""chkBasetd" & rs("FeatureID") & "-" & rs("AvDetailID") & """>" & rs("BaseGPGDescription") & "</TD>"
                
                if rs("ProductBrandID") > 0 then
                    strBase = strBase &  "<TD>Yes</TD>"
                else
                    strBase = strBase &  "<TD>No</TD>"
                end if

                if rs("SharedAV") = 0 then
                    strBase = strBase &  "<TD>No</TD>"                    
                else
                    strBase = strBase &  "<TD>Yes</TD>"                   
                end if

                strBase = strBase &  "<TD>" & rs("Status") & "</TD>"
                strBase = strBase &  "</tr>" & vbcrlf
				
				strRows = strRows & "<TD width=150>" & rs("AVNo") & "&nbsp;</TD>"
		        strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
		        
                'Task 16448------add id for span element-- ROWS-(used to check GPG description)-----
				strRows = strRows & "<TD width=250 ><span id=""chkRowtd" & rs("FeatureID") & "-" & rs("AvDetailID") & """ class=gpgDescription" & rs("FeatureID") & ">" & rs("BaseGPGDescription") & "</span>&nbsp;" & rs("CountryCode") & "</TD>"
                
				strRows = strRows & "<TD>" & rs("CountryCode") & "<input type=hidden name=CountryCode" & rs("FeatureID") & " value=""" & rs("CountryCode") & """></TD>"
				
				strRows = strRows & "<TD>" & rs("OptionConfig") & "<input type=hidden name=OptionConfig" & rs("FeatureID") & " value=""" & rs("OptionConfig") & """></TD>"
				
				strRows = strRows & "<TD>" & rs("Keyboard") & "</TD>"
				strRows = strRows & "<TD>" & rs("Powercord") & "</TD>"
                
                if rs("AvID") > 0 then
                    strRows = strRows & "<TD>Yes</TD>"
                else
                    strRows = strRows & "<TD>No</TD>"
                end if
                
                if rs("LocSharedAV") = false then
                    strRows = strRows &  "<TD>No</TD>"
                else
                    strRows = strRows &  "<TD>Yes</TD>"
                end if

                if rs("LocStatus") = " " then
                    strRows = strRows &  "<TD>-</TD>"
                else 
                    strRows = strRows &  "<TD>" & rs("LocStatus") & "</TD>"
                end if

                if EnableImaging = true then
                    if rs("InImage") = 0 then
                        strRows = strRows & "<TD>No</TD>"
                    else
                        strRows = strRows & "<TD>Yes</TD>"
                    end if
                end if

                if rs("SharedAV") = 1 and rs("IsUsedIOSCM") > 0 and IsNull(rs("AvId")) then
                    whereUseMessages = "<br><font size=1 face=verdana color=red>This Base AV is used in other SCM, therefore you cannot localize it.</font></br>"
                end if

				strRows = strRows & "</tr>" & vbcrlf
			rs.MoveNext
		loop

		if imagecount = 0 then
			strAllRows = "<tr><td colspan=10><font size=1 face=verdana>No active items defined for this product.</font></td></tr>"
		end if
		
        strAllRows = strAllRows & strBase
        if EnableImaging = false then 
	        strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7>" & whereUseMessages & "<br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td><td><b>Share AV</b></td><td><b>Status</b></td></thead>" & strRows & "</table><br></td></tr>"
		else 
            strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=8>" & whereUseMessages & "<br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td><td><b>Share AV</b></td><td><b>Status</b></td><td><b>In Image</b></td></thead>" & strRows & "</table><br></td></tr>"
        end if
		strRows = ""
	end if	
	rs.Close
%>

<div id="inlineProgressBarDiv" style="display:none;">
    <div style="font-family: Verdana; font-size: 12px; position: absolute; z-index: 10; position: absolute; left: 0px">
        <table align="center" cellpadding="0" cellspacing="0" style="background-color: Ivory; height: 50px;">
            <tr>
                <td width="13px"> </td>
                        <td valign="middle" align="center">
                            <img src="../Images/loading.gif"  runat="server">
                        </td>
                 <td valign="middle" style="font-family: Verdana; font-size: 12px;"> Processing, Please Wait...</td>
              </tr>
         </table>
    </div>
</div>

<h3 align="center"><TH align=center>Add <%= CategoryName %> Localized AVs</TH></h3>
<h5 align="center"><TH align="center">Please select items to create the Localized AVs.</TH></h5>

<%If BusinessID = 2 and Request("CategoryID") = 18 Then%>
<table id="ColorLegendTable" border=0 cellspacing=0 cellpadding=1 >
  <thead bgcolor=Wheat>
    <th align=left>Keyboard Categories:</th>
  </thead>
  <tr><td colspan=10><FONT size=1 face=verdana>ST  = Standard</font></td></tr>
  <tr><td colspan=10><FONT size=1 face=verdana>SB = Standard Backlight</font></td></tr>
  <tr><td colspan=10><FONT size=1 face=verdana>CL  = Colored</font></td></tr>
  <tr><td colspan=10><FONT size=1 face=verdana>CB = Colored Backlight</font></td></tr>		
</table>
<%End If %>
<!--<b><font size=2 face=verdana><%=strDeliverable & " (" & strProdName & ")"%></font></b>-->

<form ID=frmChange method=post>
<input id="hidFunction" name="hidFunction" type=HIDDEN value="<%= LCase(sFunction)%>" />
<!--<font size=2 face=verdana color=green>Note: Inactive and Released images are not displayed on this page.</font><BR><BR>-->
<!--<b><font size=2 face=verdana>Keyboards:</font><BR></b>-->

		<table width=100% border=0 cellspacing=0 cellpadding=1 id="tblItems">
			<thead bgcolor="Wheat">
				<!--<% if request("Type") = "1" then%>
				<TH align=left><input disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				<%else%>-->
				<th align="left"><input type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()" /></th>
                <!--<%end if%>-->                
				<th align="left">Description</th>
                <th align="left">Feature ID</th>
				<th align="left">Base AV Number</th>
				<!--<th align=left>ZWAR</th>-->
                <th align="left">Base GPG Description</th>
                <th align="left">In SCM</th>                          
                <th align="left">Shared AV</th>
                <th align="left">Status</th>
			</thead>
			<%=strAllRows%>
		</table>
<input style="Display:none" type="checkbox" id=chkAllChecked name=chkAllChecked />
<input type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>" />
<input type="hidden" id=txtFeatureID name=txtFeatureID value="<%=request("FeatureID")%>" />
<input type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>" />
<input type="hidden" id=txtSuccess name=txtSuccess value="<%= txtSuccess %>" />
<input type="hidden" id=txtEnableImaging name=txtEnableImaging value="<%= EnableImaging %>" />
<input type="hidden" id="txtUserName" name="txtUserName" value="<%=strUserName %>" />
<input type="hidden" id="txtPVID" name="txtPVID" value="<%=request("PVID")%>" />
<input type="hidden" id="txtBID" name="txtBID" value="<%=request("BID") %>" />
<input type="hidden" id="txtReleases" name="txtReleases" value="<%=request("Releases") %>" />
<input type="hidden" id="txtRTPDate" name="txtRTPDate" value="<%=request("RTPDate") %>" />
<input type="hidden" id="txtEMDate" name="txtEMDate" value="<%=request("EMDate") %>" />


</form>


<%
	set rs= nothing
	set cn=nothing
%>


</body>
</HTML>