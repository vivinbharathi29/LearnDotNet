<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/lib_debug.inc" --> 
<%
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

	'	strWords=replace(strWords,"'","''")
		
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
                'If strUserID = 5016 Or strUserID = 8 Or strUserID = 1396 Then
                    set cn = server.CreateObject("ADODB.connection")
                    set cmd = Server.CreateObject("ADODB.Command")
                    On Error Resume Next                       
	                cn.ConnectionString = Session("PDPIMS_ConnectionString")
	                cn.Open
                    cmd.CommandText = "usp_InsertLocalizedKeyboard_Pulsar"
                    cmd.CommandType = adCmdStoredProc
                    set cmd.ActiveConnection = cn
                    
                    If FeatureID = clng(arData(0)) and clng(arData(3)) = 0 then
                        AvParentID = NewAVParentID
                    Else
                        AvParentID = clng(arData(3))
                    End If

                    cmd.Parameters.Append cmd.CreateParameter("@PVID", adInteger, adParamInput, , clng(request("PVID")))
                    cmd.Parameters.Append cmd.CreateParameter("@BID", adInteger, adParamInput, , clng(request("BID")))
                    cmd.Parameters.Append cmd.CreateParameter("@BaseGPG", adVarChar, adParamInput, 50, ScrubSQL(BaseGPG))
                    cmd.Parameters.Append cmd.CreateParameter("@FeatureID", adInteger, adParamInput, , clng(ScrubSQL(arData(0))))
                    cmd.Parameters.Append cmd.CreateParameter("@UserName", adVarChar, adParamInput, 100, ScrubSQL(strUserName))                   
                    cmd.Parameters.Append cmd.CreateParameter("@ConfigCode", adVarChar, adParamInput, 5, ScrubSQL(arData(2)))
                    cmd.Parameters.Append cmd.CreateParameter("@CountryCode", adVarChar, adParamInput, 10, ScrubSQL(arData(1)))
                    cmd.Parameters.Append cmd.CreateParameter("@AVParentID", adInteger, adParamInput, , clng(AvParentID))
                    cmd.Parameters.Append cmd.CreateParameter("@GeoID", adInteger, adParamInput, , clng(arData(4)))
                    cmd.Parameters.Append cmd.CreateParameter("@NewAVParentID", adInteger, adParamOutput)

	                cn.BeginTrans()                                       
                    cmd.Execute

                    NewAVParentID = cmd.Parameters("@NewAVParentID")
	                'cn.Execute "usp_InsertLocalizedKeyboard_Pulsar " & clng(request("PVID")) & "," & clng(request("BID")) & ",'" & ScrubSQL(BaseGPG) & "'," & clng(ScrubSQL(arData(0))) & ",'" & ScrubSQL(strUserName) & "','" & ScrubSQL(arData(2)) & "','" & ScrubSQL(arData(1)) & "'," & clng(AvParentID) & "," & clng(arData(4))
	                'NewAVParentID = cn.Parameters("@NewAVParentID").value
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
                'Else
                '    Response.Write "PVID(" & request("PVID") & ") BID(" & request("BID") & ") baseGPG(" & BaseGPG & ") FeatureID(" & BaseID & ") CountryCode(" & arData(1) & ") ConfigCode(" & arData(2) & ") UserName(" & strUserName & ") </br>" 
                'End If                
            Next
        End If
    Next
End Sub
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function body_onload() {
        var txtSuccess = document.getElementById("txtSuccess");

        if (txtSuccess.value == "YES") {
            window.returnValue = "YES";
			window.close();
        } else {
            chkAll_onclick();
        }
}

function chkAll_onclick() {
	var i;
	var checkBoxes = document.getElementsByTagName("input");
	var chkAll = document.getElementById("chkAll");

	for (i=0;i<checkBoxes.length;i++)
    {
    	if (checkBoxes[i].className == "chkBase" && !checkBoxes[i].disabled)
		{
    	    if (checkBoxes[i].indeterminate) {
		        checkBoxes[i].indeterminate = false;
		    }
		    checkBoxes[i].checked = chkAll.checked;
		    var ids = checkBoxes[i].id.split('-');
		    chkBase_onclick(ids[0], ids[1]);
    	}

	}

}


function chkBase_onclick(FeatureID, AvDetailID) {
	var i;
	var checkBoxes = document.getElementsByName("chkRow" + FeatureID + '-' + AvDetailID);
	var chkBase = document.getElementById(FeatureID + '-' + AvDetailID);
    var hasDisabled = false;
    var hasChecked = false;
    
    if (chkBase.indeterminate) {
        chkBase.indeterminate = false;
        chkBase.checked = true;
    }

    for (i = 0; i < checkBoxes.length; i++) {
        if (!checkBoxes[i].disabled) {
            checkBoxes[i].checked = chkBase.checked;
        } else {
            hasDisabled = true;
            if (checkBoxes[i].checked) {
                hasChecked = true;
            }
        }
    }

    if (hasChecked && !chkBase.checked) {
        chkBase.indeterminate = true;        
    }
        
} 

function UpdateBase(chkClicked) {
	var i;
	var blnAllSame=true;
	var chkRow = document.getElementsByTagName("input");

	for (i = 0; i < chkRow.length; i++)
		{
		if (chkRow(i).className != "")
			if (chkRow(i).className == chkClicked.className)
				{
				if ((chkRow(i).checked != chkClicked.checked) || chkRow(i).indeterminate)
					{
						blnAllSame = false;	
					}
				}
		}

	var base = document.getElementById(chkClicked.className)
	if (blnAllSame) {
	    base.indeterminate = false;
	    base.checked = chkClicked.checked;
	}
	else {
	    base.indeterminate = true;
	    base.checked = true;
	}

	//if (chkClicked.checked)
	//	{
   	//		document.all("Lang" + chkClicked.value).innerText = document.all("Lang" + chkClicked.value).className;
	//		if (document.all("Row" + chkClicked.value)!=null)
	//			document.all("Row" + chkClicked.value).bgColor = "ivory";
	//	}

}


function chkRow_onclick() {
	UpdateBase(window.event.srcElement);
} 

function BaseRow_onmouseover(strID, avdetailid){
	if (window.event.srcElement.id == strID + "-" + avdetailid)
		{
		    window.event.srcElement.style.cursor = "default";
		    return;
		}

	window.event.srcElement.style.cursor = "hand";
	document.all("BaseRow" + strID + "-" + avdetailid).bgColor="lightsteelblue";	
}

function BaseRow_onmouseout(strID, avdetailid){
	document.all("BaseRow" + strID + "-" + avdetailid).bgColor="cornsilk";	
}

function BaseRow_onclick(strID, AvDetailID){
	if (window.event.srcElement.tagName == "INPUT")
	    return;

	if (document.all("ChildRow" + strID + "-" + AvDetailID).style.display == "")
	    document.all("ChildRow" + strID + "-" + AvDetailID).style.display = "none";
	else
	    document.all("ChildRow" + strID + "-" + AvDetailID).style.display = "";
}


function ChooseLanguages(strID,strAll){
	var strResults;
	var i;
	
	strResults = window.showModalDialog("ChangeImageLanguage.asp?SelectedLangs=" + document.all("Lang" + strID).innerText + "&AllLangs=" + strAll ,"","dialogWidth:300px;dialogHeight:240px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	if (typeof(strResults) != "undefined")
		{
			if (strResults != "")
				{
					if (strAll == strResults)
						{
						document.all("Row" + strID).bgColor = "ivory";
						document.all("chkRow" + strID).indeterminate=0;
						document.all("chkRow" + strID).checked =true;
						UpdateBase(document.all("chkRow" + strID));
						}
					else
						{
						document.all("Row" + strID).bgColor = "mistyrose";
						document.all("chkRow" + strID).indeterminate=-1;
						document.all(document.all("chkRow" + strID).className).indeterminate=-1;
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

//function ChangeThis_onclick() {
//	frmChange.optThis.checked=true;
//	frmChange.optFuture.checked=false;
//}

//function ChangeDefault_onclick() {
//	frmChange.optThis.checked=false
//	frmChange.optFuture.checked=true
//}

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
<body bgcolor="Ivory" onload="body_onload()">


<%
'function GetLanguages(strImages, strID)
'	dim strTemp
'	
'	if instr(strImages,trim(strID) & "=") = 0 then
'		GetLanguages = ""
'	else
'		strTemp = mid(strimages,instr(strImages,trim(strID) & "=")+ len(trim(strID) & "="))
'		strTemp = mid(strTemp,1,instr(strTemp,")") -1) 'Strip off )...
'		GetLanguages = strTemp
'	end if
'end function

	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
'	dim strImageSummary
'	dim strImages
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
	'dim blnImages
	dim strAllChecked
	dim TotalImageDefsChecked
	
	TotalImageDefsChecked = 0
	strAllChecked = ""
	AnyImagesChecked=false
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
'	strImageSummary = ""
'	strImages = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	
	'rs.Open "spGetImagesWithAVID " & clng(request("PVID")),cn,adOpenForwardOnly
	'strImages = ""
	'do while not rs.EOF
	'    if strImages = "" then
	'        strImages = trim(rs("ID"))
	'    else
	'        strImages = strImages & ", " & trim(rs("ID"))
	'    end if
	'    rs.MoveNext
	'loop
	'rs.Close

    'strImages = ""
	    
	'blnImages = true
	'if instr(strImages,":")>0 then
	'	if left(strImages,instr(strImages,":")-1)<>"" then
	'		blnImages = true
	'	end if
	'elseif strImages <> "" then
	'	blnImages = true
	'end if

	'if trim(strImageSummary) = "" then
	'	strImageSummary = "ALL"
	'end if

	'if trim(strImages) <> "" then
	'	strImages = ", " & strImages & ","
	'end if
	
	strAllRows = ""

	rs.open "spListKeyboardsForLocalization_Pulsar " & request("PVID") & "," & request("BID"), cn, adOpenForwardOnly

	LastFeatureID = ""
    LastAvDetailID = ""
	if rs.EOF and rs.BOF then
		strAllRows = "<tr><td colspan=10><FONT size=1 face=verdana>No Keyboards defined for this product.</font></td></tr>"
	else
		dim imagecount
		imagecount=0
		
		BusinessID = rs("BusinessID")
		
		do while not rs.EOF                
				imagecount = imagecount + 1
				if ((LastAvDetailID <> rs("AvDetailID") and LastFeatureID <> rs("FeatureID")) or (LastAvDetailID <> rs("AvDetailID") and LastFeatureID = rs("FeatureID")) or (rs("AvDetailID") = 0 and LastFeatureID <> rs("FeatureID"))) and LastFeatureID <> "" and LastAvDetailID <> "" then
					strAllRows = strAllRows & strBase
					strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td></thead>" & strRows & "</table><br></td></tr>"
					strRows = ""
					YesCount = 0
					NoCount = 0
					MixedCount=0
				end if
				LastFeatureID = rs("FeatureID")
				LastAvDetailID = rs("AvDetailID")

				strCellColor = "ivory"
								
				if rs("AvId") = "" Or IsNull(rs("AvId")) then
				    strRows = strRows & "<TR bgcolor=ivory><TD><input class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" id=chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "|" & rs("CountryCode") & "|" & rs("OptionConfig") & "|" & rs("AvParentID") & "|" & rs("GeoID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkRow_onclick()""></TD>"
					NoCount = NoCount + 1
				else
					strCellColor = "ivory"
					YesCount = YesCount + 1
					strRows = strRows & "<TR bgcolor=ivory style=""color:grey""><TD><input Disabled class=""" & trim(rs("FeatureID")) & "-" & rs("AvDetailID") & """ type=""checkbox"" checked id=chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=""chkRow" & rs("FeatureID") & "-" & rs("AvDetailID") & """ value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onclick=""return chkRow_onclick()"" style=""WIDTH:16;HEIGHT:16""></TD>"
				end if
				
			    if YesCount = 0 then' and MixedCount=0 
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>" 
				elseif NoCount=0 then'and MixedCount=0  
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")""></TD>" 
				    TotalImageDefsChecked = TotalImageDefsChecked + 1
				else
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("FeatureID") & "-" & rs("AvDetailID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" onclick=""return BaseRow_onclick(" & rs("FeatureID") &  "," & rs("AvDetailID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ name=""Base"" value=""" & rs("FeatureID") & "-" & rs("AvDetailID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("FeatureID") & "," & rs("AvDetailID") & ")"" indeterminate=-1></TD>"
				    TotalImageDefsChecked = TotalImageDefsChecked + 1
				end if

				strBase = strBase &  "<TD width=400>" & rs("Name")
                if rs("AvDetailID") = 0 then
                    strBase = strBase & "</TD>"
                else
                    strBase = strBase &  " (AvDetailID: " & rs("AvDetailID") & ")</TD>"
                end if
            
				strBase = strBase &  "<TD width=100>" & rs("AvBaseNo") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD width=250 style='display:none'><input style=""width:250; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDescription" & rs("FeatureID") & "-" & rs("AvDetailID") & " name=txtGPGDescription" & rs("FeatureID") & "-" & rs("AvDetailID") & " value=""" & rs("BaseGPGDescription") & """ onchange=""GpgDescription_onchange(" & rs("FeatureID") & "-" & rs("AvDetailID") & ")""></TD>"
                strBase = strBase &  "<TD width=250>" & rs("BaseGPGDescription") & "</TD>"
                
                if rs("ProductBrandID") > 0 then
                    strBase = strBase &  "<TD width=50>Yes</TD>"
                else
                    strBase = strBase &  "<TD width=50>No</TD>"
                end if

                if rs("SharedAV") = 0 then
                    strBase = strBase &  "<TD width=50>No</TD>"
                else
                    strBase = strBase &  "<TD width=50>Yes</TD>"
                end if

                'strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
				strBase = strBase &  "</tr>" & vbcrlf
				
				strRows = strRows & "<TD width=150>" & rs("AVNo") & "&nbsp;</TD>"
		        strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
		        
				strRows = strRows & "<TD width=250><span class=gpgDescription" & rs("FeatureID") & ">" & rs("BaseGPGDescription") & "</span>&nbsp;" & rs("CountryCode") & "</TD>"
				'strRows = strRows &  "<TD width=200> <input style=""width:200; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDesciption name=txtGPGDesciption value=""" & rs("GPGDescription") & """></TD>"
				
				strRows = strRows & "<TD>" & rs("CountryCode") & "<input type=hidden name=CountryCode" & rs("FeatureID") & " value=""" & rs("CountryCode") & """></TD>"
				'strRows = strRows & "<TD>" & rs("CountryCode") & "</TD>"
				
				strRows = strRows & "<TD>" & rs("OptionConfig") & "<input type=hidden name=OptionConfig" & rs("FeatureID") & " value=""" & rs("OptionConfig") & """></TD>"
				'strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
				
				strRows = strRows & "<TD>" & rs("Keyboard") & "</TD>"
				strRows = strRows & "<TD>" & rs("Powercord") & "</TD>"
                
                if rs("AvID") > 0 then
                    strRows = strRows & "<TD>Yes</TD>"
                else
                    strRows = strRows & "<TD>No</TD>"
                end if

				strRows = strRows & "</tr>" & vbcrlf
			rs.MoveNext
		loop
		if imagecount = 0 then
			strAllRows = "<tr><td colspan=10><font size=1 face=verdana>No active Keyboards defined for this product.</font></td></tr>"
		end if
		strAllRows = strAllRows & strBase
		strAllRows = strAllRows & "<tr style=""Display:none"" id=""ChildRow" & LastFeatureID & "-" & LastAvDetailID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><Td><b>Dash</b></td><td><b>Loacalized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td><td><b>In SCM</b></td></thead>" & strRows & "</table><br></td></tr>"
		strRows = ""
	end if	
	rs.Close

	'if TotalImageDefsChecked > 0 then
	'	strAllChecked="checked"
	'else
	'	strAllChecked=""
	'end if

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

<BR>
<h3 align="center"><TH align=center>Select Keyboards To Localize</TH></h3>
<BR>
<h5 align="center"><TH align="center">Please select only those base AVs that are designated for the selected brand (KMAT) of the product in the PDD.</TH></h5>

<%If BusinessID = 2 Then%>
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
<input id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<!--<font size=2 face=verdana color=green>Note: Inactive and Released images are not displayed on this page.</font><BR><BR>-->
<b><font size=2 face=verdana>Keyboards:</font><BR></b>

		<table id="KeyboardTable" class="KeyboardTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			<thead bgcolor=Wheat>
				<!--<% if request("Type") = "1" then%>
				<TH align=left><input disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				<%else%>-->
				<th align=left><input type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></th>
                <!--<%end if%>-->
				<th align=left>Description</th>
				<th align=left>Base AV Number</th>
				<!--<th align=left>ZWAR</th>-->
				<th align=left>Base GPG Description</th>
                <th align="left">In SCM</th>
                <th align="left">Shared AV</th>
			</thead>
			<%=strAllRows%>
		</table>
<input style="Display:none" type="checkbox" id=chkAllChecked name=chkAllChecked>
<input type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<input type="hidden" id=txtFeatureID name=txtFeatureID value="<%=request("FeatureID")%>">
<input type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<input type="hidden" id=txtSuccess name=txtSuccess value="<%= txtSuccess %>">
</form>


<%
	set rs= nothing
	set cn=nothing
%>


</body>
</HTML>
