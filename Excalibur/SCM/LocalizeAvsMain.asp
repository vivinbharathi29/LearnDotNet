<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
  Dim SuccessCount
  Dim ErrorCount
  Dim sFunction			: sFunction = Request.Form("hidFunction")
  '##############################################################################	
  '
  ' Create Security Object to get User Info
  '
	
	'm_EditModeOn = False
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	strUserID = Security.CurrentUserID()
	strUserName = Security.CurrentUserFullName()

	'm_IsSysAdmin = Security.IsSysAdmin()

	'm_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	'm_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	'm_UserFullName = Security.CurrentUserFullName()
	
	'If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		'm_EditModeOn = True
	'End If
	
	'If Not m_EditModeOn Then
		'sMode = "view"
	'End If

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
        'Response.Write item & " : " & Request.Form(item) & "</br>"
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
            GetSelectedImages BaseID, BaseGPG
        End If
    Next
End Sub

Sub GetSelectedImages(BaseID, BaseGPG)
    For Each item in Request.Form
    '
        If instr(item, "chkImage" & BaseID) > 0 Then
            ImageArray  = split(Request.Form(item),",")
            For j = lbound(ImageArray) To ubound(ImageArray)
                ImageID = trim(ImageArray(j))
                'If strUserID = 5016 Or strUserID = 8 Or strUserID = 1396 Then
                    RowsEffected = 0
                    set cm = server.CreateObject("ADODB.connection")
	                cm.ConnectionString = Session("PDPIMS_ConnectionString")
	                cm.Open
	                cm.BeginTrans()
	                cm.Execute "usp_InsertLocalizedImage " & clng(request("PVID")) & "," & clng(ImageID) & "," & clng(request("BID")) & ",'" & ScrubSQL(BaseGPG) & "'," & clng(BaseID) & ",'" & ScrubSQL(strUserName) & "'" , RowsEffected 
	                If RowsEffected <> 1 Then
	                    cm.RollbackTrans()
	                    ErrorCount = ErrorCount + 1
	                    'Response.Write "PVID(" & request("PVID") & ") BID(" & request("BID") & "imageDefID(" & BaseID & ") baseGPG(" & BaseGPG & ") imageID(" & trim(ImageArray(j)) & ") UserName(" & strUserName & ") </br>" 
	                    'response.Write "RowsEffected:" & RowEffected & "<br>"
	                Else 
	                    cm.CommitTrans()
	                    SuccessCount = SuccessCount + 1
	                End If
	                set cm=nothing
                'Else
                '    Response.Write "PVID(" & request("PVID") & ") BID(" & request("BID") & "imageDefID(" & BaseID & ") baseGPG(" & BaseGPG & ") imageID(" & trim(ImageArray(j)) & ") UserName(" & strUserName & ") </br>" 
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
            if (window.parent.frames["UpperWindow"]) {
                parent.window.parent.ReloadLocalizeAVs("YES");
            } else {
                window.parent.opener = 'X';
                window.parent.open('', '_parent', '')
                window.parent.close();
            }
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
		    chkBase_onclick(checkBoxes[i].id);
		}
	}
}


function chkBase_onclick(definitionId) {
	var i;
	var checkBoxes = document.getElementsByName("chkImage" + definitionId);
    var chkBase = document.getElementById(definitionId);
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
        chkBase.checked = true;
    }
        
} 

function UpdateBase(chkClicked){
	var i;
	var blnAllSame=true;
	var chkImage = document.getElementsByTagName("input");

	for (i = 0; i < chkImage.length; i++)
		{
		if (chkImage(i).className != "")
			if (chkImage(i).className == chkClicked.className)
				{
				if ((chkImage(i).checked != chkClicked.checked) || chkImage(i).indeterminate)
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

	if (chkClicked.checked)
		{
			document.all("Lang" + chkClicked.value).innerText = document.all("Lang" + chkClicked.value).className;
			if (document.all("Row" + chkClicked.value)!=null)
				document.all("Row" + chkClicked.value).bgColor = "ivory";
		}

}


function chkImage_onclick(){
	UpdateBase(window.event.srcElement);
} 

function BaseRow_onmouseover(strID){
	if (window.event.srcElement.id == strID)
		{
		window.event.srcElement.style.cursor = "default";
		return;
		}

	window.event.srcElement.style.cursor = "hand";
	document.all("BaseRow" + strID).bgColor="lightsteelblue";	
}

function BaseRow_onmouseout(strID){
	document.all("BaseRow" + strID).bgColor="cornsilk";	
}

function BaseRow_onclick(strID){

	if (window.event.srcElement.tagName == "INPUT")
	    return;

	if (document.all("ImageRow" + strID).style.display == "" )
		document.all("ImageRow" + strID).style.display="none";	
	else
		document.all("ImageRow" + strID).style.display="";	
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
						document.all("chkImage" + strID).indeterminate=0;
						document.all("chkImage" + strID).checked =true;
						UpdateBase(document.all("chkImage" + strID));
						}
					else
						{
						document.all("Row" + strID).bgColor = "mistyrose";
						document.all("chkImage" + strID).indeterminate=-1;
						document.all(document.all("chkImage" + strID).className).indeterminate=-1;
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

function ChangeThis_onclick() {
	frmChange.optThis.checked=true;
	frmChange.optFuture.checked=false;
}

function ChangeDefault_onclick() {
	frmChange.optThis.checked=false
	frmChange.optFuture.checked=true
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

.ImageTable TBODY TD{
	BORDER-TOP: gray thin solid;
}

.imagerows TBODY TD{
	BORDER-TOP: none;
}

.imagerows THEAD TD{
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
function GetLanguages(strImages, strID)
	dim strTemp
	
	if instr(strImages,trim(strID) & "=") = 0 then
		GetLanguages = ""
	else
		strTemp = mid(strimages,instr(strImages,trim(strID) & "=")+ len(trim(strID) & "="))
		strTemp = mid(strTemp,1,instr(strTemp,")") -1) 'Strip off )...
		GetLanguages = strTemp
	end if
	
end function

	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strImageSummary
	dim strImages
	dim strAllImages
	dim LastDefinition
	dim strLanguageList
	dim strRows
	dim strBase
	dim YesCount
	dim NoCount
	dim MixedCount
	dim strSavedLanguages
	dim strLoadedLanguages
	dim strCellColor
	dim blnImages
	dim strAllChecked
	dim TotalImageDefsChecked
	
	TotalImageDefsChecked = 0
	strAllChecked = "checked"
	AnyImagesChecked=false
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	strImageSummary = ""
	strImages = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")


	'if not blnLoadFailed then
	'	rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
	'	if rs.EOF and rs.BOF then
	'		strProdName = ""
	'		blnLoadFailed = true
	'	else
	'		strprodName = rs("name") & ""
	'	end if
	'	rs.Close
	'end if

	'if not blnLoadFailed then
	'	rs.Open "spGetVersionProperties4Web " & clng(request("VersionID")),cn,adOpenForwardOnly
	'	if rs.EOF and rs.BOF then
	'		strDeliverable = ""
	'		blnLoadFailed = true
	'	else
	'		strDeliverable = rs("name") & "&nbsp;-&nbsp;" & rs("Version")
	'		if trim(rs("Revision") & "") <> "" then
	'			strDeliverable = strDeliverable & "," & rs("Revision")
	'		end if
	'		if trim(rs("Pass") & "") <> "" then
	'			strDeliverable = strDeliverable & "," & rs("Pass")
	'		end if
	'		strDeliverable = strDeliverable &  "&nbsp;"
	'	end if	
	'	rs.Close
	'end if

	'if not blnLoadFailed then
	'	rs.Open "spGetDistributionVersion " & clng(request("ProductID")) & ","  & clng(request("VersionID")),cn,adOpenForwardOnly
	'	if rs.EOF and rs.BOF then
	'		blnLoadFailed = true
	'	else
	'		strImageSummary = rs("ImageSummary") & ""
	'		strImages = rs("Images") & ""
	'		if instr(strImages,":")>0 then
	'			strLoadedLanguages = mid(strImages,instr(strImages,":")+1)
	'		else
	'			strLoadedLanguages = ""
	'		end if			
	'	end if
	'	rs.Close
	'end if
	
	rs.Open "spGetImagesWithAVID " & clng(request("PVID")) & "," & clng(request("BID")),cn,adOpenForwardOnly
	strImages = ""
	do while not rs.EOF
	    if strImages = "" then
	        strImages = trim(rs("ID"))
	    else
	        strImages = strImages & ", " & trim(rs("ID"))
	    end if
	    rs.MoveNext
	loop
	rs.Close

    'strImages = "3001, 10041"
	    
	blnImages = true
	if instr(strImages,":")>0 then
		if left(strImages,instr(strImages,":")-1)<>"" then
			blnImages = true
		end if
	elseif strImages <> "" then
		blnImages = true
	end if

	'if trim(strImageSummary) = "" then
	'	strImageSummary = "ALL"
	'end if

	if trim(strImages) <> "" then
		strImages = ", " & strImages & ","
	end if
	
	strAllImages = ""
    rs.activeconnection.commandtimeout = 60
	rs.open "spListImagesForLocalization " & clng(request("PVID")) & "," & clng(request("BID")) & ",'"  & ScrubSQL(request("strSeriesSummary")) & "'",cn,adOpenForwardOnly
	lastDefinition = ""
	if rs.EOF and rs.BOF then
		strAllImages = "<tr><td colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></tr>"
	else
		dim imagecount
		imagecount=0
		
		do while not rs.EOF
			if rs("StatusID") <> 3 then 'and isnumeric(rs("Priority")) then 'trim(rs("Priority")) = "1" or trim(rs("Priority")) = "2"  or trim(rs("Priority")) = "3"  or trim(rs("Priority")) = "4"  or trim(rs("Priority")) = "5"  or trim(rs("Priority")) = "6" then
				imagecount = imagecount + 1
				if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
			
					strAllImages = strAllImages & strBase
					strAllImages = strAllImages &  "<tr style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV No.</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Config</b></td><td><b>Lang</b></td><td><b>KBD</b></td><td><b>Cord</b></td></thead>" & strRows & "</table><br></td></tr>"
					strRows = ""
					YesCount = 0
					NoCount = 0
					MixedCount=0
				end if
				lastdefinition = rs("DefinitionID")
				
				strLanguageList = rs("OSLanguage")
				if trim(rs("OtherLanguage") & "") <> "" then
					strLanguageList = strLanguageList & "," & rs("OtherLanguage")
				end if	
				
				strSavedLanguages = getLanguages(strImages,rs("ID"))
				if strSavedLanguages = "" then
					strSavedLanguages = strLanguageList
				end if
				
				strCellColor = "ivory"
				
				if instr(strImages,", " & rs("ID") & ",") > 0 or not blnImages then
					strCellColor = "ivory"
					YesCount = YesCount + 1
					'if request("Type") = "1" then
						strRows = strRows & "<TR bgcolor=ivory style=""color:grey""><TD><input Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=""chkImage" & rs("DefinitionID") & """ value=""" & rs("ID") & " LANGUAGE=javascript onclick=""return chkImage_onclick()"" style=""WIDTH:16;HEIGHT:16""></TD>"
					'else
					'	strRows = strRows & "<TR bgcolor=ivory><TD><input class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage" & imageGroup & " LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"							
					'end if
					
				else
					'if request("Type") = "1" then
					'	strRows = strRows & "<TR bgcolor=ivory><TD><input Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage" & imageGroup & " value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""></TD>"
					'else
						strRows = strRows & "<TR bgcolor=ivory><TD><input class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=""chkImage" & rs("DefinitionID") & """ value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""></TD>"
					'end if
					NoCount = NoCount + 1
				end if
				
				if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
					RowStyle = "style=display:none"
				else
					rowStyle=""
				end if
				if YesCount = 0 then' and MixedCount=0 
				  'if request("Type") = "1" then
					'strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
				  'else
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("DefinitionID") & """ name=""Base"" value=""" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("DefinitionId") & ")""></TD>" 
				  'end if
				elseif NoCount=0 then'and MixedCount=0  
				  'if request("Type") = "1" then
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("DefinitionID") & """ name=""Base"" value=""" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("DefinitionId") & ")""></TD>" 
				  'else
					'strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
				  'end if
				  TotalImageDefsChecked= TotalImageDefsChecked + 1
				else
				 'if request("Type") = "1" then
				 'strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
				 'else 
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("DefinitionID") & """ name=""Base"" value=""" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("DefinitionId") & ")"" indeterminate=-1></TD>"
				 'end if
				  TotalImageDefsChecked = TotalImageDefsChecked + 1
				end if
				strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD width=150>" & rs("BaseAVNumber") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD width=250> <input style=""width:250; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDescription" & rs("DefinitionID") & " name=txtGPGDescription" & rs("DefinitionID") & " value=""" & trim(rs("BaseGPGDescription")) & " " &  trim(rs("SeriesNum")) & """ onchange=""GpgDescription_onchange(" & rs("DefinitionID") & ")""></TD>"
                strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
				strBase = strBase &  "</tr>"
				
				'if trim(rs("AVNumber")) = "" And trim(rs("ImageAvId") > "") then
				'   strRows = strRows & "<TD width=175>PENDING</TD>"
				'else
				    strRows = strRows & "<TD width=150>" & rs("AVNumber") & "&nbsp;</TD>"
				'end if
				
				strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
				strRows = strRows & "<TD width=250><span class=gpgDescription" & rs("DefinitionId") & ">" & trim(rs("BaseGPGDescription")) & "&nbsp;" &  trim(rs("SeriesNum")) & "</span>&nbsp;" & rs("CountryCode") & "</TD>"
				'strRows = strRows &  "<TD width=200> <input style=""width:200; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDesciption name=txtGPGDesciption value=""" & rs("GPGDescription") & """></TD>"
				strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
				'if trim(rs("OtherLanguage") & "") <> "" then
				'  if request("Type") = "1" then
				'    strRows = strRows & "<TD class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
				'  else
				'	strRows = strRows & "<TD ID=""Row" & rs("ID") & """ bgcolor=" & strCellColor & " width=70><a hidefocus class=""" & strLanguageList & """ href=""javascript:ChooseLanguages(" & rs("ID") & ",'" & strlanguageList & "');"" id=""Lang" & rs("ID") & """>" & strSavedLanguages & "</a></TD>"
				 ' end if
				'else
					strRows = strRows & "<TD class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
				'end if				
				'strRows = strRows & "<TD width=70>" & strLanguageList & "</TD>"
				strRows = strRows & "<TD>&nbsp;" & rs("Keyboard") & "</TD>"
				strRows = strRows & "<TD>" & "&nbsp;" & rs("Powercord") & "</TD>"
				strRows = strRows & "</tr>"
	
			end if
			rs.MoveNext
		loop
		if imagecount = 0 then
			strAllImages = "<tr><td colspan=10><font size=1 face=verdana>No active images defined for this product.</font></td></tr>"
		end if
		strAllImages = strAllImages & strBase
		strAllImages = strAllImages &  "<tr style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV No.</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Config</b></td><td><b>Lang</b></td><td><b>KBD</b></td><td><b>Cord</b></td></thead>" & strRows & "</table><br></td></tr>"
		strRows = ""
		

	end if	
	rs.Close

	if TotalImageDefsChecked > 0 then
		strAllChecked="checked"
	else
		strAllChecked=""
	end if

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

<h3 align="center"><TH align=center>Select Images To Localize</TH></h3>
<!--<b><font size=2 face=verdana><%=strDeliverable & " (" & strProdName & ")"%></font></b>-->

<form ID=frmChange method=post>

<input id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<!--<font size=2 face=verdana color=green>Note: Inactive and Released images are not displayed on this page.</font><BR><BR>-->
<b><font size=2 face=verdana>Images:</font><BR></b>

		<table id="ImageTable" class="ImageTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			<thead bgcolor=Wheat>
				<!--<% if request("Type") = "1" then%>
				<TH align=left><input disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				<%else%>-->
				<th align=left><input type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TD>
                <!--<%end if%>-->
				<th align=left>ID</th>
				<th align=left>Base AV No</th>
				<th align=left>ZWAR</th>
				<th align=left>Base GPG Description</th>
				<th align=left>OS</th>
				<th align=left>Apps&nbsp;Bundle</th>
				<!--<TH align=left>BTO/CTO</TH>
				<TH align=left>Comments</TH>-->
			</thead>
			<%=strAllImages%>
			

		</table>
<input style="Display:none" type="checkbox" id=chkAllChecked name=chkAllChecked>
<input type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<input type="hidden" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
<input type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<input type="hidden" id=txtSuccess name=txtSuccess value="<%= txtSuccess %>">
</form>


<%
	set rs= nothing
	set cn=nothing
%>


</body>
</HTML>
