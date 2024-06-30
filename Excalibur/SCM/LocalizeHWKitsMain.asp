
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
            arSelectedHWKit = split(Request.Form(item),",")
            For j = lbound(arSelectedHWKit) To ubound(arSelectedHWKit)
                SelectedHWKit = trim(arSelectedHWKit(j))         
                arData = split(SelectedHWKit,"|")
                'If strUserID = 5016 Or strUserID = 8 Or strUserID = 1396 Then
                    RowsEffected = 0
                    set cm = server.CreateObject("ADODB.connection")
	                cm.ConnectionString = Session("PDPIMS_ConnectionString")
	                cm.Open
	                cm.BeginTrans()
	                cm.Execute "usp_InsertLocalizedHWKit " & clng(request("PVID")) & "," & clng(request("BID")) & ",'" & ScrubSQL(BaseGPG) & "','" & clng(BaseID) & "','" & ScrubSQL(strUserName) & "','" & ScrubSQL(arData(2)) & "','" & ScrubSQL(arData(1)) & "','" & ScrubSQL(arData(3)) & "'," & clng(arData(4)) , RowsEffected 
	                If RowsEffected <> 1 Then
	                    cm.RollbackTrans()
	                    ErrorCount = ErrorCount + 1
                        'Response.Write "PVID(" & request("PVID") & ") BID(" & request("BID") & ") baseGPG(" & BaseGPG & ") RootID(" & BaseID & ") CountryCode(" & arData(1) & ") ConfigCode(" & arData(2) & ") UserName(" & strUserName & ") </br>" 
	                    'response.Write "RowsEffected:" & RowEffected & "<br>"
	                Else 
	                    cm.CommitTrans()
	                    SuccessCount = SuccessCount + 1
	                End If
	                set cm=nothing
                'Else
                '    Response.Write "PVID(" & request("PVID") & ") BID(" & request("BID") & ") baseGPG(" & BaseGPG & ") RootID(" & BaseID & ") CountryCode(" & arData(1) & ") ConfigCode(" & arData(2) & ") UserName(" & strUserName & ") </br>" 
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
            window.parent.opener = 'X';
            window.parent.open('', '_parent', '')
           window.parent.close();    
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


function chkBase_onclick(RootID) {
	var i;
	var checkBoxes = document.getElementsByName("chkRow" + RootID);
	var chkBase = document.getElementById(RootID);
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


function chkRow_onclick(){
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

	if (document.all("ChildRow" + strID).style.display == "" )
		document.all("ChildRow" + strID).style.display="none";	
	else
		document.all("ChildRow" + strID).style.display="";	
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
	dim LastRootID
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
	strAllChecked = "checked"
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

	rs.open "spListHWKitsForLocalization " & clng(request("PVID")) & "," & clng(request("BID")) & ",'"  & ScrubSQL(request("strSeriesSummary")) & "'",cn,adOpenForwardOnly
	LastRootID = ""
	if rs.EOF and rs.BOF then
		strAllRows = "<tr><td colspan=10><FONT size=1 face=verdana>No AC Adapters defined for this product.</font></td></tr>"
	else
		dim imagecount
		imagecount=0
		
		do while not rs.EOF
				imagecount = imagecount + 1
				if LastRootID <> rs("RootID") and LastRootID <> "" then
			
					strAllRows = strAllRows & strBase
					strAllRows = strAllRows &  "<tr style=""Display:none"" id=""ChildRow" & LastRootID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><td><b>Dash</b></td><td><b>Localized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td></thead>" & strRows & "</table><br></td></tr>"
					strRows = ""
					YesCount = 0
					NoCount = 0
					MixedCount=0
				end if
				LastRootID = rs("RootID")
				
				strCellColor = "ivory"
				
				if rs("AvId") = "" Or IsNull(rs("AvId")) then
				    strRows = strRows & "<TR bgcolor=ivory><TD><input class=""" & trim(rs("RootID")) & """ type=""checkbox"" unchecked id=chkRow" & rs("RootID") & " name=""chkRow" & rs("RootID") & """ value=""" & rs("RootID") & "|" & rs("CountryCode") & "|" & rs("OptionConfig") & "|" & rs("AvParentID") & "|" & rs("GeoID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkRow_onclick()""></TD>"
					NoCount = NoCount + 1
				else
					strCellColor = "ivory"
					YesCount = YesCount + 1
					strRows = strRows & "<TR bgcolor=ivory style=""color:grey""><TD><input Disabled class=""" & trim(rs("RootID")) & """ type=""checkbox"" checked id=chkRow" & rs("RootID") & " name=""chkRow" & rs("RootID") & """ value=""" & rs("RootID") & " LANGUAGE=javascript onclick=""return chkRow_onclick()"" style=""WIDTH:16;HEIGHT:16""></TD>"
				end if
				
			    if YesCount = 0 then' and MixedCount=0 
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("RootID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("RootID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("RootID") & ")"" onclick=""return BaseRow_onclick(" & rs("RootID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("RootID") & """ name=""Base"" value=""" & rs("RootID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("RootID") & ")""></TD>" 
				elseif NoCount=0 then'and MixedCount=0  
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("RootID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("RootID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("RootID") & ")"" onclick=""return BaseRow_onclick(" & rs("RootID") & ")"" bgcolor=cornsilk><TD><input Disabled id=""" & rs("RootID") & """ name=""Base"" value=""" & rs("RootID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("RootID") & ")""></TD>" 
				    TotalImageDefsChecked = TotalImageDefsChecked + 1
				else
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("RootID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("RootID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("RootID") & ")"" onclick=""return BaseRow_onclick(" & rs("RootID") & ")"" bgcolor=cornsilk><TD><input id=""" & rs("RootID") & """ name=""Base"" value=""" & rs("RootID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick(" & rs("RootID") & ")"" indeterminate=-1></TD>"
				    TotalImageDefsChecked = TotalImageDefsChecked + 1
				end if
				strBase = strBase &  "<TD width=350>" & rs("Description") & "</TD>"
				strBase = strBase &  "<TD width=250>" & rs("AvBaseNo") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD width=250> <input style=""width:250; font-size:10; font-family:Verdana"" maxlength=40 type=text id=txtGPGDescription" & rs("RootID") & " name=txtGPGDescription" & rs("RootID") & " value=""" & left(trim(rs("BaseGPGDescription")) & " " &  trim(rs("SeriesNum")),40) & """ onchange=""GpgDescription_onchange(" & rs("RootID") & ")""></TD>"
                'strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
				'strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
				strBase = strBase &  "</tr>" & vbcrlf
				
				strRows = strRows & "<TD width=150>" & rs("AVNo") & "&nbsp;</TD>"
		        strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
		        
				strRows = strRows & "<TD width=250><span class=gpgDescription" & rs("RootId") & ">" & trim(rs("BaseGPGDescription")) & " " &  trim(rs("SeriesNum")) & "</span>&nbsp;" & rs("CountryCode") & "</TD>"
				'strRows = strRows &  "<TD width=200> <input style=""width:200; font-size:10; font-family:Verdana"" maxlength=50 type=text id=txtGPGDesciption name=txtGPGDesciption value=""" & rs("GPGDescription") & """></TD>"
				
				strRows = strRows & "<TD>" & rs("CountryCode") & "<input type=hidden name=CountryCode" & rs("RootID") & " value=""" & rs("CountryCode") & """></TD>"
				'strRows = strRows & "<TD>" & rs("CountryCode") & "</TD>"
				
				strRows = strRows & "<TD>" & rs("OptionConfig") & "<input type=hidden name=OptionConfig" & rs("RootID") & " value=""" & rs("OptionConfig") & """></TD>"
				'strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
				
				strRows = strRows & "<TD>" & rs("Keyboard") & "</TD>"
				strRows = strRows & "<TD>" & rs("Powercord") & "</TD>"
				strRows = strRows & "</tr>" & vbcrlf
			rs.MoveNext
		loop
		if imagecount = 0 then
			strAllRows = "<tr><td colspan=10><font size=1 face=verdana>No active AC Adapters defined for this product.</font></td></tr>"
		end if
		strAllRows = strAllRows & strBase
		strAllRows = strAllRows & "<tr style=""Display:none"" id=""ChildRow" & LastRootID & """ bgcolor=cornsilk ><td>&nbsp;</td><td colspan=7><br><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=ChildRows border=0><thead bgcolor=wheat><td>&nbsp;</td><td><b>AV Number</b></td><Td><b>Dash</b></td><td><b>Loacalized GPG Description</b></td><td><b>Country Code</b></td><td><b>Config</b></td><td><b>KBD</b></td><td><b>Cord</b></td></thead>" & strRows & "</table><br></td></tr>"
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
<h3 align="center"><TH align=center>Select Hardware Kits To Localize</TH></h3>
<BR>
<h5 align="center"><TH align="center">Please select only those base AVs that are designated for the selected brand (KMAT) of the product in the PDD.</TH></h5>
 
<!--<b><font size=2 face=verdana><%=strDeliverable & " (" & strProdName & ")"%></font></b>-->

<form ID=frmChange method=post>
<input id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<!--<font size=2 face=verdana color=green>Note: Inactive and Released images are not displayed on this page.</font><BR><BR>-->
<b><font size=2 face=verdana>AC Adapters:</font><BR></b>

		<table id="HWKitTable" class="HWKitTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			<thead bgcolor=Wheat>
				<!--<% if request("Type") = "1" then%>
				<TH align=left><input disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				<%else%>-->
				<th align=left><input type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TD>
                <!--<%end if%>-->
				<th align=left>Description</th>
				<th align=left>Base AV Number</th>
				<!--<th align=left>ZWAR</th>-->
				<th align=left>Base GPG Description</th>
				<!--<th align=left>OS</th>
				<th align=left>Apps&nbsp;Bundle</th>
				<TH align=left>BTO/CTO</TH>
				<TH align=left>Comments</TH>-->
			</thead>
			<%=strAllRows%>
			

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
