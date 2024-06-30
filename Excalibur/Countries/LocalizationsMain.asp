<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/bundleConfig.inc" -->
<html>
<head>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--
function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function chkLocalization_onclick(strID){
}
function ClosePropertiesDialog(strID) {
    $("#iframeDialog").dialog("close");
    if (typeof (strID) != "undefined") window.location.reload(true);
}
function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {
   
    $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
    $("#modalDialog").attr("width", "95%");
    $("#modalDialog").attr("height", "90%");
    $("#modalDialog").attr("src", QueryString);
    $("#iframeDialog").dialog("option", "title", Title);
    $("#iframeDialog").dialog("open");
}
function EditRelease_onclick(ProdBrandCountryLocalizationID) {
    
   // var sFusionRequirements = document.getElementById("inpFusionRequirements").value;
    ShowPropertiesDialog('LocalizationReleaseEdit.asp?ProdBrandCountryLocalizationID=' + ProdBrandCountryLocalizationID + '&FusionRequirements=<%=Request("FusionRequirements")%>&pvID=<%=Request("pvID")%>', "Edit Release", 700, 300);
        
}
function EditRelease_onclick_Notchecked(LocalizationID) {

    // var sFusionRequirements = document.getElementById("inpFusionRequirements").value;
    ShowPropertiesDialog('LocalizationReleaseEdit.asp?pvID=<%=Request("pvID")%>&ProdBrandCountryID=<%=request("ID")%>&LocalizationID=' + LocalizationID, "Edit Release", 700, 400);

}
//-->
</script>
<STYLE>
<!--
.TextBox
{
	font-family: Verdana;
	font-size: xx-small;
	height: 16;
	width: 60;
	border: solid 1px gray;
}
.CheckBox
{
	width:16;
	height:16;
}
.Hidden
{
	display:none;
}	
.TableCell
{
	border-top: gray thin solid;
	font-size: xx-small;
	font-family: Verdana;
	vertical-align: top;
}
.Table
{
	border-bottom: solid thin gray;
}
.TableHeader
{
	border-top: gray thin solid;
	font-weight: bold;
	font-size: xx-small;
	vertical-align: middle;
	line-height: 15px;
	font-family: Verdana;
}

//-->
</STYLE>
</head>
<body bgcolor="white">
<form id="frmCountries" action="LocalizationsSave.asp?pulsarplusDivId=<%= Request("pulsarplusDivId") %>" method="post">
<%	

	dim cn 
	dim rs
	dim cm
	dim p
	dim i
	dim strLastRegion
	dim strRegion
	dim strRegionID
	dim strProdPartner
	dim CurrentUserPartner
	dim strTAB
	dim strCons
	dim strEnt
    dim strFusionRequirements
        strFusionRequirements=Request("FusionRequirements")

if request("ID") = "" then
	Response.Write "<font face=verdana size=3><b>Choose Localizations</b></font><BR><BR>"
	Response.Write "<BR><font size=2 face=verdana>Not enough information to display this page</font>"
else
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	dim strLanguage
	
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


	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=1"
	else
		CurrentUserPartner = rs("PartnerID")
	end if 
	rs.Close
		


	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	'cm.CommandText = "spGetCountryForProductCountry"
	'Set p = cm.CreateParameter("@ID", 3, &H0001)
	'p.Value = request("ID")
	'cm.Parameters.Append p
	
	cm.CommandText = "usp_SelectBrandCountryLocalizationData2"
	
	Set p = cm.CreateParameter("@p_ProdBrandCountryID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p



	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strProdPartner) <> trim(CurrentUserPartner) then
			set rs = nothing
			set cn=nothing
			
			Response.Redirect "../NoAccess.asp?Level=1"
		end if
	end if

    dim strHeader : strHeader = ""

	if not (rs.EOF and rs.BOF) then
		Dim PDDLocked
		PDDLocked = CheckPDDLockBit(rs("ProductVersionID"))

		If PDDLocked = -500 Then
			Response.Write "<font face=verdana size=3><b>Product Release Not Found</b></font><BR><BR>"
			Response.End
		ElseIf PDDLocked Then
			'Display the DCR List
			Response.Write "<p><font face=verdana size=2><b>Choose the Applicable DCR</b></font><br>"
			Response.Write "<SELECT id=cboDcr name=cboDcr>"
			FillDcrList(rs("ProductVersionID"))
			Response.Write "</SELECT></p>"
		End If

		Response.Write "<font face=verdana size=2><b>Choose Localizations for " & rs("Country") & "</b></font><BR>"
		strProdPartner=rs("Partnerid")

		Response.Write "<Table class=Table width=100% border=0 cellpadding=1 cellspacing=0>"	
		Response.Write "<TR bgcolor=cornsilk>" & _
			"<TD class=TableHeader>&nbsp;</TD>" & _
            "<TD class=TableHeader>ID</TD>" & _
        	"<TD class=TableHeader>Localization</TD>"        
        if StrFusionRequirements="1" then
            Response.Write "<TD class=TableHeader>Releases</TD>" 
        end if

        strHeader = "<TD class=TableHeader>Code</TD>" & _
			        "<TD class=TableHeader>Dash</TD>" & _ 
			        "<TD class=TableHeader>Lang</TD>" & _
			        "<TD class=TableHeader>Keyboard</TD>" & _
			        "<TD class=TableHeader>KWL</TD>"

        if rs("PowerCordSupported") = true then
            strHeader = strHeader & "<TD class=TableHeader>Power&nbsp;Cord</TD>"
        end if

        if rs("DuckheadPowerCordSupported") = true then
            strHeader = strHeader & "<TD class=TableHeader>Duckhead&nbsp;Power&nbsp;Cord</TD>"
        end if

        if rs("DuckheadSupported") = true then
            strHeader = strHeader & "<TD class=TableHeader>Duckhead</TD>"
        end if

		strHeader = strHeader &	"<TD class=TableHeader>Doc Kit</TD>" & _
			                    "<TD class=TableHeader>OS Restore</TD>" & _
                                "<TD class=TableHeader>Comments</TD>" & _
			                    "</TR>"
        Response.Write strHeader

		do while not rs.EOF
			strLanguage = "<u>" & rs("OSLanguage") & "</u>"
			if trim(rs("OtherLanguage") & "") <> "" then
				strLanguage = strLanguage & "," & rs("OtherLanguage")
			end if
			if trim(rs("BrandCountryLocalizationID") & "") = "" And rs("Active") then
				Response.Write "<TR bgcolor=Ivory ID=Row" & rs("LocalizationID") & ">" & _
					"<TD valign=top style=""BORDER-TOP: gray thin solid"">"  & _
					"<INPUT value=""" & rs("LocalizationID") & """ class=""CheckBox"" type=""checkbox"" id=chkSelected name=chkSelected LANGUAGE=javascript onclick=""return chkLocalization_onclick(" & rs("LocalizationID")&  ")"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ class=""Hidden"" type=""checkbox"" id=chkTag name=chkTag>" & _
					"</td>"
		    elseif trim(rs("BrandCountryLocalizationID") & "") = "" And not rs("Active") then
				Response.Write "<TR bgcolor=Grey ID=Row" & rs("LocalizationID") & ">" & _
					"<TD valign=top style=""BORDER-TOP: gray thin solid"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ class=""CheckBox"" type=""checkbox"" id=chkSelected name=chkSelected LANGUAGE=javascript onclick=""return chkLocalization_onclick(" & rs("LocalizationID")&  ")"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ class=""Hidden"" type=""checkbox"" id=chkTag name=chkTag>" & _
					"</td>"
		    elseif not rs("Active") then
				Response.Write "<TR bgcolor=Grey ID=Row" & rs("LocalizationID") & ">" & _
					"<TD valign=top style=""BORDER-TOP: gray thin solid"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ checked class=""CheckBox"" type=""checkbox"" id=chkSelected name=chkSelected  LANGUAGE=javascript onclick=""return chkLocalization_onclick(" & rs("LocalizationID")&  ")"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ checked class=""Hidden"" type=""checkbox"" id=chkTag name=chkTag>" & _
					"</td>"
			else
				Response.Write "<TR bgcolor=Ivory ID=Row" & rs("LocalizationID") & ">" & _
					"<TD valign=top style=""BORDER-TOP: gray thin solid"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ checked class=""CheckBox"" type=""checkbox"" id=chkSelected name=chkSelected  LANGUAGE=javascript onclick=""return chkLocalization_onclick(" & rs("LocalizationID")&  ")"">" & _
					"<INPUT value=""" & rs("LocalizationID") & """ checked class=""Hidden"" type=""checkbox"" id=chkTag name=chkTag>" & _
					"</td>"
			end if

            Response.Write "<TD class=TableCell >" & rs("LocalizationID") & "</TD>"
			Response.Write "<TD class=TableCell >" & rs("DisplayName") & "</TD>"
            if StrFusionRequirements="1" then
                dim sReleasetext, sReleaseURL
                sReleasetext = rs("Releases")
                if sReleasetext ="" then
                    sReleasetext = "Add"
                   
                end if
                sReleaseURL = "<a href='javascript:EditRelease_onclick(" & rs("BrandCountryLocalizationID") & ");'>"
                if trim(rs("BrandCountryLocalizationID") & "") = "" then
                     sReleaseURL = "<a href='javascript:EditRelease_onclick_Notchecked(" & rs("LocalizationID") & ");'>"
                end if
                
		        Response.Write "<TD class=TableCell >" & sReleaseURL & sReleasetext & "&nbsp;" & "</a>" & "</TD>"
            end if
			Response.Write "<TD class=TableCell >" & rs("OptionConfig") & "&nbsp;</TD>"
			Response.Write "<TD class=TableCell >" & rs("Dash") & "&nbsp;</TD>"
			Response.Write "<TD class=TableCell >" & strLanguage & "&nbsp;</TD>"
			
			If trim(rs("ModifiedKeyboard") & "") > ""  then
			    Response.Write "<TD class=TableCell >" & rs("Keyboard") & "<BR>" & _
				"<input name=Kbd" & rs("LocalizationID") & " type=text maxlength=15 class=TextBox value=" & rs("ModifiedKeyboard") & ">" & _
				"<input name=hidKbd" & rs("LocalizationID") & " type=text class=Hidden value=" & rs("ModifiedKeyboard") & "></TD>"
			Else
			    Response.Write "<TD class=TableCell >" & rs("Keyboard")& "&nbsp;</TD>"
			End If
			
			If trim(rs("ModifiedKWL") & "") > ""  then
			    Response.Write "<TD class=TableCell>" & rs("KWL") & "<BR>" & _
				"<input name=KWL" & rs("LocalizationID") & " type=text maxlength=7 class=TextBox value=" & rs("ModifiedKWL") & ">" & _
				"<input name=hidKWL" & rs("LocalizationID") & " type=text class=Hidden value=" & rs("ModifiedKWL") & "></TD>"
			Else
			    Response.Write "<TD class=TableCell >" & rs("KWL")& "&nbsp;</TD>"
			End If
			
            if rs("PowerCordSupported") = true then
			    If trim(rs("ModifiedPowerCord") & "") > ""  then
			       Response.Write "<TD class=TableCell >" & rs("PowerCord") & "<BR>" & _
				    "<input name=PwrCord" & rs("LocalizationID") & " type=text maxlength=20 class=TextBox value=" & rs("ModifiedPowerCord") & ">" & _
				    "<input name=hidPwrCord" & rs("LocalizationID") & " type=text class=Hidden value=" & rs("ModifiedPowerCord") & "></TD>"
			    Else
			        Response.Write "<TD class=TableCell >" & rs("PowerCordGEO")& "&nbsp;</TD>"
			    End If
			end if

            if rs("DuckheadPowerCordSupported") = true then
                Response.Write "<TD class=TableCell >" & rs("DuckheadPowerCordGEO")& "&nbsp;</TD>"
            end if

            if rs("DuckheadSupported") = true then
                Response.Write "<TD class=TableCell >" & rs("DuckheadGEO")& "&nbsp;</TD>"
            end if

			If trim(rs("ModifiedDocKits") & "") > ""  then
			  Response.Write "<TD class=TableCell >" & rs("DocKits") & "<BR>" & _
				"<input name=DocKit" & rs("LocalizationID") & " type=text maxlength=30 class=TextBox value=" & rs("ModifiedDocKits") & ">" & _
				"<input name=hidDocKit" & rs("LocalizationID") & " type=text class=Hidden value=" & rs("ModifiedDocKits") & "></TD>"
			Else
			    Response.Write "<TD class=TableCell >" & rs("DocKits")& "&nbsp;</TD>"
			End If
			
			If trim(rs("ModifiedRestoreMedia") & "") > ""  then
			  Response.Write "<TD class=TableCell >" & rs("RestoreMedia") & "&nbsp;<BR>" & _
				"<input name=Media" & rs("LocalizationID") & " type=text maxlength=25 class=TextBox value=" & rs("ModifiedRestoreMedia") & ">" & _
				"<input name=hidMedia" & rs("LocalizationID") & " type=text class=Hidden value=" & rs("ModifiedRestoreMedia") & "></TD>"
			Else
			    Response.Write "<TD class=TableCell >" & rs("RestoreMedia")& "&nbsp;</TD>"
			End If
			
            Response.Write "<TD class=TableCell >" & rs("Comments") & "&nbsp;</TD>"
			rs.Movenext
		loop
		Response.Write "</Table>"	
	else
		Response.Write "<font face=verdana size=3><b>Choose Localizations</b></font><BR><BR>"
		strProdPartner = ""
		Response.Write "<font size=2 face=verdana>No localizations defined for this country.</font>"
	end if
	rs.Close
	
	set rs = nothing
	cn.Close
	set cn = nothing
end if


  %>
  <input type="hidden" id="txtID" name="txtID" value="<%= request("ID")%>" />
    <div style="display: none;">
        <div id="iframeDialog" title="Coolbeans">
            <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
        </div>
    </div>
</form>
</body>
</html>
<%
Function CheckPDDLockBit(ProductVersionID)
	Dim cn, cmd, p, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	Set cm = Server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "usp_GetPDDLockStatus"

	Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
	p.Value = ProductVersionID
	cm.Parameters.Append p

	Set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorType = adOpenStatic
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	Set cm = nothing
	Set p = nothing

	If rs.EOF and rs.BOF Then
		CheckPDDLockBit = -500
	Else
		CheckPDDLockBit = rs("PDDLocked")
	End If
	
	rs.Close
	
	Set rs = nothing
	Set cn = nothing
	
End Function

Sub FillDcrList(ProductVersionID)
	Dim Result
	Dim cn, cmd, p, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	Set cm = Server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListApprovedDCRs"

	Set p = cm.CreateParameter("@ProdID", 3, &H0001)
	p.Value = ProductVersionID
	cm.Parameters.Append p

	Set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorType = adOpenStatic
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	Set cm = nothing
	Set p = nothing

	Response.Write "<option value=""0"">-- Please Select A DCR --</option>"					
	
	Do until rs.eof
		Response.Write "<option value=""" & rs("ID") & """>" & rs("ID") & ":" & server.HTMLEncode(rs("Summary")) & "</option>"					
		rs.movenext
	Loop

	rs.close

	set cn = nothing
	set cm = nothing
	set rs = nothing
	set p  = nothing
End Sub
%>