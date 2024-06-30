<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<%

Response.AddHeader "Pragma", "No-Cache"
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim m_Table
Dim m_ListBox
Dim m_IsSysAdmin
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_EditModeOn
Dim m_JSArray
Dim m_ProductVersionID
dim m_IsPulsarProduct : m_IsPulsarProduct = 0
dim m_Releases
dim Security, sUserFullName, CurrentUserID

m_IsSysAdmin = false
m_IsProgramManager = false
m_IsSysEngProgramManager = false
m_IsSysTeamLead = false
m_EditModeOn = false
m_ProductVersionID = Trim(Request("PVID"))
m_IsPulsarProduct = Request("IsPulsarProduct")
m_Releases = ""


Call Main()

Sub Main()

	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsProgramManager = Security.IsProgramManager(m_ProductVersionID)
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(m_ProductVersionID)
	m_IsSysTeamLead = Security.IsSystemTeamLead(m_ProductVersionID)
	sUserFullName = Security.CurrentUser()
    CurrentUserID = Security.CurrentUserID()
  
    Set Security = Nothing

    'Check if user has Regions.Edit Permission
    Dim HasPemission
    HasPemission = 0
    dim cm, rs, p, cn

    set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.CommandTimeout =120
		cn.Open
    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")
    cm.CommandType = 4
	cm.CommandText = "usp_USR_ValidatePermission"
	
	Set p = cm.CreateParameter("@p_intUserId", 200, &H0001, 15)
	p.Value = CurrentUserID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@p_PName", 200, &H0001, 100)
	p.Value = "Regions.Edit"
	cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
    if not(rs.EOF and rs.BOF) then
		HasPemission = rs("HasPermission")
	end if	  
    rs.Close 
    
   if (HasPemission = 1) then
        m_EditModeOn = True
    end if

  set rs = nothing
  cn.close
  set cn= nothing


	If ucase(request("SaveMode")) = ucase("true") Then
		SaveChanges
	Else
		DrawScreen
	End If
End Sub

Sub SaveChanges()
	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Dim dw
	Dim cn
	Dim cmd
	Dim RecordCount
	Dim ReturnValue

	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
        
	Set cmd = dw.CreateCommandSP(cn, "usp_CopyLocalizationSettings")
	dw.CreateParameter cmd, "@p_SrcProductBrandID", adInteger, adParamInput, 8, Request.Form("ddlBrand")
	dw.CreateParameter cmd, "@p_DestProductBrandID", adInteger, adParamInput, 8, Request.Form("BID")
	dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
    dw.CreateParameter cmd, "@p_SrcProductVersionID", adInteger, adParamInput, 8, Request.Form("ddlProducts")
    dw.CreateParameter cmd, "@p_DestReleases", adVarChar, adParamInput, 250, Request.Form("chkCopyReleases")
    RecordCount = dw.ExecuteNonQuery(cmd)
			
	Set cmd = nothing

	'Response.End
	Response.Write "<input type=hidden id=myReturnValue value=""1"">"
	
End Sub

Sub DrawScreen()
	If Len(Trim(Request("PVID"))) = 0 Or Len(Trim(Request("BID"))) = 0 Then
		Response.Write "<H3>Insufficient information provided.</H3>"
		Response.End
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3>"
		Response.End
	End If

	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_ListProductVersionBrands")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, 0 'pass ProductBrand zero to include target product to the list of source product list.
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
	If rs.eof And rs.bof Then
		Response.Write "<H3>No Active Products were found.</H3>"
		Response.End
	End If
	
	Dim Rows
	Dim ReturnCode
	Dim ProductVersionID
	Dim sProductOptions
	Dim sLastProductVersion	

	Do Until rs.EOF
		'
		' Create Javascrip Array
		'
		If sLastProductVersion <> rs("ProductVersionID") Then
			sLastProductVersion = rs("ProductVersionID")
			m_JSArray = m_JSArray & "'EOF');" & vbcrlf & "assocArray['ddlProducts=" & rs("ProductVersionID") & "'] = new Array("
			sProductOptions = sProductOptions & "<OPTION Value=" & rs("ProductVersionID") 
			If CLng(sLastProductVersion) = CLng(Request("PVID")) Then
				sProductOptions = sProductOptions & " SELECTED"
			End If
			sProductOptions = sProductOptions & ">" & Trim(rs("ProductName")) & " " & Trim(rs("ProductVersion")) & "</OPTION>"
		End If
		m_JSArray = m_JSArray & "'" & rs("ProductBrandID") & "','" & Replace(rs("BrandName"),"'","\'") & "',"
		
		rs.movenext
	Loop
	
	m_JSArray = Right(m_JSArray, LEN(m_JSArray) - 7) & "'EOF');"   
	m_JSArray = "<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>" & vbcrlf & _
		"<!--" & vbcrlf & _
		"var assocArray = new Object();" & vbcrlf & _
		m_JSArray & vbcrlf & _
		"//-->" & vbcrlf & _
		"</SCRIPT>"		
    
	m_ListBox = sProductOptions
	
	rs.close
    Set cmd = nothing
    Set cmd = dw.CreateCommandSP(cn, "usp_Product_GetProductReleases")
	dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, m_ProductVersionID
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
    do while not rs.eof
        if trim(rs("ReleaseCount")) = "1" then
            m_Releases = "<input checked id=""chkCopyReleases"" name=""chkCopyReleases"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """  > " & rs("ReleaseName") & "&nbsp;"                                    
        else
            m_Releases = m_Releases & "<input id=""chkCopyReleases"" name=""chkCopyReleases"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """   > " & rs("ReleaseName") & "&nbsp;"
        end if
        rs.movenext
    loop

    rs.close
	set rs = nothing
	
End Sub

%>
<HTML>
<HEAD>
<TITLE>Copy Main</TITLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
<!--

function DocumentOnLoad()
{
	if (window.myReturnValue)
	{
	    var pulsarplusDivId = document.getElementById('hdnTabName');
	    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	       // parent.window.parent.reloadFromPopUp(pulsarplusDivId);
	        // For Closing current popup if Called from pulsarplus
	        parent.window.parent.closeExternalPopup();
	    }
	    else {
	        if (parent.window.parent.document.getElementById('modal_dialog')) {
	            parent.window.parent.modalDialog.cancel(true);
	        } else {
	            window.parent.close();
	        }
	    }
	    /*window.returnValue = window.myReturnValue.value;
		this.close();*/
	}

	listboxItemSelected(frmMain.ddlProducts,frmMain.ddlBrand);
}

function comboItemSelected(oList1,oList2)
{
	if (oList2!=null)
	{
		clearComboOrList(oList2);
		if (oList1.selectedIndex == -1)
			oList2.options[oList2.options.length] = new Option('Please make a selection from the list', '');
		else 
			fillCombobox(oList2, oList1.name + '=' + oList1.options[oList1.selectedIndex].value);
	}
}

function listboxItemSelected(oList1,oList2)
{
	if (oList2!=null)
	{
		clearComboOrList(oList2);
		if (oList1.selectedIndex == -1)
			oList2.options[oList2.options.length] = new Option('Please make a selection from the list', '');
		else 
			fillListbox(oList2, oList1.name + '=' + oList1.options[oList1.selectedIndex].value);
	}
}

function clearComboOrList(oList)
{
	for (var i = oList.options.length - 1; i >= 0; i--)
	{
		oList.options[i] = null;
	}
	oList.selectedIndex = -1;
	if (oList.onchange)	oList.onchange();
}

function fillCombobox(oList, vValue)
{
	if (vValue != '') 
	{
		if (assocArray[vValue])
		{
			oList.options[0] = new Option('Please make a selection', '');
			var arrX = assocArray[vValue];
			for (var i = 0; i < arrX.length; i = i + 2)
			{
				if (arrX[i] != 'EOF') oList.options[oList.options.length] = new Option(arrX[i + 1].split('&amp;').join('&'), arrX[i]);
			}
			if (oList.options.length == 1)
			{
				oList.selectedIndex=0;
				if (oList.onchange) oList.onchange();
			}
		} 
		else 
		{
			oList.options[0] = new Option('None found', '');
		}
	}
}

function fillListbox(oList, vValue)
{
	if (vValue != '') 
	{
		if (assocArray[vValue])
		{
			oList.options[oList.options.length] = new Option('-- Make A Selection --');
			var arrX = assocArray[vValue];
			for (var i = 0; i < arrX.length; i = i + 2)
			{
				if (arrX[i] != 'EOF') oList.options[oList.options.length] = new Option(arrX[i + 1].split('&amp;').join('&'), arrX[i]);
			}
			if (oList.options.length == 1)
			{
				oList.selectedIndex=0;
				if (oList.onchange)	oList.onchange();
			}
		} 
		else 
		{
			oList.options[0] = new Option('None found', '');
		}
	}
}

function cmdAllRelease_onclick() {
    var checkboxes = document.getElementsByName('chkCopyReleases');
    for (var i = 0; i < checkboxes.length; i++ )
        checkboxes[i].checked = true;
}

//-->
</SCRIPT>
<%= m_JSArray %>
<STYLE>
<!--
.Region
{
    FONT-FAMILY: Verdana;
    FONT-WEIGHT: bold;
    FONT-SIZE: xx-small;
    COLOR: black;
    BACKGROUND-COLOR: MediumAquamarine;
    border-top: black thin solid;
}
.TextBox
{
	font-family: Verdana;
	font-size: xx-small;
	height: 16;
	width: 160;
	border: solid 1px gray;
}
//-->
</STYLE>
<LINK rel="stylesheet" type="text/css" href="../style/general.css">
</HEAD>
<BODY bgcolor="ivory" onload="DocumentOnLoad()">
<FORM id="frmMain" method="post" action=CopyMain.asp >
<p>Please select the Product and Brand to import from.</p>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><th>Products</th><th>&nbsp;&nbsp;</th><th>Brand</th></tr>
<tr><td>
<SELECT class=TextBox id=ddlProducts name=ddlProducts onChange='listboxItemSelected(this.form.ddlProducts,this.form.ddlBrand);'>
<%= m_ListBox%>
</SELECT>
</td><td>&nbsp;</td><td>
<SELECT class=TextBox id=ddlBrand name=ddlBrand></SELECT>
</td></tr>
</table>
<%if request("IsPulsarProduct") = "1" then %>
<div style="margin-top:100px;">
    <p>Import To:</p>
    <table width=100% border=0 cellspacing=0 cellpadding=0>    
    <tr>
        <td><b>Releases:</b><font color=red>*</font>&nbsp&nbsp<%=m_Releases%>&nbsp&nbsp<input type="button" value=" All " id="cmdAllRelease" name="cmdAllRelease" onclick="return cmdAllRelease_onclick()"></td>
    </tr>
    </table>
</div>
<% end if %>
<input type=hidden id=SaveMode name=SaveMode value=true >
<input type=hidden id=bid name=bid value="<%= Request("bid")%>">
<input type=hidden id=pvid name=pvid value="<%= Request("pvid")%>">
<input type=hidden id=hidEdit name=hidEdit value="<%= m_EditModeOn%>">
<INPUT type="hidden" id=IsPulsarProduct name=IsPulsarProduct value="<%= request("IsPulsarProduct")%>">
<input type="hidden" id="hdnTabName" value="<%=Request("pulsarplusDivId")%>" />
</FORM>
</BODY>
</HTML>
