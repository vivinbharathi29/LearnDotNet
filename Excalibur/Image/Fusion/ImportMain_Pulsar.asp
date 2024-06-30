<%@ Language=VBScript %>
<!-- #include file="../../includes/no-cache.asp" -->
<%
    
Dim AppRoot
AppRoot = Session("ApplicationRoot")
 %>
<HTML>
<HEAD>
<link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />

</HEAD>
    <script language="javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
<script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>

<SCRIPT type="text/javascript" LANGUAGE=javascript>
<!--

function chkAll_onclick() {
	var i;
	if (typeof(frmImport.chkSelected)!="undefined")
	{
		if (typeof(frmImport.chkSelected.length)=="undefined")			
		{
		    if (frmImport.chkAll.checked)
			    frmImport.chkSelected.checked = true;
		    else
			    frmImport.chkSelected.checked = false;
		}		
		else
		{
		    for (i=0;i<frmImport.chkSelected.length;i++)
			{
		        if (frmImport.chkAll.checked)
					frmImport.chkSelected[i].checked = true;
				else
					frmImport.chkSelected[i].checked = false;
			}
		}
	}
}

function cboProduct_onchange() {
    var strID = frmImport.cboProduct.value;
    var targetID = frmImport.txtID.value;

    if (strID == "0" )
		{
        tblImages.innerHTML = "<Table ID=\"ImageTable\" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=\"width:16;height:16;\" type=\"checkbox\" id=chkALL name=chkAll LANGUAGE=javascript onclick=\"return chkAll_onclick()\"></TD><TD><font size=1 face=verdana><b>Product&nbsp;Drop</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Software&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Type&nbsp;&nbsp;</b></font></TD></tr><TR><TD colspan=4><font size=1 face=verdana>No Product Selected</font></TD></TR></table>";
		}
	else
		{
		
       var ajaxurl = "<%=AppRoot %>/Image/Fusion/ImportRSget_Pulsar.asp?SourceProductID=" + strID + "&TargetProductID=" + targetID + "&ProductReleaseID= <%=request("ProductReleaseID")%>";
        $.ajax({
	        url: ajaxurl,
	        type: "POST",
	        success: function (data) {
	            if (data != "") {
	                tblImages.innerHTML = data;
	            }
	        },
	        error: function (xhr, status, error) {
	            tblImages.innerHTML = error;
	        }

	    });
			
		
		}
}

//-->
</SCRIPT>

<BODY bgcolor=Ivory style="overflow:auto">


<form id=frmImport  method=post action="ImportSave_Pulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
<font color=black size=4 face=verdana><b>Select Images to Import</b></font><BR><BR>
<font size=1 face=verdana>Select a product to see all images on that product. Only Pulsar Products are candidates for import<BR><BR><b>Note: You may only import from one product at a time.</b><BR><BR></font>
<TABLE width=100% border=0><TR><TD align=right><font size=2 face=verdana><b>Product:&nbsp;</b></font>
<SELECT id=cboProduct name=cboProduct LANGUAGE=javascript onchange="return cboProduct_onchange()">
<OPTION selected value=0></OPTION>
<%
	dim cn
	dim cm
	dim p
	dim rs
	
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
      	Response.Redirect "../../NoAccess.asp?Level=1"
    else
            CurrentUserPartner = rs("PartnerID")
    end if 
    rs.Close

	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductPartner"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProductID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
			set rs = nothing
			set cn=nothing
				
			Response.Redirect "../../NoAccess.asp?Level=1"
		end if
		rs.close
	end if




	rs.Open "usp_Image_GetProductList 1",cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(rs("ID")) <> trim(request("ProductID")) then
			if CurrentUserPartner = "1" or (trim(rs("PartnerID")) = trim(CurrentUserPartner)) then
				Response.Write "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
			end if
		end if
		rs.MoveNext
	loop
	rs.Close

	set rs = nothing
	cn.Close
	set cn = nothing

%>
</SELECT>
</td></tr></table>
<DIV id=tblImages>
	<Table ID="ImageTable" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style="width:16;height:16;" type="checkbox" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>SKU&nbsp;Number</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Release&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Software&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Type&nbsp;&nbsp;</b></font></TD></tr>
	<TR><TD colspan=4><font size=1 face=verdana>No Product Selected.</font></TD></TR></table>
</DIV>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ProductID")%>"> 
<INPUT type="hidden" id=txtImportLocalizations name=txtImportLocalizations value="0">
<INPUT type="hidden" id=productReleaseID name=productReleaseID value="<%=request("ProductReleaseID")%>">
</form>

</BODY>
</HTML>
