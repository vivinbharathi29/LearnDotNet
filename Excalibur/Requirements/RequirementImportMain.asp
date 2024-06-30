<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function chkAll_onclick() {
	var i;
	
	if (typeof(frmRequirement.chkSelected)!="undefined")
		{
			if (typeof(frmRequirement.chkSelected.length)=="undefined")			
				{
				if (frmRequirement.chkAll.checked)
					frmRequirement.chkSelected.checked = true;
				else
					frmRequirement.chkSelected.checked = false;
				}
		
			else
				{
				for (i=0;i<frmRequirement.chkSelected.length;i++)
					{
						if (frmRequirement.chkAll.checked)
							frmRequirement.chkSelected(i).checked = true;
						else
							frmRequirement.chkSelected(i).checked = false;
					}
				}
		}
}



function cboProduct_onchange() {
	var strID = frmRequirement.cboProduct.value;
	var oldLocation;
	
	if (strID == "" )
		{
		tblRequirements.innerHTML = "<Table ID=\"ReqTable\" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=\"width:16;height:16;\" type=\"checkbox\" id=chkALL name=chkAll LANGUAGE=javascript onclick=\"return chkAll_onclick()\"></TD><TD><font size=1 face=verdana><b>Requirement</b></font></TD><TD><font size=1 face=verdana><b>Specification&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Deliverables&nbsp;&nbsp;</b></font></TD></tr><TR><TD colspan=4><font size=1 face=verdana>No Product Selected</font></TD></TR></table>";
		}
		
	oldLocation = window.frames["RemoteFrame"].location;
	document.all.RemoteFrame.src="RequirementImportRFget.asp?ImportID=" + strID + "&ID=" + frmRequirement.txtID.value;
}



function RemoteFrame_onload() {
	tblRequirements.innerHTML = window.frames["RemoteFrame"].document.body.innerHTML;
}

//-->
</SCRIPT>
</HEAD>


<BODY bgcolor=Ivory>


<form ID=frmRequirement action=RequirementImportSave.asp method=post>

<%	

	dim cn 
	dim rs
	dim strDeliverables


if request("ProductID") = "" then
	Response.Write "<BR><font size=2 face=verdana>Not enough information to display this page</font>"
else
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")
	rs.open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product</font>"
		rs.Close
	else
		Response.Write "<font color=black size=4 face=verdana><b>Select Requirements to Import to " & rs("Name") & "</b></font><BR><BR>"
		Response.Write "<font size=1 face=verdana>Select a product to see all requirements on that product which do not exist for " & rs("Name") & ".<BR><BR><b>Note: You may only import from one product at a time.</b><BR><BR></font>"
		Response.Write "<TABLE width=100% border=0><TR><TD align=right><font size=2 face=verdana><b>Product:&nbsp;</b></font>"
		rs.Close
		Response.Write "<SELECT id=cboProduct name=cboProduct LANGUAGE=javascript onchange=""return cboProduct_onchange()"" ><option value="""" selected></option>"
		rs.Open "spGetProducts",cn,adOpenForwardOnly
		do while not rs.EOF
			if rs("ID") <> request("ProductID") then
				Response.Write "<option value=""" & rs("ID") & """>" & rs("Name") & " " & rs("version") & "</option>"
			end if
			rs.MoveNext
		loop
		rs.Close
		Response.Write "</SELECT></td></TR></table>"
		Response.Write "<SPAN id=""tblRequirements"">"
		Response.Write "<Table ID=""ReqTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Requirement</b></font></TD><TD><font size=1 face=verdana><b>Specification&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Deliverables&nbsp;&nbsp;</b></font></TD></tr>"
		Response.Write "<TR><TD colspan=4><font size=1 face=verdana>No Product Selected</font></TD></TR>"
		Response.Write "</table>"
		Response.Write "</SPAN>"
	  
	end if
	set rs = nothing
	set rs2 = nothing
	cn.Close
	set cn = nothing
end if


  %>
    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
  <INPUT type="hidden" id=txtID name=txtID value="<%= request("ProductID")%>">
</form>
<IFRAME style="display:none" ID=RemoteFrame LANGUAGE=javascript onload="return RemoteFrame_onload()">

</IFRAME>
</BODY>
</HTML>
