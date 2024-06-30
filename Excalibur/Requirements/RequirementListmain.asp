<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function chkAll_onclick() {
	var i;
	
	for (i=0;i<frmRequirement.chkSelected.length;i++)
		{
			if (frmRequirement.chkAll.checked)
				frmRequirement.chkSelected(i).checked = true;
			else
				frmRequirement.chkSelected(i).checked = false;
		}
}


function AddRequirement(){
	var strID = new Array;
	
	strID = window.showModalDialog("UpdateMasterList.asp","","dialogWidth:500px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
	if (typeof(strID) != "undefined")
		{
		var NewRow;
		var NewCell;
	
		NewRow = ReqTable.insertRow(1);
		NewRow.bgColor = "Ivory";
		NewRow.vAlign = "top";
		NewRow.name = "Row" + strID[0];
		NewRow.id = "Row" + strID[0];
	
		NewCell = NewRow.insertCell();
		NewCell.innerHTML = "<INPUT value=\"" + strID[0] + "\" style=\"width:16;height:16;\" type=\"checkbox\" id=chkSelected checked name=chkSelected><INPUT value=\"" + strID[0] + "\" style=\"width:16;height:16;display:none\" type=\"checkbox\" id=chkTag name=chkTag>";
		NewCell = NewRow.insertCell();
		NewCell.innerHTML = "<font size=1 face=verdana><b>" + strID[1] + "</b></font>";
		NewCell = NewRow.insertCell();
		NewCell.innerHTML = "&nbsp;";
		NewCell = NewRow.insertCell();
		NewCell.innerHTML = "&nbsp;";

		}
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=White>

<form ID=frmRequirement action=RequirementListSave.asp method=post>

<%	

	dim cn 
	dim rs
	dim strDeliverables


if request("ID") = "" then
	Response.Write "<BR><font size=2 face=verdana>Not enough information to display this page</font>"
else
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")
	rs.open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product</font>"
		rs.Close
	else
		Response.Write "<font color=black size=4 face=verdana><b>Select " & rs("name") & " Requirements</b></font>"
		Response.Write "<TABLE style=""display:none"" width=100% border=0><TR><TD align=right><font size=2 face=verdana><a href=""javascript: AddRequirement();"">Add Unlisted Requirement</a></font></td></TR></table>"
		rs.Close
		rs.Open "spListRequirementsByProductWeb " & clng(request("ID")),cn,adOpenForwardOnly
		Response.Write "<Table ID=""ReqTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Requirement</b></font></TD><TD><font size=1 face=verdana><b>Specification&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Deliverables&nbsp;&nbsp;</b></font></TD></tr>"
		do while not rs.EOF 
			if rs("ID") <> 1960 then
				rs2.Open "splistdeliverablesbyrequirement " & clng(rs("ID")) & "," & clng(request("ID")), cn, adOpenForwardOnly
				strDeliverables = ""
				do While Not rs2.EOF
					strDeliverables = strDeliverables & "-" & rs2("Name") &  "<BR>"
					rs2.MoveNext
				Loop
				rs2.Close
		
				if strDeliverables = "" then
					strDeliverables = "&nbsp;"
				end if		
				
				if not isnull(rs("ProductID")) then
					Response.Write "<TR ID=""Row" & rs("ID") & """ valign=top bgcolor=lightsteelblue><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("ID") & """ checked style=""width:16;height:16;"" type=""checkbox"" id=chkSelected name=chkSelected><INPUT value=""" & rs("ID") & """ checked style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
				elseif rs("active") then
					Response.Write "<TR ID=""Row" & rs("ID") & """ valign=top bgcolor=Ivory><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("ID") & """ style=""width:16;height:16;"" type=""checkbox"" id=chkSelected name=chkSelected><INPUT value=""" & rs("ID") & """ style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
				end if
				response.write "<TD nowrap style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana><b>" & rs("Requirement") & "&nbsp;</b></font></TD>"
				response.write "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Spec") & "&nbsp;</font></TD>"
				response.write "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & strDeliverables  & "</font></TD></tr>"
			end if
			rs.MoveNext
		loop
		rs.Close
		Response.Write "</table>"
	  
	end if
	set rs = nothing
	set rs2 = nothing
	cn.Close
	set cn = nothing
end if


  %>
  <INPUT type="hidden" id=txtID name=txtID value="<%= request("ID")%>">
    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</form>
</BODY>
</HTML>
