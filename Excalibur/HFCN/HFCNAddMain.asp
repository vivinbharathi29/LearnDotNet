<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<%if request("ID") = "" then%>
	<Title>Add Deliverable</title>
<%else%>
	<Title>Edit Deliverable</title>
<%end if%>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		if (event.srcElement.name == "cboDeliverable")
			cboDeliverable_onchange();
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}

function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }

function cmdOK_onclick() {
	var OutArray = new Array();
	var blnFailed;
	if (AddCat.txtName.value == "")
		{
			window.alert("Deliverable name is required.");
			AddCat.txtName.focus();
			blnFailed = true;
		}
	else if (AddCat.cboCategory.selectedIndex ==0)
		{
			window.alert("Category is required.");
			AddCat.cboCategory.focus();
			blnFailed = true;
		}

	if (!blnFailed)
		{
		if (AddCat.txtRootID.value == "")
			{
			OutArray[0]= AddCat.txtName.value;
			OutArray[1]= AddCat.cboCategory.value;
	
			window.returnValue = OutArray;
			window.close();
			}
		else
			{
				AddCat.submit();
			}
		}
	
}

function cmdCancel_onclick() {
	var OutArray = new Array();

	OutArray[0]= "";
	OutArray[1]= "";
	window.returnValue = OutArray;
	window.close();
}

function window_onload() {
	AddCat.txtName.focus();
}

function txtName_onkeypress() {
	//if (event.keyCode == 13)
	
	//	if (AddCat.cboCategory.selectedIndex < 1)
	//		AddCat.cboCategory.focus();
	//	else
	//		cmdOK_onclick();
}

function AddManager_onclick() {
	var strID = new Array();
	strID = window.showModalDialog("http://16.81.19.70/mobilese/today/employee.asp","","dialogWidth:460px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strID) != "undefined")
		{
			AddCat.cboPM.options[AddCat.cboPM.length] = new Option(strID[1],strID[0]);
			AddCat.cboPM.selectedIndex = AddCat.cboPM.length - 1
		}
}

//-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<%
	dim cn
	dim rs 
	dim cm
	dim p
	dim strName
	dim strCategoryID
	dim strDescription
	dim strDevManID
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")


	if request("ID") = "" then
		strName = ""
		strCategoryID = ""
		strDescription = ""
		strDevManID = ""
	else
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetRootProperties4Version"
		
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p
	
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'strSQL = "spGetRootProperties4Version " & request("ID")
		'rs.Open strSQL,cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strName = rs("Name") & ""
			strCategoryID = rs("Category") & ""
			strDescription = rs("Description") & ""
			strDevManID = rs("ManagerID") & ""
		end if
		rs.Close
	end if


	strCategories = ""
		
	strSQL = "spListHFCNCategories"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		if trim(Request("CatID")) = trim(rs("ID")) or trim(strCategoryID) = trim(rs("ID")) then
			strCategories = strCategories & "<Option selected Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		else
			strCategories = strCategories & "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close

%>	

<table border=0 width=100%>
<TD>
<%if request("ID") = "" then
	Response.Write "<h3>Add Deliverable</h3>"
else
	Response.Write "<h3>Edit Root Deliverable</h3>"
end if

%>
<font size=2 color=blue face=verdana>
Recommended guidelines for naming deliverables.<BR>
1. Do not include version information.<BR>
2. Do not use Product name unless necessary.<BR>
3. Do not use Vendor name unless necessary.<BR>
4. Describe common attributes accross all suppliers and versions.<BR>


</font>

<Form ID="AddCat" method="post" action="HFCNSaveRoot.asp">
<INPUT style="Display:none" type="text" id=text2 name=text2>
<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=Request("ID")%>">
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="150" nowrap><b>Deliverable&nbsp;Name:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT style="WIDTH:100%" type="text" id=txtName name=txtName value="<%=strName%>"  LANGUAGE=javascript onkeypress="return txtName_onkeypress()" maxlength="120"></td>
	</tr>
    <tr>
		<td nowrap><b>Category:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<select id="cboCategory" name="cboCategory" style="WIDTH: 100%" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">			
				<option selected></option>
				<%=strCategories%>
			</select>			
			</td>
    </tr>
    <%if request("ID") <> "" then%>
    
   <tr>
		<td nowrap><b>Dev.&nbsp;Manager:</b> <font color="#ff0000" size="1">*</font></td>
		<td><select id="cboPM" style="WIDTH: 335px" name="cboPM" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
      
		<%
		strSQL = "spGetEmployees"
		rs.Open strSQL,cn,adOpenForwardOnly
		do while not rs.EOF
			if strDevManID = rs("ID") & "" then
				Response.Write "<Option selected Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				Response.Write "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close
		%>
      
      
      </select>&nbsp;<input type="button" value="Add" id="cmdAddManager" name="cmdAddManager" LANGUAGE="javascript" onclick="return AddManager_onclick()" tabindex=-1></td>
    </tr>
	<%end if%>
<!--    
    <tr>
		<td nowrap><b>Vendor:</b> <font color="#ff0000" size="1">*</font> </td>
		<td colspan="10"><select id="cboVendor" name="cboVendor" style="WIDTH: 250px" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">			
		<option selected></option>
		<%
		strSQL = "spGetVendorList"
		rs.Open strSQL,cn,adOpenForwardOnly
		do while not rs.EOF
			if strVendorID = rs("ID") & "" then
				Response.Write "<Option selected Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else		
				Response.Write "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close
		%>

			</select>&nbsp;<input type="button" value="Add" id="cmdAddVendor" name="cmdAddVendor" LANGUAGE="javascript" onclick="return cmdAddVendor_onclick()">
		</td>
		
	</tr>
    -->
    <%if request("ID") <> "" then%>
    <tr>
		<td nowrap valign="top"><b>Description:</b></td>
		<td colspan="10">
		<TEXTAREA style="WIDTH: 100%; HEIGHT: 60px" rows=3 cols=51 id=txtDescription name=txtDescription><%=strDescription%></TEXTAREA>
		</td>
	</tr>
	<%end if%>
</table>
</form>
</TD></TR>
<TR><TD align=right>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK  LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
</TD></TR>
</TABLE>
<%
	set rs = nothing
	set cn = nothing
	
%>

</BODY>
</HTML>
