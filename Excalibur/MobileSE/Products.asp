<%@ Language=VBScript %>
<html>
<head>
<title>Product Information - HP Restricted</title>
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function lstCompare_onclick() {
	var SelectedCount = 0;
	var SelectedIDs = "";
	
	for (i=0;i<lstCompare.length;i++)
		if (lstCompare.options[i].selected)
			{
			SelectedIDs = SelectedIDs + "_" + lstCompare.options[i].value;
			SelectedCount = SelectedCount + 1;
			}
	
	if (SelectedCount < 2)
		alert("You must select at least 2 products to compare.");
	else
		{
		var IncludeImages;
		if (chkIncludeImageSummary.checked)
			IncludeImages = "&Images=1";
		else
			IncludeImages = "&Images=0";		
			
		SelectedIDs = SelectedIDs.substring(1,SelectedIDs.length);
		if (chkDelta.checked)
			window.location.href = "Deliverables.asp?ID=" + SelectedIDs  + "&Report=2" + IncludeImages
		else
			window.location.href = "Deliverables.asp?ID=" + SelectedIDs + "&Report=1" + IncludeImages

		}	
	
}

function cmdExcel_onclick() {
	if (cboSingleProduct.selectedIndex >0)
		{
		strID = cboSingleProduct.options[cboSingleProduct.selectedIndex].value;
		window.open ("../Reports/DelSummary.asp?ID=" + strID + "&Type=Excel");
		}
	else
		alert("Select a product first.");
}

function cmdHTML_onclick() {
	if (cboSingleProduct.selectedIndex >0)
		{
		strID = cboSingleProduct.options[cboSingleProduct.selectedIndex].value;
		window.open ("../Reports/DelSummary.asp?ID=" + strID);
		}
	else
		alert("Select a product first.");
}

//-->
</SCRIPT>
</head>




<body background="images/shadow.gif">

<!-- This is the Headder Table -->
<table border="0" cellPadding="1" cellSpacing="1" width="100%" border="0">
  <tr>
    <td nowrap width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
	<td><font size="5" face="Tahoma" color="#006697"><b>Excalibur Reports</b></font></td>
<!--    <td><IMG height=50 src="images/information.gif" width=137></td>    <td align="right"><IMG height=50 src="images/programoffice.gif" width=283></td> --> </tr>
</table>
   
<!-- #include file = "menubar.asp" -->
    
<!-- Web Page Stuff Starts Here -->
<%
				dim rs	
				dim cn

				set cn = server.CreateObject("ADODB.Connection")
				cn.ConnectionString = Session("PDPIMS_ConnectionString")
				cn.Open
			
				set rs = server.CreateObject("ADODB.recordset")
				rs.ActiveConnection = cn

%>

	<h2>Product Information<hr></h2>

       <ul>
        <li><strong>Extract Software Deliverable List:</strong><br>
		<SELECT id=cboSingleProduct name=cboSingleProduct style="WIDTH: 217px">
			<OPTION></OPTION>			<%				rs.Open "SELECT v.ID, f.name + ' ' + isnull(v.version,'') as name FROM productfamily f with (NOLOCK), productversion v with (NOLOCK) where (v.productstatusid<5) and v.productfamilyid = f.id order by f.name,v.version;"
				rs.MoveFirst
				do while not rs.EOF
					Response.Write "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" & vbcrlf
					rs.MoveNext
				loop
				rs.Close
						%>
		</SELECT>
		<INPUT type="button" value="Excel" id=cmdExcel name=cmdExcel LANGUAGE=javascript onclick="return cmdExcel_onclick()">&nbsp;
		<INPUT type="button" value="HTML" id=cmdHTML name=cmdHTML LANGUAGE=javascript onclick="return cmdHTML_onclick()">
		</Ul>
       <ul>
        <li><strong>Compare Deliverable Lists:</strong><br>
        <font face=verdana size=1 color=green>Select all products to Compare.  Use &LT;CTRL&GT; or &LT;SHIFT&GT; to multi-select.</font>
        <table border="0" cellPadding="1" cellSpacing="1" width="75%">
          
          <tr>
            <td Width="220"><select id="lstCompare" name="lstCompare" size="2" style="HEIGHT: 350px; WIDTH: 217px" multiple>
<%
	
				rs.Open "SELECT v.ID, f.name + ' ' + v.version as name FROM productfamily f with (NOLOCK), productversion v with (NOLOCK) where  v.productfamilyid = f.id order by f.name,v.version;"
				rs.MoveFirst
				do while not rs.EOF
					Response.Write "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" & vbcrlf
					rs.MoveNext
				loop
				rs.Close
				set rs=nothing
				set cn=nothing
%>               
              </select><br>
<!--              <input type="radio" id="optCompReq" name="optCompare" checked="true">Compare Requirements<br>              <input type="radio" id="optCompDel" name="optCompare">Compare Deliverables<br>              <input type="radio" id="optCompBoth" name="optCompare">Compare Product Definitions<br>              <input type="radio" id="optCompSchedule" name="optCompare">Compare Schedules (Text)<br>              <input type="radio" id="optCompScheduleGraph" name="optCompare">Compare Schedules (Graph)<br>-->              
              
              </td>
            <td vAlign="top">
              <input align="top" id="cmdCompare" name="cmdCompare" type="button" value="Compare" LANGUAGE=javascript onclick="return lstCompare_onclick()"><BR>
              <font size=2 face=verdana><BR><b>Show:</b><BR></font>
              <INPUT type="radio" id=chkAll name=chkShow checked> All<BR>
              <INPUT type="radio" id=chkDelta name=chkShow> Delta<BR>
              <font size=2 face=verdana><BR><b>Include:</b><BR></font>
              <INPUT type="checkbox" id=chkIncludeImageSummary name=chkIncludeImageSummary> Image Summary
            </td></tr></table>
            
        </td>
        </tr>       
        </table>

<div id="cc"><h1>HP Restricted</h1></div>


</body>
</html>
