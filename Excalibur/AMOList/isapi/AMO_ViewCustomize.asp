<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!------------------------------------------------------------------- 
'Description: AMO DATA
'----------------------------------------------------------------- //-->    
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataAVL.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
Call ValidateSession

	dim rsColumn, rsColumnSelected, sHideColumns, nParent
	
	
	nParent = Request.QueryString("nParent")
	
	set rsColumnSelected = Nothing
	
	set rsColumn = Server.CreateObject ("ADODB.Recordset")
	rsColumn.Fields.Append "Description", 129, 100
	rsColumn.Fields.Append "ID", 3, 4
	rsColumn.CursorLocation = 2	
    rsColumn.Open
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Select Availability (SA)"
    rsColumn.Fields("ID").Value = 1
	if IsODM = 0 then
		rsColumn.AddNew
		rsColumn.Fields("Description").Value = "AMO Price"
		rsColumn.Fields("ID").Value = 2
	end if
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "End of Sales (ES)"
    rsColumn.Fields("ID").Value = 3
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "PHweb (General) Availability (GA)"
    rsColumn.Fields("ID").Value = 4
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "End of Manufacturing (EM)"
    rsColumn.Fields("ID").Value = 5
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Global Series Config EOL"
    rsColumn.Fields("ID").Value = 6
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Hide from PRL"
    rsColumn.Fields("ID").Value = 7
	if IsODM = 0 then
		rsColumn.AddNew
		rsColumn.Fields("Description").Value = "AMO Cost"
		rsColumn.Fields("ID").Value = 8
		rsColumn.AddNew
		rsColumn.Fields("Description").Value = "Actual Cost"
		rsColumn.Fields("ID").Value = 9
	end if
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Alternative"
    rsColumn.Fields("ID").Value = 10
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Export Weight"
    rsColumn.Fields("ID").Value = 11
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Net Weight"
    rsColumn.Fields("ID").Value = 12
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Air Packed Weight"
    rsColumn.Fields("ID").Value = 13
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Air Packed Cubic"
    rsColumn.Fields("ID").Value = 14
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Export Cubic"
    rsColumn.Fields("ID").Value = 15
    'rsColumn.AddNew
    'rsColumn.Fields("Description").Value = "Owned By"
    'rsColumn.Fields("ID").Value = 16
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Warranty Code"
    rsColumn.Fields("ID").Value = 17
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Manufacture Country"
    rsColumn.Fields("ID").Value = 18
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Long Description"
    rsColumn.Fields("ID").Value = 19
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Rules Description"
    rsColumn.Fields("ID").Value = 20
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Replacements Description"
    rsColumn.Fields("ID").Value = 21
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Order Instruction"
    rsColumn.Fields("ID").Value = 22
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Product Line"
    rsColumn.Fields("ID").Value = 23
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "SKU Visibility"
    rsColumn.Fields("ID").Value = 24
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Replacement"
    rsColumn.Fields("ID").Value = 25
    rsColumn.AddNew
    rsColumn.Fields("Description").Value = "Hide from SCM"
    rsColumn.Fields("ID").Value = 26
    'rsColumn.AddNew
    'rsColumn.Fields("Description").Value = "Hide from SCL"
    'rsColumn.Fields("ID").Value = 27
    
    
   	rsColumn.Sort = "Description ASC"
   	
   	sHideColumns = GetDBCookie("AMO Hide Column")
   	if sHideColumns = "" then 
   		sHideColumns = 0
   	end if

%>

<HTML>
<head>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="AMO - View Customize" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title>AMO - View Customize</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/wizard%20style.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>


<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">

<title></title>
<script language="JavaScript" src="../library/scripts/formChek.js" type="text/javascript"></script>
<script language="JavaScript" src="../library/scripts/calendar.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript type="text/javascript">
<!--
//var oPopup = window.createPopup();


function btnSave_Click() {


	var sColumnIds, i;
	SelectAll(thisform.lbxSelectedColumn);
	sColumnIds = "";
	for (i = 0; i < thisform.lbxSelectedColumn.options.length; i++) {
	    sColumnIds = sColumnIds + "," + thisform.lbxSelectedColumn[i].value.trim();
	}
	if (sColumnIds != ""){
	    sColumnIds = sColumnIds.slice(1);
	}
	
	document.getElementsByName("nColumnIDs").item(0).value = sColumnIds
	thisform.action = "<%=Trim(nParent)%>.asp?nMode=10";	
	thisform.target = "_self"
	thisform.submit();	
								
}


function btnCancel_Click() {

	thisform.action = "<%=Trim(nParent)%>.asp?nMode=9";	
	thisform.target = "_self"
	thisform.submit();	

}

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="gray">
<!-- #include file="../library/includes/popup.inc" -->
<FORM name=thisform method=post>
	<table border=0 cellspacing=3 cellpadding=3>
		<TR>
			<td><font color='blue'><b>Customize which columns appear in the AMO list</b></font></td>
		</TR>
		<TR>
			<td colspan='2'><% 
				DualListboxRs_GetHTML6_Write rsColumn, "Description", "ID", rsColumnSelected, _
			"Description", "ID", true, true, sHideColumns, "Show Columns", "Hide Columns", _
			"Column", true, 300, 250, false, true, false, 300, 10
		
		%></td>
		</TR>
		
	</table>
	
	<table>
	<tr>
		<td>
		<INPUT id=btnSave name=btnSave type=button value="Save" LANGUAGE=javascript onclick="return btnSave_Click();">
		<INPUT id=btnCancel name=btnCancel type=button value="Cancel"  LANGUAGE=javascript onClick="btnCancel_Click();">
		</td>
	</tr>
	</table>
	<INPUT type="hidden" id=nColumnIDs name=nColumnIDs value="">
</FORM>
<%
set rsColumnSelected = Nothing
set rsColumn = Nothing

%>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->
