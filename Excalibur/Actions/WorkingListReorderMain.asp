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
<script language="JavaScript" src="../_ScriptLibrary/jsrsClient.js"></script>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
		if (String(KeyString).length >0)
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


function sortList(strName) {
	var lb = document.getElementById(strName);
	arrTexts = new Array();

	for(i=0; i<lb.length; i++)  
		{
		arrTexts[i] = lb.options[i].text;
		}

	arrTexts.sort(sortNumber);

	for(i=0; i<lb.length; i++)  
	{
		lb.options[i].text = arrTexts[i];
		lb.options[i].value = arrTexts[i];
	}

}
function sortNumber(a,b)
	{
	return a - b
	}


function SetOrder(ID){
	var SortRequired=false;
	var PreviousValue;
	if (frmMain.cboNew.selectedIndex >= frmMain.cboNew.length-1)
		PreviousValue = "1" ;//frmMain.cboNew.options[0].value;
	else
		PreviousValue = frmMain.cboNew.options[frmMain.cboNew.selectedIndex+1].value;
	
	if (document.all("NewPositionLink" + ID).innerText != "Set")
		{
		frmMain.cboNew.options[frmMain.cboNew.options.length] = new Option(document.all("NewPositionLink" + ID).innerText,document.all("NewPositionLink" + ID).innerText);
		SortRequired = true
		}
	document.all("NewPositionLink" + ID).innerText = frmMain.cboNew.options[frmMain.cboNew.selectedIndex].value;
	document.all("txtValueList" + ID).innerText = frmMain.cboNew.options[frmMain.cboNew.selectedIndex].value;
	if (frmMain.cboNew.selectedIndex < frmMain.cboNew.length-1)
		{
		frmMain.cboNew.selectedIndex = frmMain.cboNew.selectedIndex + 1
		if (frmMain.txtReportOption.value!="2")
		    frmMain.cboNew.options[frmMain.cboNew.selectedIndex-1] = null;
		}
	else if (frmMain.txtReportOption.value!="2")
        {
		frmMain.cboNew.options[frmMain.cboNew.selectedIndex] = null;
	    }
	    
	if (SortRequired)
		{
		sortList ("cboNew");
		
		for (i=0;i<frmMain.cboNew.options.length;i++)
			
			if(frmMain.cboNew.options(i).value==PreviousValue)
				{
				frmMain.cboNew.selectedIndex=i;
				return;
				}
		}
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	Font-Family: Verdana;
	Font-Size: xx-small;
}
</STYLE>
<BODY bgcolor="ivory">

<%

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function

	dim cn
	dim rs
	dim i
	dim cm
	dim p
	dim CurrentUser
	dim CurrentUserID
	dim strID
	dim strEmployee
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
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
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	strID = trim(Request("ID"))
	if strID = "" then
		strID = 0
	end if


strEmployeeName = ""
if strID <> "" and isnumeric(strID) then
	rs.Open "spGetEmployeeByID " & clng(request("ID")),cn,adOpenStatic
	if not(rs.EOF and rs.BOF) then
		strEmployeeName = rs("Name") & ""
	end if
	rs.Close
end if

strProjectName = ""
if request("ProjectID") <> "" and isnumeric(request("ProjectID")) then
	rs.Open "spGetProductVersion " & clng(request("ProjectID")),cn,adOpenStatic
	if not(rs.EOF and rs.BOF) then
		strProjectName = rs("DOTSName") & ""
	end if
	rs.Close
end if
if strID = "" and strEmployeeName = "" then
	Response.Write "Not enough information supplied to display this page."
else
    if trim(request("ReportOption")) = "2" then
	    Response.Write "<font size=3 face=verdana><b>Reorder " & strProjectName & " Task List</b></font>"
        rs.Open "spListActionItemsForTools " & clng(request("ProjectID")) & ",2,0",cn,adOpenForwardOnly
    elseif trim(request("ReportOption")) = "3" then
	    Response.Write "<font size=3 face=verdana><b>Reorder " & strProjectName & " Task List for " & shortname(strEmployeeName) & "</b></font>"
        rs.Open "spListActionItemsForTools " & clng(request("ProjectID")) & ",2," & clng(strID),cn,adOpenForwardOnly
	else
	    Response.Write "<font size=3 face=verdana><b>Reorder " & strProjectName & " Working List for " & shortname(strEmployeeName) & "</b></font>"
    	rs.Open "spListActionItemWorkingList " & clng(strID),cn,adOpenForwardOnly
	end if
	
	if rs.EOF and rs.BOF then
		Response.Write "No Working List action items found for this product."
	else
		Response.Write "<form id=frmMain method=post action=""WorkingListReorderSave.asp"">"
		Response.Write "<font face=verdana size=2><b>Next Position: </b></font>"		Response.write "<SELECT sorted id=cboNew name=cboNew>"
		for i = 1 to 100
			Response.Write "<OPTION value=""" & i & """>" & i & "</OPTION>"
		next		Response.Write "</SELECT>"
		Response.Write "<Table cellscaping=0 cellpadding=0 border=0><TR bgcolor=beige>"
		Response.Write "<TD><b>Old</b></TD>"
		Response.Write "<TD><b>New</b></TD>"
        if trim(request("ReportOption")) = "2" then
		    Response.Write "<TD><b>Owner</b></TD>"
		    Response.Write "<TD><b>ID</b></TD>"
		    Response.Write "<TD><b>Priority</b></TD>"
		    Response.Write "<TD><b>Summary</b></TD>"
        elseif trim(request("ReportOption")) = "3" then
		    Response.Write "<TD><b>ID</b></TD>"
		    Response.Write "<TD><b>Priority</b></TD>"
		    Response.Write "<TD><b>Summary</b></TD>"
		else
		    Response.Write "<TD><b>ID</b></TD>"
		    Response.Write "<TD><b>Priority</b></TD>"
		    Response.Write "<TD><b>Product</b></TD>"
		    Response.Write "<TD><b>Summary</b></TD>"
		end if
		Response.Write "</TR>"
		i=0
		do while not rs.EOF
			Response.Write "<TR>"
			Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("Displayorder") & "</TD>"
			Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid""><a id=NewPositionLink" & trim(i) & " href=""javascript:SetOrder(" & i & ");"">Set</a><INPUT type=""hidden"" id=txtValueList" & trim(i) & " name=txtValueList value=""0""></TD>"
            if trim(request("ReportOption")) = "2" then
		    	Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"" nowrap>" & rs("Owner") & "&nbsp;&nbsp;&nbsp;</TD>"
    			Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("ID") & "<INPUT type=""hidden"" id=txtIDList name=txtIDList value=""" & trim(rs("ID")) & """></TD>"
	    		Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"" align=middle>" & rs("Priority") & "</TD>"
			    Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("Summary") & "</TD>"
	        elseif trim(request("ReportOption")) = "3" then
    			Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("ID") & "<INPUT type=""hidden"" id=txtIDList name=txtIDList value=""" & trim(rs("ID")) & """></TD>"
	    		Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"" align=middle>" & rs("Priority") & "</TD>"
			    Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("Summary") & "</TD>"
		    else
    			Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("ID") & "<INPUT type=""hidden"" id=txtIDList name=txtIDList value=""" & trim(rs("ID")) & """></TD>"
	    		Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"" align=middle>" & rs("Priority") & "</TD>"
		    	Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"" nowrap>" & shortname(rs("Product") & "") & "&nbsp;&nbsp;&nbsp;</TD>"
			    Response.Write "<TD style=""BORDER-TOP: gainsboro thin solid"">" & rs("Summary") & "</TD>"
		    end if
			Response.Write "</TR>"
			i=i+1
			rs.MoveNext
		loop
		


					
	end if
	rs.Close

%>

	<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
	<INPUT type="hidden" id=txtProjectID name=txtProjectID value="<%=request("ProjectID")%>">
	<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
	<INPUT type="hidden" id=txtReportOption name=txtReportOption value="<%=trim(request("ReportOption"))%>">
	
	</form>
<%

end if

set rs = nothing
cn.Close
set cn=nothing

%>


</BODY>
</HTML>


