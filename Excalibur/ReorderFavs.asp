<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript" src="_ScriptLibrary/jsrsClient.js"></script>
<script src="includes/client/jquery.min.js" type="text/javascript"></script>
<script src="includes/client/jquery-ui.min.js" type="text/javascript"></script>
<script src="includes/client/jquery.blockUI.js" type="text/javascript"></script>

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
    
function lstAvailable_ondblclick() {
	if (lstAvailable.selectedIndex != "undefined")
		{
			if (lstAvailable.selectedIndex != -1 )
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[lstAvailable.selectedIndex].text,lstAvailable.options[lstAvailable.selectedIndex].value);
				lstAvailable.options[lstAvailable.options.selectedIndex]=null;
			}
		}

}

function lstNew_ondblclick() {
	if (lstNew.selectedIndex != "undefined")
		{
			if (lstNew.selectedIndex != -1 )
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[lstNew.selectedIndex].text,lstNew.options[lstNew.selectedIndex].value);
				lstNew.options[lstNew.options.selectedIndex]=null;
			}
		}
}

function cmdAdd_onclick() {
	var i;
	for (i=0;i<lstAvailable.length;i++)
		{
			if (lstAvailable.options[i].selected)
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[i].text,lstAvailable.options[i].value);
			}
		}

	for (i=lstAvailable.length-1;i>-1;i--)
		{
			if (lstAvailable.options[i].selected)
			{
				lstAvailable.options[i]=null;
			}
		}
		
		
}

function cmdRemove_onclick() {
	var i;
	for (i=0;i<lstNew.length;i++)
		{
			if (lstNew.options[i].selected)
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[i].text,lstNew.options[i].value);
			}
		}

	for (i=lstNew.length-1;i>-1;i--)
		{
			if (lstNew.options[i].selected)
			{
				lstNew.options[i]=null;
			}
		}
		


}


function myCallback(returnstring) {

    if (returnstring == 1)
        window.returnValue = "RefreshLeftTree";
    else
        window.returnValue = "";

	window.close();
} 


function cmdOK_onclick() {
	var i;
	var strFavs="";
	var FavCount=0;
	
	for(i=0;i<lstAvailable.length;i++)
		lstAvailable.options[i].selected=true;
	cmdAdd_onclick();
	
	for(i=0;i<lstNew.length;i++)
	{
		strFavs = strFavs + lstNew.options[i].value + ","
		FavCount = FavCount + 1;
	}

	if (strFavs != txtFavs.value) {
	    //jsrsExecute("FavoritesRSupdate.asp", myCallback, "UpdateFavs",Array(strFavs,String(FavCount),txtEmployeeID.value));
	    ajaxurl = "FavoritesRSupdate.asp?CurrentUserID=" + txtEmployeeID.value + "&FavCount=" + String(FavCount) + "&Favorites=" + strFavs;
	    $.ajax({
	        url: ajaxurl,
	        type: "POST",
	        success: function (data) {
	            if (data == 1)
	                window.returnValue = "RefreshLeftTree";
	            else
	                window.returnValue = "";

	            window.close();
	        },
	        error: function (xhr, status, error) {
	            alert(error);
	        }

	    });
	}
	else {
	    window.returnValue = 0;
	    window.close();
	}
}


function cmdCancel_onclick() {
	window.returnValue = 0;
	window.close();
}

function cmdImageButton_onmouseover() {
	window.event.srcElement.style.cursor = "default";
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
}

function cmdImageButton_onmouseout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
}

function cmdImageButton_onmousedown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";

}

function cmdImageButton_onmouseup() {
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	ImageButton_Pressed(window.event.srcElement.name);
}

function cmdImageButton_onkeydown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";
}

function cmdImageButton_onkeyup() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	if (window.event.keyCode !=9)
		ImageButton_Pressed(window.event.srcElement.name);
}

function ImageButton_Pressed(ID){
	if (ID=="cmdSel")
		cmdAdd_onclick();
	else if (ID=="cmdRemove")
		cmdRemove_onclick();
	else if (ID=="cmdSelProd")
		SelectProducts();
	else if (ID=="cmdRemoveProd")
		RemoveProducts();
	else if (ID=="cmdSelDel")
		SelectDel();
	else if (ID=="cmdRemoveDel")
		RemoveDel();
}

function cmdImageButton_onfocusout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
	window.event.srcElement.style.backgroundColor = "gainsboro";

}

function SelectProducts() {
	var i;
	for (i=0;i<lstAvailable.length;i++)
		{
			if (lstAvailable.options[i].value.substring(0,1)=="P")
			{
				lstSort.options[lstSort.length] = new Option(lstAvailable.options[i].text,lstAvailable.options[i].value);
			}
		}

	for (i=lstAvailable.length-1;i>-1;i--)
		{
			if (lstAvailable.options[i].value.substring(0,1)=="P")
			{
				lstAvailable.options[i]=null;
			}
		}

	sortListBox (lstSort,1);

	for (i=0;i<lstSort.length;i++)
		{
			if (lstSort.options[i].value.substring(0,1)=="P")
			{
				lstNew.options[lstNew.length] = new Option(lstSort.options[i].text,lstSort.options[i].value);
			}
		}

	for (i=lstSort.length-1;i>-1;i--)
		{
			if (lstSort.options[i].value.substring(0,1)=="P")
			{
				lstSort.options[i]=null;
			}
		}


}

function RemoveProducts() {
	var i;
	for (i=0;i<lstNew.length;i++)
		{
			if (lstNew.options[i].value.substring(0,1)=="P")
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[i].text,lstNew.options[i].value);
			}
		}

	for (i=lstNew.length-1;i>-1;i--)
		{
			if (lstNew.options[i].value.substring(0,1)=="P")
			{
				lstNew.options[i]=null;
			}
		}
}


function SelectDel() {
	var i;
	for (i=0;i<lstAvailable.length;i++)
		{
			if (lstAvailable.options[i].value.substring(0,1)!="P")
			{
				lstSort.options[lstSort.length] = new Option(lstAvailable.options[i].text,lstAvailable.options[i].value);
			}
		}

	for (i=lstAvailable.length-1;i>-1;i--)
		{
			if (lstAvailable.options[i].value.substring(0,1)!="P")
			{
				lstAvailable.options[i]=null;
			}
		}
	
	
	sortListBox (lstSort,1);
	
	for (i=0;i<lstSort.length;i++)
		{
			if (lstSort.options[i].value.substring(0,1)!="P")
			{
				lstNew.options[lstNew.length] = new Option(lstSort.options[i].text,lstSort.options[i].value);
			}
		}

	for (i=lstSort.length-1;i>-1;i--)
		{
			if (lstSort.options[i].value.substring(0,1)!="P")
			{
				lstSort.options[i]=null;
			}
		}

}

function RemoveDel() {
	var i;
	for (i=0;i<lstNew.length;i++)
		{
			if (lstNew.options[i].value.substring(0,1)!="P")
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[i].text,lstNew.options[i].value);
			}
		}

	for (i=lstNew.length-1;i>-1;i--)
		{
			if (lstNew.options[i].value.substring(0,1)!="P")
			{
				lstNew.options[i]=null;
			}
		}
}


function compareItemText(a,b) {
  return a.text!=b.text ? a.text<b.text ? -1 : 1 : 0;

}
function compareItemNumber(a,b) {
  return a.text - b.text;

}

function sortListBox(list,sortType) {
	var items = list.options.length;
  	var tmp = new Array(items);
  
	for ( i=0; i<items; i++ )
		tmp[i] = new Option(list.options[i].text,list.options[i].value);
  
  	if(sortType==1)
		tmp.sort(compareItemText);
	else
		tmp.sort(compareItemNumber);
  
	for ( i=0; i<items; i++ )
		list.options[i] = new Option(tmp[i].text,tmp[i].value);


}

//-->
</script>
</head>
<body bgcolor="ivory">
<br>
<table style="width:100%"><tr><td width="10">&nbsp;</td><td>
<%


dim cn
dim rs
dim p
dim cm

set cn = server.CreateObject("ADODB.Connection")
cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
cn.Open

set rs = server.CreateObject("ADODB.recordset")

dim strFavs
dim FavArray
dim i
dim strFav
dim CurrentUSer
dim CurrentUSerID

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

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		strFavs = rs("Favorites")
	end if
	rs.Close


'strFavs = trim(Request.Cookies("Favorites"))
if len(strFavs) <> "" then
	if right(strFavs,1) = "," then
		strFavs = left(strFavs,len(strFavs)-1)
	end if
end if

FavArray = split(strfavs,",")




%>


<font size="3" face="verdana"><b>Reorder Favorites</b></font><br><br>
<table border="0" width="100%">
<tr>
<td nowrap width="50%">
	<font face="verdana" size="2"><b>Current Order:</b></font><br>
	<select style="WIDTH:100%" size="15" id="lstAvailable" name="lstAvailable" multiple LANGUAGE="javascript" ondblclick="return lstAvailable_ondblclick()">
		<%
			for i = lbound(FavArray) to ubound(FavArray)
				strFav=""
				if left(trim(ucase(FavArray(i))),1) = "P" then
					rs.Open "spGetProductVersionName " & clng(mid(trim(ucase(FavArray(i))),2)),cn,adOpenForwardOnly
				else
					rs.Open "spGetDeliverableRootName " & clng(trim(ucase(FavArray(i))) ),cn,adOpenForwardOnly
				end if
				if not (rs.EOF and rs.BOF) then
					Response.Write "<OPTION value=""" & FavArray(i) & """>" & rs("Name") & "</OPTION>"
				end if
				rs.Close
			next 
		%>
	</select>
</td></tr>
<td nowrap width="50%">
	<table width="100%"><tr><td>
	<font face="verdana" size="2"><b>New Order:</b></font>
	</td>
	<td align="right">
<input type="image" src="images/downproduct.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdSelProd" name="cmdSelProd" title="Select Products (Sorted Alphabetically)" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20"><input type="image" src="images/upproduct.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdRemoveProd" name="cmdRemoveProd" title="Remove Products" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20">
&nbsp;
<input type="image" src="images/downdel.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdSelDel" name="cmdSelDel" title="Select Deliverables (Sorted Alphabetically)" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20"><input type="image" src="images/updel.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdRemoveDel" name="cmdRemoveDel" title="Remove Deliverables" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20">
&nbsp;
<input type="image" src="images/downarrow.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdSel" name="cmdSel" title="Select" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20"><input type="image" src="images/uparrow.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdRemove" name="cmdRemove" title="Remove" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="40" HEIGHT="20">&nbsp; 
	
		<input style="display:none;width:60" type="button" value="Add" id="cmdAdd" name="cmdAdd" LANGUAGE="javascript" onclick="return cmdAdd_onclick()">
		<input style="display:none;width:60" type="button" value="Remove" id="cmdRemove" name="cmdRemove" LANGUAGE="javascript" onclick="return cmdRemove_onclick()">
	</td></tr></table>
	<select style="WIDTH:100%" size="13" id="lstNew" name="lstNew" multiple LANGUAGE="javascript" ondblclick="return lstNew_ondblclick()">
	</select>
</td>
</tr>
</table>

<table width="100%">
<tr><td align="right"><input type="button" value="OK" id="cmdOK" name="cmdOK" LANGUAGE="javascript" onclick="return cmdOK_onclick()">&nbsp;<input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick()"></td></tr>
</table>

<%
set rs = nothing
cn.Close
set cn = nothing
%><input type="hidden" id="txtFavs" name="txtFavs" value="<%=strFavs & ","%>">
<input type="hidden" id="txtEmployeeID" name="txtEmployeeID" value="<%=CurrentUserID%>">
</td>
</tr>
</table><select style="display:none" size="2" id="lstSort" name="lstSort"></select>
</body>
</html>
