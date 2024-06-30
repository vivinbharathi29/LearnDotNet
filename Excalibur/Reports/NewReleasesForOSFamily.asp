<%@ Language=VBScript %>
<%
	if request("Type") = "Excel" then
		Response.ContentType = "application/vnd.ms-excel"
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript><!--function PageLoad(){    ProductSelectCell.innerHTML = txtProductSelection.value;}function AdvancedTarget (ProdID,VerID,RootID){
	var strResult;
	strResult = window.showModalDialog("../Target/TargetAdvanced.asp?ProductID=" + ProdID + "&VersionID=" + VerID + "&RootID=" + RootID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	//if (typeof(SelectedRow) != "undefined")
	//	{
	//	SelectedRow.style.color="black";
	//	SelectedRow = null;
	//	}
	if (typeof(strResult) != "undefined")
		{
//            if (typeof(window.parent.frames["TitleWindow"])!="undefined")
//                if (typeof(window.parent.frames["TitleWindow"].SavedValue)!="undefined")
  //      		    window.parent.frames["TitleWindow"].SavedValue.value = window.document.body.scrollTop;

			window.location.reload(true);
		}
	
}

function Releases_onmouseout() {
//	if(typeof(oPopup) == "undefined") 
//		return;

//	if (! oPopup.isOpen)
//		{
		if (window.event.srcElement.className == "text")
	    	window.event.srcElement.parentElement.parentElement.style.color = "black";
		else if (window.event.srcElement.className == "cell")
	    	window.event.srcElement.parentElement.style.color = "black";
//		}
}

function Releases_onmouseover() {
//	if(typeof(oPopup) == "undefined") 
//		return;
//	if (! oPopup.isOpen)
//		{
		if (window.event.srcElement.className == "cell")
			{
    		window.event.srcElement.parentElement.style.color = "red";
			window.event.srcElement.parentElement.style.cursor = "hand";
			}
		else if (window.event.srcElement.className == "text")
			{
    		window.event.srcElement.parentElement.parentElement.style.color = "red";
			window.event.srcElement.parentElement.parentElement.style.cursor = "hand";		
			}

//		if (typeof(SelectedRow) != "undefined")
//			if (SelectedRow != null)
//				SelectedRow.style.color="black";

//		}
}

function SelectProducts(){    SelectLink.style.display = "none";    ProductDisplay.style.display = "";}function CancelFilter(){    SelectLink.style.display = "";    ProductDisplay.style.display = "none";}function FilterReport (){    frmMain.submit();}function CheckAll(){    var i;    for (i=0;i<frmMain.lstProducts.length;i++)        frmMain.lstProducts[i].checked = true;}function UncheckAll(){    var i;    for (i=0;i<frmMain.lstProducts.length;i++)        frmMain.lstProducts[i].checked = false;}//--></SCRIPT>
</HEAD>
<STYLE>
    Body
    {
        FONT-Family: verdana;
        FONT-Size: x-small;	
    }
    Table
    {
        FONT-Family: verdana;
        FONT-Size: xx-small;	
    }
    A:link
    {
        COLOR: blue
    }
    A:visited
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }
</STYLE>
<BODY LANGUAGE=javascript onload="return PageLoad();"> 
<font size=2 face=verdana><b>New Releases for Win7</b></font>
<form id=frmMain action="NewReleasesForOSFamily.asp" method=get>
<table id=SelectLink width="100%"><tr><td align=right><a href="javascript:SelectProducts();">Select Products</a></td></tr></table>
<table style="display:none" id=ProductDisplay width="100%">
<tr><td bgcolor=gainsboro colspan=2><b>Choose Products to Display</b></td></tr>
<tr><td id=ProductSelectCell width="100%"></td>
    <td valign=top nowrap align=right>
        <a href="javascript:FilterReport();">Filter&nbsp;Report</a>&nbsp;|&nbsp;<a href="javascript:CancelFilter();">Cancel</a>
    </td></tr></table>
<%
    dim cn, rs
    dim cnstring
    dim blnCanEdit
    dim strProductIDs
    dim strProducts
    dim ProductCount
    blnCanEdit = false
    
  	set cn = server.CreateObject("ADODB.Connection")
  	set rs = server.CreateObject("ADODB.recordset")
	cnString =Session("PDPIMS_ConnectionString")
   	cn.ConnectionString = cnString
   	cn.Open
    
    
    
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		if rs("SystemAdmin") then
		    blnCanEdit = true
		end if
    end if
	rs.Close

    if not blnCanEdit then
        rs.Open "spListPMsActive 0",cn,adOpenForwardOnly
		do while not rs.EOF
    	    if trim(rs("ID")) = trim(currentuserid) then
    	        blnCanEdit = true
    	        exit do
    	    end if
    		rs.MoveNext				
	    loop
		rs.Close	
    end if
    
    rs.open "spListNewReleases4OSFamily 10",cn
    if rs.eof and rs.bof then
        response.Write "No New Releases Found."
    else
        response.Write "<table bgcolor=ivory cellpadding=2 cellspacing=0 border=1 width=""100%"">"
        response.Write "<tr bgcolor=beige><td><b>Product</b></td><td><b>Deliverable</b></td><td><b>Version</b></td></tr>"
        strProductIDs = ""
        strProducts = "<Table><TR><td colspan=8><a href=""javascript:CheckAll();"">Check&nbsp;All</a>&nbsp;|&nbsp;<a href=""javascript:UncheckAll();"">Uncheck&nbsp;All</a></td></tr><tr>"
        ProductCount = 0
        do while not rs.eof
            if instr("," & strProductIDs & ",","," & trim(rs("ProductID")) & "," ) = 0 then
                if ProductCount mod 8 = 0 then
                    strProducts = strProducts &  "</tr><tr>"
                end if
                strProductIDs = strProductIDs & "," & rs("ProductID")
                if request("lstProducts") = "" or  instr("," & trim(replace(request("lstProducts")," ","")) & ",","," & trim(rs("ProductID") & "") & ",") <> 0 then
                    strProducts = strProducts & "<TD nowrap><input checked id=""lstProducts"" name=""lstProducts"" type=""checkbox"" value=""" & rs("ProductID") & """>" & rs("Product") & "</td>"
                else
                    strProducts = strProducts & "<TD nowrap><input id=""lstProducts"" name=""lstProducts"" type=""checkbox"" value=""" & rs("ProductID") & """>" & rs("Product") & "</td>"
                end if
                ProductCount = ProductCount + 1
            end if
 
            if request("lstProducts") = "" or  instr("," & trim(replace(request("lstProducts")," ","")) & ",","," & trim(rs("ProductID") & "") & ",") <> 0 then
 
                if blnCanEdit then
                    response.Write "<tr LANGUAGE=javascript onmouseover=""return Releases_onmouseover()"" onmouseout=""return Releases_onmouseout()"" onclick=""return AdvancedTarget(" & clng(rs("productID")) & "," & clng(rs("versionID")) & "," & clng(rs("rootID")) & ");"">"
                else
                    response.Write "<tr>"
                end if
                response.Write "<td class=""cell"">" & rs("Product") & "</td>"
                response.Write "<td class=""cell"">" & rs("Deliverablename") & "</td>"
                response.Write "<td class=""cell"">" & rs("Version") & "," & rs("Revision") & "," & rs("Pass") & "</td>"
                response.Write "</tr>"
                end if
            rs.movenext
        loop
        response.Write "</tr></table>"
        strProducts = strProducts & "</Table>"
    end if
    rs.close  
    
    
    cn.close
    set cn = nothing    
%>

</form>
<textarea style="display:none" id="txtProductSelection" rows="2" cols="20"><%=server.htmlencode(strProducts)%></textarea>
</BODY>
</HTML>
