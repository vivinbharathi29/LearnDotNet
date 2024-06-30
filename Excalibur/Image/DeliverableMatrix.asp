<%@ Language=VBScript %>
	
<%
	
Select Case request("Type")
	Case 1
		Response.ContentType = "application/vnd.ms-excel"
	Case 2
		Response.ContentType = "application/msword"
	Case 3
		Response.ContentType = "application/vnd.ms-excel"
	Case Else
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
End Select

%>
	
<HTML>
<HEAD>
<Title>Deliverable in Image</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


var oPopup = window.createPopup();

function Delrows_onmouseover(strID) {
	
	if (window.event.srcElement.className != "link")
		{
		//document.all("Row" + strID).style.cursor = "hand";
		//document.all("Row" + strID).style.color = "red";
		window.event.srcElement.parentElement.style.cursor = "hand";
		window.event.srcElement.parentElement.style.color = "red";
		}
}

function Delrows_onmouseout(strID) {
//	document.all("Row" + strID).style.color = "black";
	window.event.srcElement.parentElement.style.color = "black";
}

function DisplayVersion (VersionID,RootID){
	var strID;
	
	//strID = window.showModalDialog("../WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
	strID = window.open("../Query/DeliverableVersionDetails.asp?Type=1&RootID=" + RootID + "&ID=" + VersionID)
		
	//if (typeof(strID) != "undefined")
	//	{
	//		window.location.reload(true);
	//	}
}


function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function DelMenu (strID,RootID,blnHasPath){

    // The variables "lefter" and "topper" store the X and Y coordinates
    // to use as parameter values for the following show method. In this
    // way, the popup displays near the location the user clicks. 
    
    
	if (window.event.srcElement.className == "link")
		return;
    
    
    var lefter = event.clientX;
    var topper = event.clientY;
    var popupBody;
	//var strPath = trim(document.all("DelPath" + strID).innerText);
	//var strPath = "\\16.81.19.70";
	if(typeof(oPopup) == "undefined") 
		{
		DisplayVersion(strID);
		
		return;
		}

  
  
  //if (strPath != "" && strPath != ".")
	if (blnHasPath)
	{
    popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:GetVersion(" + strID + ")'\" >&nbsp;&nbsp;Download</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayChanges(" + strID + ")'\" >&nbsp;&nbsp;Display Changes</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayPartNumbers(" + strID + "," + RootID + ")'\" >&nbsp;&nbsp;Show&nbsp;CD&nbsp;Part&nbsp;Number&nbsp;&nbsp;&nbsp;&nbsp;</SPAN></FONT></DIV>";

	popupBody = popupBody + "<DIV>";
	popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

	popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:DisplayVersion(" + strID + "," + RootID + ")'\" ><FONT face=Arial size=2>&nbsp;&nbsp;Properties</FONT></SPAN></DIV>";

	popupBody = popupBody + "</DIV>";

    oPopup.document.body.innerHTML = popupBody; 
	
	oPopup.show(lefter, topper, 140, 92, document.body);
	
	
	//Adjust window size
	var NewHeight;
	var NewWidth;
	
	if (oPopup.document.body.scrollHeight> 140 || oPopup.document.body.scrollWidth> 92)
		{
		NewHeight = oPopup.document.body.scrollHeight;
		NewWidth = oPopup.document.body.scrollWidth;
		oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
		}

	
	}
	else
		DisplayVersion(strID,RootID);
	

}

function DisplayChanges(ID){
	var strID
	strID = window.showModalDialog("../Properties/FT.asp?ID=" + ID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.reload(true);
		}
		
}

function DisplayPartNumbers(VersionID, RootID){
	var strID
	strID = window.showModalDialog("DisplayPartNumbers.asp?RootID=" + RootID + "&VersionID=" + VersionID,"","dialogWidth:600px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;maximize:No;resizable: Yes;status: No") 
	if (typeof(strID) != "undefined")
		{
		window.location.reload(true);
		}
		
}

function GetVersion(VersionID){
	//var strPath = trim(document.all("DelPath" + VersionID).innerText);    
	//window.open ("file://" + strPath);
	window.open ("../FileBrowse.asp?ID=" + VersionID);

}


function Export(strID){
	window.open (window.location.href + "&Type=" + strID);
}

function instr(MyString,Find){
	return MyString.indexOf(Find,0);
	
}

function ShowLanguage(strLang){
	var strPage = window.location.href;
	
	if (instr(strPage,"&ImageLang=") > -1)	
		window.location.href = strPage.substr(0,instr(strPage,"&ImageLang=")) + "&ImageLang=" + strLang;
	else
		window.location.href = strPage + "&ImageLang=" + strLang;

}


//-->
</SCRIPT>
</HEAD>

<STYLE>
.InfoTable TD{
	FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}

A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
A
{
    COLOR: blue
}

td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }

</STYLE>


<BODY >




<%


	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function

	function ShowImages (strImageList, strFilterList)
		dim FilterArray
		dim i
		if strImageList = "" or strFilterList="" then
			ShowImages = true
		else
			FilterArray = split(strFilterList,",")
			ShowImages = false
			for i = lbound(FilterArray) to ubound(FilterArray)
				if instr(", " & strImageList & ",", ", " &  FilterArray(i) & ",") > 0 then
					ShowImages = true
					exit for	
				end if
			next
		end if
	end function


	dim cn
	dim rs
	dim cn2
	dim rs2
	dim cm
	dim p
	dim strSR
	dim strARCD
	dim strDRDVD
	dim strWEB
	dim strDIB
	dim strPreinstall
	dim strPreinstallBrand
	dim strPreload	
	dim strPreloadBrand
	dim strRoyalty	
    dim strHPPI
	dim strTray
	dim strPanel
	dim strInfoCenter
	dim strDesktop
	dim strMenu
	dim strTile
	dim strTaskbarIcon
	dim blnLoadOK
	dim InfoRow
	dim strCert
	dim strSWSetupCat
	dim strProductName
	dim strImgLang
	dim strLanguages
	dim strCurrentLanguage
	dim LangArray
	dim strLangLinks
	dim	HideDTIcon
	dim strDashType
	dim ImagesToShow
	dim strProductID
	dim strPreinstallTeam
	dim strConveyorServerID 
	dim strlockeddeliverableList
	dim blnImageSnapshotsSaved
	dim strStatusID
	dim strDevCenter
    dim intProductStatus
	
	strDevCenter = ""
	strDashType=""
    intProductStatus = 0

	ImagesToShow = request("ImageFilter")
	if ImagesToShow <> "" then
		ImagesToShow = "," & ImagesToShow & ","
	end if
	HideDTIcon = false
	
	strLanguages = ""

	if request("ID") <> "" or request("ProdID") <> "" then
		blnLoadOk = true
	else
		blnloadok = false
	end if


	if blnLoadOK then
		'Create Database Connection
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.Recordset")
		
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		strDaveEmail = ""
					

		function ScrubSQL(strWords) 
	
			dim badChars 
			dim newChars 
			dim i
			
			strWords=replace(strWords,"'","''")
			
			badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
			newChars = strWords 
			
			for i = 0 to uBound(badChars) 
				newChars = replace(newChars, badChars(i), "") 
			next 
			
			ScrubSQL = newChars 
		
		end function 

		function FormatSKUList(strImages)
			dim strDash
			dim strLanguages
			dim strImageString
			
			FormatSKUList = ""
			if instr(strImages,":") > 0 then
				strImageString = left(strImages,instr(strImages,":")-1)
			else
				strImageString = strImages
			end if
			if right(strImageString,1) = "," then
				strImageString = left(strImageString,len(strImageString) -1)
			end if

			
			if strImages <> "" then
		
				set rs2 = server.CreateObject("ADODB.Recordset")
				rs2.Open "Select i.id, idef.SKUNumber, r.OSLanguage, r.Otherlanguage, r.dash, i.Priority " & _
						"from images i with (NOLOCK), ImageDefinitions idef with (NOLOCK), Regions r with (NOLOCK) " & _
						"where i.ImageDefinitionID = idef.id " & _
						"and r.id = i.regionid " & _
						"and idef.active = 1 " & _
						"and idef.statusid < 2 " & _
						"and i.ID in (" & ScrubSQL(strImageString) & ") ",cn,adOpenForwardOnly
				
				do while not rs2.EOF
					if instr(strImages,rs2("ID") & "=") > 0 then
						strlanguages = mid(strImages,instr(strImages,rs2("ID") & "=")+ len(rs2("ID")& "") + 1)
						strLanguages = "(" & left(strLanguages,instr(strLanguages,")"))
					else
						strLanguages = ""	
					end if
					if right(rs2("SKUNUmber") & "",2) = "##" then
						'if request("Type")<> "" then
                            if instr(FormatSKUList, ", " & rs2("SKUNumber")) = 0 then
							    FormatSKUList = FormatSKUList & ", " & rs2("SKUNumber") & strLanguages & ""
                            end if
						'else
                        '    if instr(FormatSKUList, rs2("SKUNumber")) = 0 then
    					'		FormatSKUList = FormatSKUList & ", <a class=""link"" href=""DeliverableMatrix.asp?ID=" & rs2("id") & "&PINTest=" & request("PINTest") & """>" & rs2("SKUNumber") & strLanguages & "</a>"
	                    '    end if
    					'end if
					elseif instr(rs2("SKUNUmber") & "","xx")=0 then
						if request("Type")<> "" then
							FormatSKUList = FormatSKUList & ", " & rs2("ID") & strLanguages & ""
						else
							FormatSKUList = FormatSKUList & ", <a class=""link"" href=""DeliverableMatrix.asp?ID=" & rs2("id") & "&PINTest=" & request("PINTest") & """>" & rs2("ID") & strLanguages & "</a>"
						end if
					else	
						strDash = rs2("Dash") & ""
						if strDash <> "" then
							strDash = left(mid(strDash,2),2)
							if request("Type")<> "" then
								FormatSKUList = FormatSKUList & ", " & replace(rs2("SKUNumber"),"xx",strDash) & strLanguages & ""
							else
								FormatSKUList = FormatSKUList & ", <a class=""link"" href=""DeliverableMatrix.asp?ID=" & rs2("id") & "&PINTest=" & request("PINTest") & """>" & replace(rs2("SKUNumber"),"xx",strDash) & strLanguages & "</a>"
							end if
						else
							if request("Type")<> "" then
								FormatSKUList = FormatSKUList & ", " & rs2("SKUNumber") & strLanguages & ""
							else
								FormatSKUList = FormatSKUList & ", <a class=""link"" href=""DeliverableMatrix.asp?ID=" & rs2("id") & "&PINTest=" & request("PINTest") & """>" & rs2("SKUNumber") & strLanguages & "</a>"
							end if
						end if
					end if
	
					rs2.MoveNext
				loop
				rs2.Close
				set rs2 = nothing
				if len(FormatSKUList) > 0 then
					FormatSKUList = mid(FormatSKUList,3)
				else
					FormatSKUList = "No Images in Development"'"All Images Have Been Released to the Factory"
				end if
			end if
		end function



		dim CurrentUser 
		dim CurrentUserName
		dim CurrentUserGroup
		dim CurrentUserID
	
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

'		rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
		if (rs.EOF and rs.BOF) then
			set rs = nothing
        	set cn=nothing
        	Response.Redirect "../NoAccess.asp?Level=0"
        else
            CurrentUserPartner = rs("PartnerID")
			CurrentUserName = rs("Name") & ""
			CurrentUserID = rs("ID")
			CurrentUserGroup = rs("WorkgroupID")
		end if
		rs.Close

		'Verify Access is OK. Partner 1 is HP and 9 is Moduslink which can access any product
		if trim(CurrentUserPartner) <> "1" and trim(CurrentUserPartner) <> "9" then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetProductPartner"
			
	
			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			if request("ProdID") = "" then
				p.Value = 0
			else
				p.Value = request("ProdID")
			end if
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@ImageID", 3, &H0001)
			if request("ID") = "" then
				p.Value = 0
			else
				p.Value = request("ID")
			end if
			cm.Parameters.Append p
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				rs.close	
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../NoAccess.asp?Level=0"
			end if
			rs.close	
		end if

		InfoRow = ""
		if request("ProdID") <> "" then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetProductBrands"
		

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("ProdID")
			cm.Parameters.Append p
	

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			'rs.open "spGetProductBrands " & request("ProdID"),cn,adOpenForwardOnly
			HideDTIcon = true
			do while not rs.EOF
				if rs("BusinessID")=2  then 'or (lcase(trim(rs("Name"))) <> "hp nc" and lcase(trim(rs("Name"))) <> "hp nw" and lcase(trim(rs("Name"))) <> "hp nx") then
					HideDTIcon = false
					exit do
				end if
				rs.MoveNext
			loop
			rs.Close
		else
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetImageProperties"
			

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("ID")
			cm.Parameters.Append p
		

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

		
			'rs.Open "spGetImageProperties " & request("ID"),cn,adOpenForwardOnly
			if rs.eof and rs.BOF then	
				blnLoadOK = false	
			else
				if trim(rs("Skunumber") & "") = "" then
					strImageID = rs("ID")
				else
					if instr(rs("Skunumber"),"-xx")>0 then
						strImageID = replace(rs("Skunumber"),"-xx","-" & mid(rs("Dash"),2,2))
					else
						strImageID = rs("Skunumber") & rs("Dash")
					end if
				end if

				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spGetMUILanguages " & rs("OSID") & "," & rs("RegionID"),cn,adOpenForwardOnly
				if rs2.eof and rs2.bof then
					strMUILanguages = ""
				else
					strMUILanguages = rs2("MUILanguages") & "&nbsp;"
				end if
				rs2.close		
				set rs2=nothing


				on error resume next

				'Load Houston Conevyor server to check SKU Rev	
				strConveyorServerID = 0
				set rs2 = server.CreateObject("ADODB.recordset")
				set cn2 = server.CreateObject("ADODB.Connection")
				cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=houbnbcvr01.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;"
				cn2.Open
				rs2.Open "select max(skurevision) as Rev from sku with (NOLOCK) where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
				if not (rs2.EOF and rs2.BOF) then
					strSKURevision =  rs2("rev") & ""
					strConveyorServerID = 1
					if trim(strSKURevision) = "" then
						strSKURevision = "0"
					end if
				else
					strSKURevision =  "0"
				end if
				rs2.Close
				set rs2=nothing
				cn2.Close
				set cn2=nothing

				'Load TDC Conevyor server to check SKU Rev	
				set rs2 = server.CreateObject("ADODB.recordset")
				set cn2 = server.CreateObject("ADODB.Connection")
				cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=16.159.144.23" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
				'cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=tpopsgcvr3.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
				cn2.Open
				rs2.Open "select max(skurevision) as Rev from sku with (NOLOCK) where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
				if not (rs2.EOF and rs2.BOF) then
					if trim(rs2("rev") & "") <> "" then
						if isnumeric(rs2("rev") & "") then
							if clng(rs2("rev") & "") > clng(strSKURevision) then
								strSKURevision =  rs2("rev") & ""
								strConveyorServerID = 2
							end if
						end if
					end if
				end if
				rs2.Close
				set rs2=nothing
				cn2.Close
				set cn2=nothing

				'Load CDC Conevyor server to check SKU Rev	
'				set rs2 = server.CreateObject("ADODB.recordset")
'				set cn2 = server.CreateObject("ADODB.Connection")
'				cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=shgcdccvr01.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
'				cn2.Open
'				rs2.Open "select max(skurevision) as Rev from sku where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
'				if not (rs2.EOF and rs2.BOF) then
'					if trim(rs2("rev") & "") <> "" then
'						if isnumeric(rs2("rev") & "") then
'							if clng(rs2("rev") & "") > clng(strSKURevision) then
'								strSKURevision =  rs2("rev") & ""
'								strConveyorServerID = 5
'							end if
'						end if
'					end if
'				end if
'				rs2.Close
'				set rs2=nothing
 '   			cn2.Close
	'			set cn2=nothing
		if false then
			on error goto 0
				'Load Brazil Conveyor server to check SKU Rev	
				set rs2 = server.CreateObject("ADODB.recordset")
				set cn2 = server.CreateObject("ADODB.Connection")
				cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;User ID=viewer;Initial Catalog=conveyor;Data Source=BRAHPQCVR01.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
				cn2.Open
				rs2.Open "select max(skurevision) as Rev from sku with (NOLOCK) where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
				if not (rs2.EOF and rs2.BOF) then
					if trim(rs2("rev") & "") <> "" then
						if isnumeric(rs2("rev") & "") then
							if clng(rs2("rev") & "") > clng(strSKURevision) then
								if isnull(rs2("rev")) then
									Response.Write "1" & "<BR>"
								end if
								strSKURevision =  rs2("rev") & ""
								Response.Write ">" & strSKURevision & "<"
								strConveyorServerID = 4
							end if
						end if
					end if
				end if
				rs2.Close
				set rs2=nothing
				cn2.Close
				set cn2=nothing
		end if


'				'Load Conevyor Singapore server to check SKU Rev	
'				set rs2 = server.CreateObject("ADODB.recordset")
'				set cn2 = server.CreateObject("ADODB.Connection")
'				cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=sgpacccvr01.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
'				cn2.Open
'				rs2.Open "select max(skurevision) as Rev from sku where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
'				if not (rs2.EOF and rs2.BOF) then
'					if trim(rs2("rev") & "") <> "" then
'						if isnumeric(rs2("rev") & "") then
'							if clng(rs2("rev") & "") > strSKURevision then
'								strSKURevision =  rs2("rev") & ""
'								strConveyorServerID = 3
'							end if
'						end if
'					end if
'				end if
'				rs2.Close
'				set rs2=nothing
'				cn2.Close
'				set cn2=nothing
'				
'				'Load Conevyor TDC server to check SKU Rev if Rev does not exist on Houston server
'				if trim(strSKURevision) = "" then
'					set cn2 = server.CreateObject("ADODB.Connection")
'					cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=16.159.144.23" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;" '16.169.185.60
''					cn2.ConnectionString = "Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;User ID=viewer;Initial Catalog=conveyor;Data Source=tpopsgcvr3.auth.hpicorp.net" & ";Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;" '16.169.185.60
'					cn2.Open
'				
'					rs2.Open "select max(skurevision) as Rev from sku where skunumber = '" & ScrubSQL(strImageID) & "' and locked <> 0 AND Enabled <> 0;",cn2,adOpenForwardOnly
'					if not (rs2.EOF and rs2.BOF) then
'						strSKURevision =  rs2("rev") & ""
'					else
'						strSKURevision = "Unknown"
'					end if
'					rs2.Close
'				end if
'
'				set cn2=nothing
'				set rs2=nothing
			
				on error goto 0
			
				HideDTIcon = false
				'if (lcase(trim(rs("Brand"))) = "hp nc" or lcase(trim(rs("Brand"))) = "hp nw" or lcase(trim(rs("Brand"))) = "hp nx")   then
				if rs("Devcenter") <> 2 then
					if rs("OSID") <> 13 then 'Never Hide Win2k
						HideDTIcon = true
					end if
				end if
							

				strlockeddeliverableList = trim(rs("lockeddeliverableList") & "")
				blnImageSnapshotsSaved = trim(rs("ImageSnapshotsSaved") & "")
				strStatusID = trim(rs("StatusID") & "")
				if not isnumeric(strStatusID) then
					strStatusID = "0"
				end if
				
				InfoRow = InfoRow & "<TR><td><b>Brand: </b>" & rs("Brand") & "</TD>"
				InfoRow = InfoRow & "<TD><b>OS: </b>" & rs("OS") & "</TD>"
				InfoRow = InfoRow & "<TD><b>Apps: </b>" & rs("SW") & "</TD>"
				InfoRow = InfoRow & "<TD><b>Type: </b>" & rs("ImageType") & "</TD>"
				if trim(strDevCenter) = "2" then
					if trim(rs("OtherLanguage") & "") <> "" then
						strLanguages = rs("OSLanguage") & "," & rs("OtherLanguage")
						LangArray = split(strLanguages,",")
						
						if request("Type") <> "" then
							strLangLinks = strLanguages			
						else
							strLangLinks = ""
							for i = lbound(LangArray) to ubound(LangArray)
								if trim(ucase(LangArray(i))) = trim(ucase(request("ImageLang"))) then
									strLangLinks = strLangLinks & "," & trim(ucase(LangArray(i)))
								else
									strLangLinks = strLangLinks & ",<a href=""javascript:ShowLanguage('" & LangArray(i) & "')"">" & trim(ucase(LangArray(i))) & "</a>"
								end if
							next
							if strLangLinks <> "" then
								strlangLinks = mid(strLangLinks,2)
							end if
						end if
						InfoRow = InfoRow & "<TD><b>Language: </b>" & strLangLinks & "</TD>"
					else
						InfoRow = InfoRow & "<TD><b>Language: </b>" & rs("OSLanguage") & ""
						strLanguages = rs("OSLanguage")
					end if
				end if
				InfoRow = InfoRow & "</tr>" 
				InfoRow = InfoRow & "<tr><TD><b>Region: </b>" & rs("Region") & "</TD>"
				InfoRow = InfoRow & "<TD><b>Dash: </b>" & rs("Dash") & "</TD>"
				InfoRow = InfoRow & "<TD><b>Country Code: </b>" & rs("CountryCode") & "</TD>"
				strDashType = trim(rs("ImageType") & "")
				InfoRow = InfoRow & "<TD ><b>RTMDate: </b>" & rs("RTMDate") & "</TD>"
				
				
				if trim(strDevCenter) = "2" then
					InfoRow = InfoRow & "<TD colspan=4><b>Option Config: </b>" & rs("OptionConfig") & "</TD>"
				end if
				InfoRow = InfoRow & "</tr>" 
				if trim(strDevCenter) = "2" then
					if trim(strMUILanguages) <> "" then
						InfoRow = InfoRow & "<TR><TD colspan=2><b>Comments: </b>" & rs("Comments") & "</td><TD colspan=2><b>MUI Languages:</b> " & strMUILanguages & "</TD><TD colspan=1><b>Status: </b>" & rs("Status") & "</td></TR>" 
					else
						InfoRow = InfoRow & "<TR><TD colspan=4><b>Comments: </b>" & rs("Comments") & "</td><TD colspan=1><b>Status: </b>" & rs("Status") & "</td></TR>" 
					end if
				else
					if trim(strMUILanguages) <> "" then
						InfoRow = InfoRow & "<TR><TD colspan=1><b>Comments: </b>" & rs("Comments") & "</td><TD colspan=2><b>MUI Languages:</b> " & strMUILanguages & "</TD><TD colspan=1><b>Status: </b>" & rs("Status") & "</td></TR>" 
					else
						InfoRow = InfoRow & "<TR><TD colspan=3><b>Comments: </b>" & rs("Comments") & "</td><TD colspan=1><b>Status: </b>" & rs("Status") & "</td></TR>" 
					end if
				end if
			end if
			rs.Close
		end if
		
			if blnLoadok then
				if request("ProdID") <> "" then
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetProductVersionName"
		

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = request("ProdID")
					cm.Parameters.Append p
	

					rs.CursorType = adOpenForwardOnly
					rs.LockType=AdLockReadOnly
					Set rs = cm.Execute 
					Set cm=nothing

				
					'rs.Open "spGetProductVersionName " & request("ProdID"),cn,adOpenForwardOnly
					if rs.EOF and rs.BOF then
						strproductName = ""
						strDevCenter = ""
						strPreinstallTeam = 0
                        intProductStatus = 0
					else
						strDevCenter = rs("DevCenter")
						strproductName = rs("Name") & ""
						strPreinstallTeam = rs("PreinstallTeam") & ""
                        intProductStatus = rs("ProductStatusID")
					end if	
					rs.Close
					strProductID = request("ProdID")
				else
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetProductByImage"
		
		
					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = request("ID")
					cm.Parameters.Append p
	

					rs.CursorType = adOpenForwardOnly
					rs.LockType=AdLockReadOnly
					Set rs = cm.Execute 
					Set cm=nothing

					'rs.Open "spGetProductByImage " & request("ID"),cn,adOpenForwardOnly
					if rs.EOF and rs.BOF then
						strproductName = ""
						strPreinstallTeam = 0
						strDevCenter = ""
					else
						strDevCenter = rs("DevCenter")
						strproductName = rs("Product") & ""
						strProductID = rs("ID")
						strPreinstallTeam = rs("PreinstallTeam") & ""
					end if	
					rs.Close

				end if
%>



<% if request("DIBOnly") = "" then %>
<center><font size=3 face=verdana><b><%=strproductName%> Deliverable Matrix 
    <%if request("PINTest") = "1" then%>
        <font size=2 face=verdana color="red"><br />PIN Test Matrix<BR /><BR />WARNING: This matrix has not been released and is for Preinstall team use only.</font>
    <%end if%>
<% else %>
<center><font size=3 face=verdana><b><%=strproductName%> DIB Deliverables 
<% end if 

'Only display the DIB section when the request was for DIB only
if request("DIBOnly") = "" then
strOtherServerName = ""
	if clng(strConveyorServerID) <> clng(strPreinstallTeam) and clng(strPreinstallTeam) <> 0 and clng(strConveyorServerID)<> 0 then
		
		strOtherServerName = "<font color=red> - "
		select case strConveyorServerID
			case 1
			strOtherServerName = strOtherServerName & "Houston"
			case 2
			strOtherServerName = strOtherServerName & "Taiwan"
			case 3
			strOtherServerName = strOtherServerName & "Singapore"
			case 4
			strOtherServerName = strOtherServerName & "Brazil"
			case 5
			strOtherServerName = strOtherServerName & "CDC"
		end select
		strOtherServerName = strOtherServerName & " Server</font>"
	end if

if request("ID") <> "" then%>
	<%if strSKURevision = "0" then%>
		<BR><BR>Image <%=strImageID%> (<font color=red>Rev Unknown</font>)
	<%else%>
		<BR><BR>Image <%=strImageID%> (Rev <%=strSKURevision & strOtherServerName%>)
	<%end if%>
	
	<%	if request("ImageLang") <> "" then
			Response.Write "(" & request("ImageLang") & ")"
		end if%>
<%end if%>



<%	if clng(strStatusID) > 1 and  not (blnImageSnapshotsSaved and trim(strlockeddeliverableList & "") <> "") then 													
		Response.Write "<BR><BR><font size=1 face=verdana color=red>Warning: This deliveable list may be inaccurate because this image was not locked when it was released to the factory or deactivated.</font><BR>"					
	end if
%>

</b></font><BR><BR>
	<%if strproductName = "Ford 1.1" and request("prodID") = "" then%>
		<center><font size=2 face=verdana color=red>Note: This image is shared with Thruman 1.1</font></center><BR>
	<%elseif trim(request("ProdID")) = "267" then%>
		<center><font size=2 face=verdana color=red>Note: Image deliverables are shared with Thruman 1.1</font></center><BR>
	<%elseif trim(request("ProdID")) = "268" then%>
		<center><font size=2 face=verdana color=red>Note: These image deliverables only reflect the Thurman 1.1 SMB images. Please refer to Ford 1.1 for consumer image deliverables.</font></center><BR>
	<%elseif trim(request("ProdID")) = "349" then%>
		<center><font size=2 face=verdana color=red>Note: Titan and Altima are using shared image strategy. Please refer to Titan Excalibur for the up-to-date Software and Image deliverables targets.</font></center><BR>
	<%end if%>
<font size=1 face=verdana><%=Now()%></font><BR><BR>
<%if request("FilterName") <> "" then%>
	<font size=1 face=verdana>Displaying deliverables for selected images only: <%=request("FilterName")%></font><BR><BR>
<%end if%>
</center>
<%if request("ID") <> "" then%>
	<font size=2 face=verdana><b>Image Details:</b></font><BR>
<%end if%>
<table width=100% class=InfoTable border=1 bgcolor=beige cellspacing=0 cellpadding=3>
<%=InfoRow%>
</table><BR>

<%
	dim strMatrixHeader	
	if request("ID") <> "" then
		strMatrixHeader	= "Deliverables in this Image" 
	else
		strMatrixHeader	= "Software Deliverables" 
	end if

    dim OSLinks
    OSLinks = ""

    if intProductStatus < 3 and trim(request("ProdID")) <> "" then
         if request("OSFamilyID") = "12" then
            OSLinks = OSLinks & "&nbsp;&nbsp;&nbsp;Image OS: <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=10&FilterName=Win7 Images"">Win7</a> | Win8 | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=18&FilterName=Win10 Images"">Win10</a> |<a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & """>All</a>&nbsp;&nbsp;"
        elseif request("OSFamilyID") = "10" then
            OSLinks = OSLinks & "&nbsp;&nbsp;&nbsp;Image OS: Win7 | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=12&FilterName=Win8 Images"">Win8</a> | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=18&FilterName=Win10 Images"">Win10</a> |<a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & """>All</a>&nbsp;&nbsp;"
        elseif request("OSFamilyID") = "18" then
            OSLinks = OSLinks & "&nbsp;&nbsp;&nbsp;Image OS: <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=10&FilterName=Win7 Images"">Win7</a>  | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=12&FilterName=Win8 Images"">Win8</a> | Win10 | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & """>All</a>&nbsp;&nbsp;"
		else
            OSLinks = OSLinks & "&nbsp;&nbsp;&nbsp;Image OS: <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=10&FilterName=Win7 Images"">Win7</a> | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=12&FilterName=Win8 Images"">Win8</a> | <a href=""deliverablematrix.asp?ProdID=" & request("ProdID") & "&PINTest=" & request("PINTest") & "&OSFamilyID=18&FilterName=Win10 Images"">Win10</a> | All&nbsp;&nbsp;"
        end if
    end if

%>
<%if request("Type")= "" then%>
<table width=100%><TR><TD align=left><font size=2 face=verdana><b><%=strMatrixHeader%></b></font><font size=1 face=verdana><%=OSLinks%></font><BR></td><td valign=bottom align=right><font size=1 face=verdana>&nbsp;Export:&nbsp;<a href="javascript: Export(1);">Excel</a><!--&nbsp;|&nbsp;<a href="javascript: Export(2);">Word</a>--></td></tr></table>
<%else%>
<font size=2 face=verdana><b><%=strMatrixHeader%></b></font>
<font size=1 face=verdana><BR></font>
<%end if
'end if for DIB only deliverables
end if
%>


<!--<TABLE border=1 cellspacing=0 cellpadding=1><TR style="HEIGHT: 20px" height=20><TD bgcolor=powderblue><font size=1 face=verdana>&nbsp;&nbsp;Versions&nbsp;Targeted&nbsp;But&nbsp;Not&nbsp;In&nbsp;Image&nbsp;&nbsp;</font></TD><TD bgcolor=mistyrose ><font size=1 face=verdana>&nbsp;&nbsp;Versions&nbsp;Being&nbsp;Removed&nbsp;From&nbsp;Image&nbsp;&nbsp;</font></TD><TD bgcolor=ivory ><font size=1 face=verdana>&nbsp;&nbsp;Versions&nbsp;Targeted&nbsp;<u>and</u>&nbsp;In&nbsp;Image&nbsp;&nbsp;</font></TD></TR></TABLE>-->




<%	

	dim strDIBRows
	dim strWebRows
	dim strOOCDIBRows
	dim strDRCDRows
	dim strOOCWebRows
	dim strRowData
	dim strAllNotes
	dim strOOC
	dim strLastCategory
	dim strWebOSList
	dim strDistributionChange
	dim DistributionChangeCount
	dim strMultiLanguage
	dim strCDPartNumber
	dim strKitNumber
	dim blnHasPath
	
	'Setup table row HTML for seconday color cells
	PreRow = "<TR bgcolor=Ivory valign=top><TD>"
	MidRow = "</TD><TD valign=top  class=""cell"">"
	MidRowNR = "</TD><TD nowrap valign=top  class=""cell"">"
	MidRowCenter = "</TD><TD align=center valign=top class=""cell"">"
	MidRowCenterBorder = "</TD><TD align=center valign=top style=""BORDER-LEFT: gray thin solid;"" class=""cell"">" 
	PostRow = "</FONT></TD></TR>"

if request("DIBOnly") = "" then
	if request("ProdID") = "" then
		strImgLang= "<TD align=center width=20 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
	else
		strImgLang = ""
	end if
	
	Response.Write "<TABLE ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
	if blnImageSnapshotsSaved and trim(strlockeddeliverableList & "") <> "" then 													
		Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
	else
		Select Case request("Type")
			Case 1
				Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD rowspan=2 align=left width=80 valign=bottom><font face=verdana color=black size=1><b>Developer Approved</b></font></TD><TD colspan=7 align=right width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD rowspan=2 valign=bottom><font face=verdana color=black size=1><b>HPPI</b></font></TD><TD rowspan=2 valign=bottom><font face=verdana color=black size=1><b>Royalty Bearing</b></font></TD><TD colspan=7 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>SWSetup</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=200 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
				Response.Write "<tr height=52 bgcolor=beige><TD valign=bottom><font face=verdana color=black size=1><b>Preinstall</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>Preload</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>DIB</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>WEB</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>SR</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>DRCD</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>DRDVD</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><font face=verdana color=black size=1><b>CTL</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>Desktop</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>StartMenu</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>SysTray</b></font></TD><TD valign=bottom><font face=verdana color=black size=1><b>InfoCenter</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Start Menu Tile</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Task Bar Icon</b></font></TD></tr>"
			Case 3
				Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD></tr>"
			Case Else
				Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD width=26 rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_developerApproved.GIF""></TD><TD style=""BORDER-LEFT: gray thin solid;"" colspan=7 align=right width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD rowspan=2 valign=bottom><IMG SRC=""../images/col_HPPI.gif""></TD><TD rowspan=2 valign=bottom><IMG height=45 width=24 width=11 SRC=""../images/col_Royalty.gif""></TD><TD colspan=7 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>SWSetup</b></font></TD><TD align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=200 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
				Response.Write "<tr height=52 bgcolor=beige><TD valign=bottom style=""BORDER-LEFT: gray thin solid;""><IMG height=56 width=14 SRC=""../images/col_Prinstall.gif""></TD><TD valign=bottom><IMG height=48 width=15 SRC=""../images/col_Preload.gif""></TD><TD valign=bottom><IMG height=29 width=12 SRC=""../images/col_Dib.gif""></TD><TD valign=bottom><IMG height=29 width=14 SRC=""../images/col_Web.gif""></TD><TD valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD valign=bottom><IMG height=44 width=11 SRC=""../images/col_DRDVD.gif""></TD><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD><TD valign=bottom><IMG SRC=""../images/col_infocenter.GIF""></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Start Menu Tile</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Task Bar Icon</b></font></TD></tr>" 
		End Select
	end if
end if ' for DIB only	
	if request("ProdID") <> "" and trim(request("OSFamilyID")) <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListDeliverablesInImageByOSFamily"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSFamilyID", 3, &H0001)
        p.Value = clng(request("OSFamilyID"))
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@PINTest", 3, &H0001)
		if request("PINTest") = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p	

        Set p = cm.CreateParameter("@latestML", 3, &H0001)
		p.Value = 1
		cm.Parameters.Append p	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
    elseif request("ProdID") <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListDeliverables4Product"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@PINTest", 3, &H0001)
		if request("PINTest") = "1" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p	

        Set p = cm.CreateParameter("@latestML", 3, &H0001)
		p.Value = 1
		cm.Parameters.Append p	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

'		rs.Open "spListDeliverables4Product " & request("ProdID") ,cn,adOpenForwardOnly
	else
		if blnImageSnapshotsSaved and trim(strlockeddeliverableList & "") <> "" then 
			if instr(strlockeddeliverableList & "",":")> 0 then				strSQl = left(strlockeddeliverableList & "",instr(strlockeddeliverableList & "",":")-2)			else				strSQL = left(strlockeddeliverableList & "",len(strlockeddeliverableList & "") -1)			end if
					
			strSQl = "Select r.hppi,v.ID, v.DeliverableName, v.Version, v.Revision, v.Pass, e.name as Developer, r.categoryid, v.imagepath, r.id as RootID, r.royaltyBearing, r.icondesktop, r.iconinfocenter, r.iconpanel, r.iconmenu, r.certrequired, r.icontray, r.IconTile, r.IconTaskbarIcon, v.vendorversion, v.irspartnumber, v.partnumber, c.otscategory as Category,r.typeid, 1 as InOfficialImage, 0 as DeveloperNotificationStatus, 0 as DIBHWReq, v.CDKitNumber, v.CDPartnumber, '' as targetNotes, '' as PreinstallNotes, '' as ImageSummary, 0 as CertificationStatus, 1 as Preinstall, 0 as preload, 1 as Targeted, 0 as OSID,  0 as RACD_Americas, 0 as OOCRelease, '' as DistributionChange, 0 as OSCD, 0 as DocCD,0 as RACD_EMEA, 0 as RACD_APD, '' as PreinstallBrand, '' as PreloadBrand, 0 as Patch, 0 as ImageID, 0 as DRDVD, 0 as DropInBox, 0 as Web, 0 as ProdID, 0 as ARCD, 0 as selectiverestore, 1 as inimage, r.SWSetupCategoryID, '' as Images " & _					 "From DeliverableVersion v with (NOLOCK)" & _
					 "inner join deliverableroot r on v.deliverablerootid = r.id" & _
					 "inner join deliverablecategory c on c.id = r.categoryid" &_
					 "inner join employee e on e.id = v.developerid" & _
					 "where v.id in (" & strSQL & ") "  & _
					 "order by c.otscategory, v.deliverablename, v.id desc "			rs.open strSQL, cn,adOpenStatic
		else
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spListDeliverablesInImage"
		

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("ID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@PINTest", 3, &H0001)
			if trim(request("PINTest")) = "1" then
                p.Value = 1
            else
                p.Value = 0
            end if
			cm.Parameters.Append p
	
            Set p = cm.CreateParameter("@latestML", 3, &H0001)
		    p.Value = 1
		    cm.Parameters.Append p	

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
		end if
'		rs.Open "spListDeliverablesInImage " & request("ID") ,cn,adOpenForwardOnly
	end if
	DistributionChangeCount = 0
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD colspan=12><font face=Verdana size=2>No deliverables found for this program.</font>"
	else
	
		set rs2 = server.CreateObject("ADODB.Recordset")
		rs2.open "spListProductOS " & clng(strProductID) & ",2"  ,cn,adOpenForwardOnly
		strWebOSList = ""
		do while not rs2.eof
			strWebOSList = strWebOSList & "," & trim(rs2("id"))
			rs2.movenext	
		loop
		rs2.close
		set rs2=nothing
		LastBucket = ""
		strDIBRows = ""
		strWebRows = ""
		strOOCDIBRows = ""
		strOOCWebRows = ""
		strRowData = ""
		strDRCDRows = ""
		strAllNotes = ""
		strDistributionChange = ""
		strMultiLanguage = ""
		strCDPartNumber = ""
		strKitNumber = ""		
		
		dim HideTools
		if CurrentUSerGroup=23 or CurrentUSerGroup=12 or request("HideTools") <> "" then
			HideTools = true
		else
			HideTools = false
		end if		
		strDaveEmail = strDaveEmail & "<BR>RS Opened<BR>"
		do while not rs.EOF
			'	if currentuserid=31 then
             '       response.write "<BR>." & rs("name") & "_" & rs("InIMage")
              '  end if
			strDaveEmail = strDaveEmail & rs("ID") & "<BR>"
			'if (not((HideTools) and (rs("CategoryID")=137 or rs("CategoryID")=170 or rs("CategoryID")=171))) and rs("Name") <> "Test Deliverable"  and ShowImages(trim(rs("Images") & ""),ImagesToShow ) and (((rs("Preinstall") or rs("Preload") or rs("SelectiveRestore") or rs("ARCD") or rs("DRDVD") or rs("DropInBox") or rs("Web")))  or (request("ProdID") <> "" and rs("TypeID") > 1 )     ) then
			if (not((HideTools) and (rs("CategoryID")=137 or rs("CategoryID")=170 or rs("CategoryID")=171))) and ShowImages(trim(rs("Images") & ""),ImagesToShow ) and (((rs("Preinstall") or rs("Preload") or rs("SelectiveRestore") or rs("ARCD") or rs("DRDVD") or rs("DropInBox") or rs("Web")))  or (request("ProdID") <> "" and rs("TypeID") > 1 )     ) then
				strDaveEmail = strDaveEmail & "."
                if request("ProdID") <> "" or (trim(rs("Images") & "") = "" or instr(", " & rs("Images") & "," , ", " & trim(rs("ImageID")) & ",") > 0) then
					strDaveEmail = strDaveEmail & "."
					CurrentCategory = trim(rs("Category") & "")
					'if instr(CurrentCategory,"-")>0 then
					'	CurrentCategory = left(CurrentCategory,instr(CurrentCategory,"-")-1)
					'end if
				
					strPreinstallBrand = trim(rs("PreinstallBrand") & "")
					strPreloadBrand = trim(rs("PreloadBrand") & "")
				
					if strLastCategory <> CurrentCategory and request("DIBOnly") = "" and (rs("Preinstall") or rs("Preload") or rs("SelectiveRestore"))then
						if blnImageSnapshotsSaved and trim(strlockeddeliverableList & "") <> "" then 													
							Response.Write "<TR bgcolor=#cccc99><TD colspan=8><font size=1 face=verdana>" & CurrentCategory & "</font></TD></TR>" 
						else
							Response.Write "<TR bgcolor=#cccc99><TD colspan=28><font size=1 face=verdana>" & CurrentCategory & "</font></TD></TR>" 
						end if
						strLastCategory = CurrentCategory
					end if

					strDistribution = ""
					if trim(rs("TypeID") & "") = "1" then
						strDistribution = "<b>-</b>"
					else
						if rs("Preinstall") then
							strDistribution = ",Preinstall"
							if strPreinstallBrand <> "" then
								strDistribution = strDistribution & "(" & strPreinstallBrand & ")"
							end if
						end if
						if rs("Preload") then
							strDistribution = strDistribution & ",Preload"
							if strPreloadBrand <> "" then
								strDistribution = strDistribution & "(" & strPreloadBrand & ")"
							end if
						end if	
						if rs("DropInBox") then
							strDistribution = strDistribution & ",DIB"
						end if
						if rs("Web") then
							strDistribution = strDistribution & ",Web"
						end if	
						if rs("DRDVD") then
							strDistribution = strDistribution & ",DRDVD"
						end if

						if rs("RACD_Americas") then
							strDistribution = strDistribution & ",RACD-Americas"
						end if

						if rs("RACD_EMEA") then
							strDistribution = strDistribution & ",RACD-EMEA"
						end if

						if rs("RACD_APD") then
							strDistribution = strDistribution & ",RACD-APD"
						end if

						if rs("OSCD") then
							strDistribution = strDistribution & ",OSCD"
						end if
					
						if rs("DocCD") then
							strDistribution = strDistribution & ",Doc CD"
						end if

						if trim(rs("Patch")&"") <> "0" then
							strDistribution = strDistribution & ",Patch"
						end if

						if trim(rs("DistributionChange") & "") <> "" then
							strDistribution = strDistribution & "<font color= ""red""><b>*</font></b>"
							DistributionChangeCount = DistributionChangeCount + 1
						end if
						
						if strDistribution <> "" then
							strDistribution = mid(strDistribution,2)
						else
							strDistribution = "&nbsp;"
						end if
					end if

					strOOC = rs("OOCRelease") & ""
			
					Version = rs("Version") & ""
					if rs("Revision") <> "" then
						Version  = Version & "," &  rs("Revision")
					end if
					if rs("Pass") <> "" then
						Version  = Version & "," &  rs("Pass")
					end if
					if rs("Targeted") then
						strTargeted = "X"
					else
						strTargeted = "&nbsp;"
					end if
					if rs("InImage") then
						strInImage = "X"
					elseif rs("Preinstall") or rs("Preload") or rs("SelectiveRestore") or rs("ARCD") or rs("DRDVD") then
						strInImage = "&nbsp;"
					else
						strInImage = "-"
					end if

					if rs("Preinstall") then
						if strPreinstallbrand <> "" and ucase(trim(strDashType)) <> "CTO" then
							strPreinstall = strPreinstallbrand
						else
							strPreinstall = "X"
						end if
					else
						strPreinstall = "&nbsp;"
					end if
					if rs("Preload") then
						if strPreloadbrand <> "" and ucase(trim(strDashType)) <> "CTO" then
							strPreload = strPreloadbrand
						else
							strPreload = "X"
						end if
					else
						strPreload = "&nbsp;"
					end if
					if rs("DropInBox") then
						strDIB = "X"
					else
						strDIB = "&nbsp;"
					end if
					if rs("Web") then
						strWEB = "X"
					else
						strWEB = "&nbsp;"
					end if	
					if rs("SelectiveRestore") then
						strSR = "X"
					else
						strSR = "&nbsp;"
					end if

					if rs("ARCD") then
						strARCD = "X"
					else
						strARCD = "&nbsp;"
					end if

					if rs("DRDVD") then
						strDRDVD = "X"
					else
						strDRDVD = "&nbsp;"
					end if

					if rs("HPPI") then
						strHPPI = "X"
					else
						strHPPI = "&nbsp;"
					end if

					if rs("RoyaltyBearing") then
						strRoyalty = "X"
					else
						strRoyalty = "&nbsp;"
					end if

					if false then
						if rs("RootID") = 8001 and rs("IconDesktop") then
							strDesktop = "X"
						elseif rs("OSID") = 13 and rs("RootID") = 3230 and rs("ID") >= 2398 then 'Help and support on Win2K always has icon
							strDesktop = "X"
						elseif rs("OSID") = 0 and rs("RootID") = 3230 and rs("ID") >= 2398 then 'Help and support on Win2K always has icon
							strDesktop = "2K"
						elseif rs("IconDesktop") and (not HideDTIcon) then
							strDesktop = "X"
						else
							strDesktop = "&nbsp;"
						end if
					
					else
						if rs("IconDesktop") and (not HideDTIcon) then
							strDesktop = "X"
						elseif rs("OSID") = 0 and rs("IconDesktop") and HideDTIcon then 'Commercial Product-Level
							strDesktop = "[1]"
						elseif rs("OSID") <> 0 and rs("IconDesktop") and HideDTIcon then 'Commercial Image-Level
							strDesktop = "[1]"
						else
							strDesktop = "&nbsp;"
						end if
					
					end if
					
					
'					if currentuserid=31 then
'						Response.Write rs("DeliverableName") & " - " & HideDTIcon & "<BR>" 
'					end if
					
					if rs("IconPanel") then
						strPanel = "X"
					else
						strPanel = "&nbsp;"
					end if
					
					if rs("IconInfoCenter") then
						strInfoCenter = "X"
					else
						strInfoCenter = "&nbsp;"
					end if					

					if rs("IconMenu") then
						strMenu = "X"
					else
						strMenu = "&nbsp;"
					end if

					if rs("IconTray") then
						strTray = "X"
					else
						strTray = "&nbsp;"
					end if
					
					if rs("IconTile") then
						strTile = "X"
					else
						strTile = "&nbsp;"
					end if
					
					if rs("IconTaskbarIcon") then
						strTaskbarIcon = "X"
					else
						strTaskbarIcon = "&nbsp;"
					end if
					
					
					if trim(rs("TypeID") & "") = "1" then
						strImageSummary = "<b>-</b>"
					else
						if trim(rs("ImageSummary") & "") = "" then
							strImageSummary = "ALL"
						else
							strImageSummary = rs("ImageSummary")
						end if
						if rs("Images") <> "" and request("ProdID") <> ""  then
							if rs("Preinstall") or rs("Preload") then
								strImageSummary = "<B>" & strImageSummary & "</b>" & "<BR>" & "" & FormatSKUList(rs("Images") & "") & ""
							else
								strImageSummary = "<B>" & strImageSummary & "</b>" 
							end if
						end if
					end if
					
					select case trim(rs("SWSetupCategoryID") & "")
						case "2"
						strSWSetupCat =  "Driver"
						case "3"
						strSWSetupCat =  "Recommended"
						case "4"
						strSWSetupCat =  "Optional"
						case else
						strSWSetupCat =  "&nbsp;"	
					end select
					
					
						if rs("CertRequired") & "" = "0" or rs("CertRequired") = "" then
							strCert = "&nbsp;"
						else
							select case trim(rs("CertificationStatus") & "")
							case "0",""
								strCert  = "Required"
							case "1"
								strCert  = "Submitted"
							case "2"
								strCert  = "Approved"
							case "3"
								strCert  = "Failed"
							case "4"
								strCert  = "Waiver"
							case else
								strCert = "&nbsp;"
							end select
						end if
						
						if request("ProdID") = "" then
							if instr(rs("Images") & "",rs("imageID") & "=")> 0 then
								strCurrentLanguage = mid(rs("Images") & "",instr(rs("Images") & "",rs("imageID") & "=")+ len(rs("imageID")& "") + 1)
								strCurrentLanguage = left(strCurrentLanguage,instr(strCurrentLanguage,")")-1)
								if strCurrentLanguage = "" then
									strCurrentLanguage = strLanguages
								else
									strCurrentLanguage = "<B>" & strCurrentLanguage & "</b>"
								end if
							else
								strCurrentLanguage = strLanguages
							end if
							strImgLang= midrowNR & "<font size=1 face=verdana  class=""text"">" & strCurrentLanguage & "&nbsp;" & "</TD>"
						else
							strImgLang = ""
						end if
					
						strAllNotes  = ""
						strAllNotes = rs("PreinstallNotes") & ""
						if rs("TargetNotes") & "" <> "" then
							if strAllNotes <> "" then
								strAllNotes = "PIN: " & strAllNotes & "<BR>PM: " & rs("TargetNotes")
							else
								strAllNotes = strAllNotes & rs("TargetNotes")
							end if
						end if
						if rs("DropInBox") and trim(rs("DIBHWReq") & "") <> "" then
							if strAllNotes = ""  then
								strAllNotes = "Required Optical Drive: " & rs("DIBHWReq") 
							else
								strAllNotes = "Required Optical Drive: " & rs("DIBHWReq") & "<BR>" & strAllNotes
							end if
						end if
						
						if trim(rs("DeliverableName") & "") = "" then
							strDeliverablename = rs("Name")
						else
							strDeliverablename = rs("DeliverableName")
						end if

						

						strDaveEmail = strDaveEmail & "Checking Language: ReqLang:" & request("ImageLang") & " CurrentLanguage: " & strCurrentLanguage & "<BR>"
						
						if request("ImageLang") = "" or  instr(strCurrentLanguage,request("ImageLang")) > 0 then

							strDaveEmail = strDaveEmail & "Language OK<BR>"
							
							Response.Write "<INPUT type=""hidden"" id=Path" & rs("ID") & " name=Path" & rs("ID") & " value=""" & rs("ImagePath") & """>"

							set rs2 = server.CreateObject("ADODB.Recordset")
							rs2.Open "spGetRootProperties " & clng(rs("RootID")),cn,adOpenForwardOnly
							strMultiLanguage = rs2("MultiLanguage") & ""
							rs2.close
							if strMultiLanguage <> "1" then
								strCDPartNumber = "By Language"
								strKitNumber = "By Language"
							else
								strCDPartNumber = rs("CDPartNumber")
								strKitNumber = rs("CDKitNumber")
							end if

							
							'kz 10/12/04, In Pin Matrix, highlight the deliverables that are in Pin Image but not in official Image
							if ( request("PINTest") = "1" and rs("InImage") and (not rs("InOfficialImage")) ) then
								strRowData = "<tr bgcolor=mistyrose "
							elseif (rs("Preinstall") or rs("Preload")  or rs("SelectiveRestore")) and rs("targeted") and ((not rs("InImage")) or isnull(rs("inimage"))) then
								strRowData = "<tr bgcolor=ivory " 'powderblue
							elseif (rs("Preinstall") or rs("Preload") or rs("ARCD") or rs("DRDVD") or rs("SelectiveRestore")) and (not rs("targeted")) and (rs("InImage") or isnull(rs("inimage"))) then
								strRowData = "<tr bgcolor= ivory " 'mistyrose
							else
								strRowData = "<tr bgcolor=ivory "
							end if
							if trim(rs("imagePath") & "")= "" then
								blnHasPath = 0
							else
								blnHasPath = 1
							end if
							strRowData = strRowData & "ID=""Row" & rs("ID") & """ valign=top LANGUAGE=javascript onmouseover=""return Delrows_onmouseover(" & rs("ID") & ")"" onmouseout=""return Delrows_onmouseout(" & rs("ID") & ")"" onclick=""return DelMenu(" & rs("ID") & "," & rs("RootID") & "," & blnHasPath & ")"" oncontextmenu=""DelMenu(" & rs("ID") & "," & rs("RootID")  &  "," & blnHasPath & ");return false;"">" 
							strShortRowData = strRowData
							strWebRowData = strRowData

							strRowData = strRowData &  "<TD class=""cell"" nowrap><INPUT type=""hidden"" id=txtDelName" & trim(request("ID") & "_" & rs("ID")) & " name=txtDelName value=""" & strDeliverablename & " " & Version & """><font size=1 face=verdana class=""text"">" & strDeliverablename & "</font>" & midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>" & midrow & "<font size=1 face=verdana class=""text"">" & rs("vendorversion") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("irspartnumber") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("CDPartNumber") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("CDKitnumber") & "&nbsp;" & "</font>" & midrow & "<font size=1 face=verdana  class=""text"">" & shortname(rs("Developer")) & "</TD>"
							if blnImageSnapshotsSaved and trim(strlockeddeliverableList & "") <> "" then 													
								strRowData = strRowData & "<td style=""Display:none""><font size=1 ID=""DelPath" & trim(rs("ID")) & """>" & rs("ImagePath") & "</font></td>" & postrow
							else

								if rs("DeveloperNotificationStatus") = "1" then
									strRowData = strRowData &  "<TD nowrap bgcolor=SpringGreen valign=top class=""cell""><font size=1 face=verdana class=""text"">" & "OK" 
								elseif rs("DeveloperNotificationStatus") = "2" then
									strRowData = strRowData  & "<TD nowrap bgcolor=red valign=top class=""cell""><font size=1 face=verdana class=""text"">" & "NO" 
								else
									strRowData = strRowData &  "<TD nowrap bgcolor=Yellow valign=top class=""cell""><font size=1 face=verdana class=""text"">" & "TBD" 
								end if	
								strRowData = strRowData & midrowcenterborder & "<font size=1 face=verdana class=""text"">" & strPreinstall & midrowcenter & "<font size=1 face=verdana class=""text"">" & strPreload & midrowcenter & "<font size=1 face=verdana class=""text"">" & strDIB & midrowcenter & "<font size=1 face=verdana class=""text"">" & strWEB & midrowcenter & "<font size=1 face=verdana class=""text"">" & strSR & midrowcenter & "<font size=1 face=verdana class=""text"">" & strARCD & midrowcenter & "<font size=1 face=verdana class=""text"">" & strDRDVD & midrowcenter & "<font size=1 face=verdana class=""text"">" & strHPPI & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strRoyalty & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strPanel & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strDesktop  & midrowcenter&  "<font size=1 face=verdana class=""text"">" & strMenu  & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strTray  & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strInfoCenter  & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strTile  & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strTaskbarIcon  & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strSWSetupCat & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strCert  & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strAllNotes & "&nbsp;" & "</TD>" & strImgLang & "<TD nowrap valign=top  class=""cell"">" &  "<font size=1 face=verdana class=""text"">" & strImageSummary  &  "</td><td style=""Display:none""><font size=1 ID=""DelPath" & trim(rs("ID")) & """>" & rs("ImagePath") & "</font></td>" & postrow
							end if

							'strRowData = strRowData &  "<TD class=""cell"" nowrap><INPUT type=""hidden"" id=txtDelName" & trim(request("ID") & "_" & rs("ID")) & " name=txtDelName value=""" & strDeliverablename & " " & Version & """><font size=1 face=verdana class=""text"">" & strDeliverablename & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & rs("vendorversion") & "&nbsp;" & "</font>"  & midrownr & "<font size=1 face=verdana class=""text"">" & rs("partnumber") & "&nbsp;" & "</font>"  & midrow & "<font size=1 face=verdana  class=""text"">" & shortname(rs("Developer")) & midrow & "<font size=1 ID=""DistCell" & rs("ID") & "_" &  request("ID") & """ face=verdana  class=""text"">" & strDistribution   & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strSR  & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strARCD   & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strPanel & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strDesktop  & midrowcenter&  "<font size=1 face=verdana class=""text"">" & strMenu  & midrowcenter &  "<font size=1 face=verdana class=""text"">" & strTray  & midrowcenterborder &  "<font size=1 face=verdana class=""text"">" & strCert  & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strAllNotes & "&nbsp;" & "</TD>" & strImgLang & "<TD nowrap valign=top  class=""cell"">" &  "<font size=1 face=verdana class=""text"">" & strImageSummary  &  "</td><td style=""Display:none""><font size=1 ID=""DelPath" & trim(rs("ID")) & """>" & rs("ImagePath") & "</font></td>" & postrow
							strShortRowData = strShortRowData &  "<TD class=""cell"" nowrap><INPUT type=""hidden"" id=txtDelName" & trim(request("ID") & "_" & rs("ID")) & " name=txtDelName value=""" & strDeliverablename & " " & Version & """><font size=1 face=verdana class=""text"">" & strDeliverablename & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & rs("vendorversion") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("irspartnumber") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & strCDpartnumber & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & strCDKitnumber & "&nbsp;" & "</font>" & midrow & "<font size=1 face=verdana  class=""text"">" & shortname(rs("Developer")) &  midrowcenter &  "<font size=1 face=verdana class=""text"">" & strCert  & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strAllNotes & "&nbsp;" & "</TD>" & strImgLang & "<TD style=""display:none"" nowrap valign=top  class=""cell"">" &  "<font size=1 face=verdana class=""text"">" & strImageSummary  &  "</td><td style=""Display:none""><font size=1 ID=""DelPath" & trim(rs("ID")) & """>" & rs("ImagePath") & "</font></td>" & postrow 

							if rs("Web") then
								strOSList = ""
								set rs2 = server.CreateObject("ADODB.Recordset")
								rs2.open "spGetSelectedOS " & rs("ID"),cn,adOpenForwardOnly
								do while not rs2.eof
									if strWebOSList="" or instr("," & strWebOSList & ",","," & trim(rs2("ID")) & ",")>0 then
										strOSList = strOSList & "," & rs2("Shortname")
									end if
									rs2.movenext	
								loop
								rs2.close	
								set rs2 = nothing
								if strOSList = "" then
									strOSList = "&nbsp;"
								else
									strOSList = mid(strOSList,2)
								end if
								strWebRowData  = strWebRowData  &  "<TD class=""cell"" nowrap><INPUT type=""hidden"" id=txtDelName" & trim(request("ID") & "_" & rs("ID")) & " name=txtDelName value=""" & strDeliverablename & " " & Version & """><font size=1 face=verdana class=""text"">" & strDeliverablename & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & Version & "</font>"  & midrow & "<font size=1 face=verdana class=""text"">" & rs("vendorversion") & "&nbsp;" & "</font>"  & midrownr & "<font size=1 face=verdana class=""text"">" & rs("irspartnumber") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("CDpartnumber") & "&nbsp;" & "</font>" & midrownr & "<font size=1 face=verdana class=""text"">" & rs("CDKitnumber") & "&nbsp;" & "</font>" & midrow & "<font size=1 face=verdana  class=""text"">" & shortname(rs("Developer")) &  midrowcenter &  "<font size=1 face=verdana class=""text"">" & strCert  & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strOSList & "&nbsp;" & "</TD><TD nowrap valign=top  class=""cell""><font size=1 face=verdana  class=""text"">" & strAllNotes & "&nbsp;" & "</TD>" & strImgLang & "<TD style=""display:none"" nowrap valign=top  class=""cell"">" &  "<font size=1 face=verdana class=""text"">" & strImageSummary  &  "</td><td style=""Display:none""><font size=1 ID=""DelPath" & trim(rs("ID")) & """>" & rs("ImagePath") & "</font></td>" & postrow
							end if
							

							

' 							if (rs("DropInBox") or rs("ARCD")) then
'							if strOOC = "True" then
'									strOOCDIBRows = strOOCDIBRows & strShortRowData 'strRowData
'								else
'								strDIBRows = strDIBRows & strShortRowData 'strRowData
'								end if
'							end if
'							
'							if rs("Web") then
'								if strOOC = "True" then
'									strOOCWebRows = strOOCWebRows & strShortRowData 'strRowData
'								else
'									strWebRows = strWebRows & strShortRowData 'strRowData
'								end if
'							end if
'							
'							if rs("ARCD") and rs("InImage") then
'								strDRCDRows = strDRCDRows & strRowData
'							end if
'							
'							if rs("InImage") then
'								Response.Write strRowData
'							end if
'
 							if (rs("DropInBox")) then 'or rs("ARCD")) and (not (rs("Preinstall") or rs("Preload") or rs("SelectiveRestore")  )) then
								if strOOC = "True" then
									strOOCDIBRows = strOOCDIBRows & strShortRowData 'strRowData
								else
									strDIBRows = strDIBRows & strShortRowData 'strRowData
								end if
							end if
							if rs("Web") then 'and (not (rs("Preinstall") or rs("Preload") or rs("ARCD") or rs("SelectiveRestore")  )) then
								if strOOC = "True" then
									strOOCWebRows = strOOCWebRows & strWebRowData 'strRowData
								else
									strWebRows = strWebRows & strWebRowData 'strRowData
								end if
							end if
                        
							if (rs("ARCD") or rs("DRDVD")) and rs("InImage") and request("DIBOnly") = "" then
								strDRCDRows = strDRCDRows & strRowData
								Response.Write strRowData
			                    strDaveEmail = strDaveEmail & "DRCDRow<BR>"
							elseif rs("InImage") and request("DIBOnly") = ""  then
								Response.Write strRowData
			                    strDaveEmail = strDaveEmail & "RegularRow<BR>"
			                else
			                    strDaveEmail = strDaveEmail & "RowSkipped: " & rs("ID") & " InImage: " & rs("InImage") & " DIBOnly: >" & request("DIBOnly") & "<<BR>"
							end if
						end if
					end if
				end if
				rs.MoveNext
			loop
		end if

		Response.Write "</TABLE>"
		if DistributionChangeCount <> "0" and request("DIBOnly") = "" then
'			Response.Write "<font color=""red"">*</font><font size=2> Means distribution has changed recently, preinstall team is in the process of updating the deliverable with the new distribution</font>"
		end if

		rs.Close

	if strDIBRows <> "" then
%>		
<BR><BR><table width=660 class=InfoTable border=1 bgcolor=beige cellspacing=0 cellpadding=3>
<%=InfoRow%>
</table><br>
<%if trim(strDevCenter) <> "2" then%>
<table border=1 cellpadding=2 cellspacing=0><tr><td><font size=1 face=verdana>
[1] - Commercial Product Desktop Icon Rules:<br>
<div style="margin-left: 30px">
<%
    rs.open "spGetCycleIconRule " & clng(strproductID),cn
    do while not rs.EOF
        response.write rs("Rule") & ""
        rs.MoveNext
    loop
    rs.Close

%>
</div>
</font>
</TD></TR></TABLE><br>
<%end if%>

<%		
			if request("ProdID") = "" then
				strImgLang= "<TD align=center width=20 valign=bottom><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
			else
				strImgLang = ""
			end if
			if request("DIBOnly") = "" then 
			Response.Write "<BR><font size=2 face=verdana><b>DIB Deliverables:</b></font><BR>"
			end if
			Response.Write "<TABLE width=100% ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
'			Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD rowspan=2 align=left width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD width=26 style=""BORDER-LEFT: gray thin solid;"" rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD rowspan=2 valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD colspan=4 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Image&nbsp;Support</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
'			Response.Write "<tr height=52 bgcolor=beige><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD></tr>"
			Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD align=center width=40 valign=bottom><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
			Response.Write strDIBRows
			Response.Write "</TABLE>"
		end if
		
		
	'else
	'	Response.Write "<BR><font size=2 face=verdana>Unable to find the selected Image</font>"
	'end if

		if strOOCDIBRows <> "" and request("DIBOnly") = "" then
			if request("ProdID") = "" then
				strImgLang= "<TD align=center width=20 valign=bottom><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
			else
				strImgLang = ""
			end if
			Response.Write "<BR><BR><font size=2 face=verdana><b>OOC: DIB Deliverables:</b></font><BR>"
			Response.Write "<TABLE width=100% ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
'			Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD rowspan=2 align=left width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD width=26 style=""BORDER-LEFT: gray thin solid;"" rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD rowspan=2 valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD colspan=4 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Image&nbsp;Support</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
'			Response.Write "<tr height=52 bgcolor=beige><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD></tr>"
			Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD align=center width=40 valign=bottom><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
			Response.Write strOOCDIBRows
			Response.Write "</TABLE>"
		end if



		if strWebRows <> "" and request("DIBOnly") = "" then
			if request("ProdID") = "" then
				strImgLang= "<TD align=center width=20 valign=bottom><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
			else
				strImgLang = ""
			end if
			Response.Write "<BR><BR><font size=2 face=verdana><b>Web Release Deliverables:</b></font><BR>"
			Response.Write "<TABLE width=100% ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
			'Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD rowspan=2 align=left width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD width=26 style=""BORDER-LEFT: gray thin solid;"" rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD rowspan=2 valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD colspan=4 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Image&nbsp;Support</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
			'Response.Write "<tr height=52 bgcolor=beige><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD></tr>"
			Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD align=center width=40 valign=bottom><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Supported OS</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
			Response.Write strWEBRows
			Response.Write "</TABLE>"
		end if
		
		
		if strOOCWebRows <> "" and request("DIBOnly") = "" then
			if request("ProdID") = "" then
				strImgLang= "<TD align=center width=20 valign=bottom><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
			else
				strImgLang = ""
			end if
			Response.Write "<BR><BR><font size=2 face=verdana><b>OOC: Web Release Deliverables:</b></font><BR>"
			Response.Write "<TABLE width=100% ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
			'Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD rowspan=2 align=left width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD width=26 style=""BORDER-LEFT: gray thin solid;"" rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD rowspan=2 valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD colspan=4 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Image&nbsp;Support</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>"
			'Response.Write "<tr height=52 bgcolor=beige><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD></tr>"
			Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1>CD&nbsp;Part&nbsp;Number</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1>Kit&nbsp;Number</b></font></TD><TD align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD align=center width=40 valign=bottom><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Supported OS</b></font></TD><TD align=center width=80 valign=bottom ><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
			Response.Write strOOCWEBRows
			Response.Write "</TABLE>"
		end if
		
		if strDRCDRows <> "" and request("DIBOnly") = "" and request("Type") <> 3 then
			if request("ProdID") = "" then
				strImgLang= "<TD align=center width=20 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Img. Lang.</b></font></TD>"
			else
				strImgLang = ""
			end if
			Response.Write "<BR><BR><font size=2 face=verdana><b>Restore Media Deliverables:</b></font><BR>"
			Response.Write "<TABLE width=100% ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
			Response.Write "<tr bgcolor=beige><TD rowspan=2 align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Vendor&nbsp;Version</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>CD&nbsp;Part&nbsp;Number</b></font></TD><TD rowspan=2 align=center valign=bottom width=50><font face=verdana color=black size=1><b>Kit&nbsp;Number</b></font></TD><TD rowspan=2 align=left width=60 valign=bottom><font face=verdana color=black size=1><b>Developer</b></font></TD><TD width=26 rowspan=2 valign=bottom><IMG height=54 width=23 SRC=""../images/col_developerApproved.GIF""></TD><TD style=""BORDER-LEFT: gray thin solid;"" colspan=7 align=right width=80 valign=bottom><font face=verdana color=black size=1><b>Distribution</b></font></TD><TD rowspan=2 valign=bottom><IMG SRC=""../images/col_HPPI.gif""></TD><TD rowspan=2 valign=bottom><IMG height=45 width=24 width=11 SRC=""../images/col_Royalty.gif""></TD><TD colspan=7 style=""BORDER-LEFT: gray thin solid;"" align=center width=80 valign=bottom><font face=verdana color=black size=1><b>Icons&nbsp;Installed</b></font></TD><TD style=""BORDER-LEFT: gray thin solid;"" align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>SWSetup</b></font></TD><TD align=center width=40 valign=bottom rowspan=2><font face=verdana color=black size=1><b>WHQL</b></font></TD><TD align=center width=200 valign=bottom rowspan=2><font face=verdana color=black size=1><b>Notes</b></font></TD>" & strImgLang & "<TD rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Images&nbsp;In&nbsp;Development</b></font></TD><TD style=""display:none"" rowspan=2 align=left width=120 valign=bottom><font face=verdana color=black size=1><b>Path</b></font></TD></tr>" 
			Response.Write "<tr height=52 bgcolor=beige><TD valign=bottom style=""BORDER-LEFT: gray thin solid;""><IMG height=56 width=14 SRC=""../images/col_Prinstall.gif""></TD><TD valign=bottom><IMG height=48 width=15 SRC=""../images/col_Preload.gif""></TD><TD valign=bottom><IMG height=29 width=12 SRC=""../images/col_Dib.gif""></TD><TD valign=bottom><IMG height=29 width=14 SRC=""../images/col_Web.gif""></TD><TD valign=bottom><IMG height=54 width=23 SRC=""../images/col_selectiverestore.gif""></TD><TD valign=bottom><IMG height=44 width=11 SRC=""../images/col_arcd.gif""></TD><TD valign=bottom><IMG height=44 width=11 SRC=""../images/col_DRDVD.gif""></TD><TD style=""BORDER-LEFT: gray thin solid;"" valign=bottom><IMG height=44 width=22 SRC=""../images/col_controlpanel.gif""></TD><TD valign=bottom><IMG height=50 width=14 SRC=""../images/col_desktop.gif""></TD><TD valign=bottom><IMG height=32 width=22 SRC=""../images/col_menu.gif""></TD><TD valign=bottom><IMG height=45 width=24 SRC=""../images/col_systemtray.gif""></TD><TD valign=bottom><IMG SRC=""../images/col_infocenter.gif""></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Tile</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Task Bar Icon</b></font></TD></tr>" 
			Response.Write strDRCDRows
			Response.Write "</TABLE>"
		end if
        if trim(request("ID")) <> "" then
            rs.open "spListPatchesOnMatrix " & clng(strProductID),cn
			Response.Write "<BR><BR><font size=2 face=verdana><b>Factory Patches:</b></font><BR>"
			Response.Write "<TABLE ID=TableDeliverables  border=1 cellpadding=1 cellspacing=0>"
            intPatchCount = 0
            do while not rs.eof
                if trim(rs("images") & "") = "" or instr("," & replace(rs("images") & ""," ","") & ",","," & trim(request("ID")) & ",") > 0 then
                    intPatchCount = intPatchCount + 1
                    if intpatchcount = 1 then
            			Response.Write "<tr bgcolor=beige><TD align=left valign=bottom><font face=verdana color=black size=1><b>Name</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>Version</b></font></TD><TD align=center valign=bottom width=50><font face=verdana color=black size=1><b>RTM</b></font></TD></tr>" 
                    end if
                    strVersion = rs("Version") & ""
                    if rs("revision") <> "" then
                        strVersion = strversion & "," & rs("revision")
                    end if
                    if rs("pass") <> "" then
                        strVersion = strversion & "," & rs("pass")
                    end if
                    response.write "<tr>"
                    response.write "<td nowrap>" & rs("name") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
                    response.write "<td nowrap>" & strVersion & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
                    if rs("RTM") = 1 then
                        response.write "<td align=""center"">X</td>"
                    else
                        response.write "<td>&nbsp;</td>"
                    end if
                    response.write "</tr>"
                end if
                rs.movenext
            loop
            rs.close
			Response.Write "</TABLE>"
            if intpatchcount = 0 then
                response.write "<b>none</b><BR>"
            end if
            response.write "<BR>"
        end if
	else
		Response.Write "<BR><font size=2 face=verdana>Unable to find the selected Image</font>"
	end if
	
	end if
if trim(strDevCenter) <> "2" then
%>
<br>
<table border=1 cellpadding=2 cellspacing=0><tr><td><font size=1 face=verdana>
[1] - Commercial Product Desktop Icon Rules:<br>
<div style="margin-left: 30px">
<%
    rs.open "spGetCycleIconRule " & clng(strproductID),cn
    do while not rs.EOF
        response.write rs("Rule") & ""
        rs.MoveNext
    loop
    rs.Close

	set rs = nothing
	set cn = nothing
%>
</div>
</font>
</TD></TR></TABLE>
<%end if%>
</BODY>
</HTML>
