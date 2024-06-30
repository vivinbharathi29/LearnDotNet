<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<STYLE>
TD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: ivory;
}
TH
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: cornsilk;
}

.SummaryTH
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: LightSteelBlue;
}

.SummaryTD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: gainsboro;
}
</STYLE>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function DisplayTargetIssues(){
	TargetIssuesRow.style.display = "";	
}
function CompareLines(strTable){
	var i;
		document.all("frmCompare" + strTable).submit();
}

//-->

</SCRIPT>
</HEAD>
<BODY>
<%
	Server.ScriptTimeout = 1200
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

	dim StartDate
	StartDate = now()

	dim cn
	dim cn2
	dim cm
	dim p
	dim rs
	dim rs2

	if request("ProdID") = "" then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
	else
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")

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
		if (rs.EOF and rs.BOF) then
			set rs = nothing
			set cn=nothing
			Response.Redirect "../../NoAccess.asp?Level=0"
		else
			CurrentUserPartner = rs("PartnerID")
		end if 
		rs.Close		
		
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
		strPreinstallTeam = 0
'		rs.Open "spGetProductVersionName " & request("ProdID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product in Excalibur.</font><BR><BR>"
		else

			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					set rs = nothing
					set cn=nothing
					
					Response.Redirect "../../NoAccess.asp?Level=0"
				end if
			end if
			
			'Response.Flush
		
			strPreinstallTeam = rs("PreinstallTeam")
			
		if cint(strPreinstallTeam) = 1 then
    			strConveyor = "houbnbcvr01.auth.hpicorp.net"'"16.101.60.73"
	    		strServerLocation = "Houston"
'		elseif cint(strPreinstallTeam) = 3 then
'				strConveyor = "SGPACCCVR01.auth.hpicorp.net"
'				strServerLocation = "Singapore"
			'elseif cint(strPreinstallTeam) = 4 then
				'strConveyor = "BRAHPQCVR01.auth.hpicorp.net"
				'strServerLocation = "Brazil"
			'elseif cint(strPreinstallTeam) = 5 then
		'		strConveyor = "SHGCDCCVR01.auth.hpicorp.net"
		'		strServerLocation = "China"
			else
				strConveyor = "16.159.144.23"'"tpopsgcvr3.auth.hpicorp.net"'"tpopsgcvr2.auth.hpicorp.net"
				strServerLocation = "Taiwan"
			end if
			strCompareType = rs("Name") & ""
			strproduct = rs("Name") & ""
			Response.Write "<DIV ID=ReportTitle style=""display:""><font size=3 face=verdana><b><center>" &  rs("name") & "</b></font><BR><BR><font size=2 face=verdana>Conveyor Server: " & strServerLocation & "</font></center><BR>"

			Response.Write "<font size=2 face=verdana><u><b>Results</b></u></font><BR><BR></div>"

			rs.Close
			
		end if'product found
		
	
	
       strSQl = "Select distinct i.skunumber " & _
                "from images i with (NOLOCK), imagedefinitions id with (NOLOCK) " & _
                "where id.id = i.imagedefinitionid " & _
                "and lockeddeliverablelist like '% " & request("VersionID") &  ",%' " & _
                "and i.skunumber is not null " & _
                "and productversionid = " & request("ProdID")
                
        rs.open strSQL
        do while not rs.eof
	        response.write rs("SKUNUmber") & "<BR>"
	        rs.movenext 
	    loop
	    rs.close
	    
	    response.write "--------------" & "<BR>"
	
	
	    strSQl = "Select images " & _
                 "from product_deliverable with (NOLOCK) " & _
                 "where productversionid = " & request("ProdID") & " " & _
                 "and deliverableversionid = " & request("VersionID")
	    rs.open strSQl
	    if rs.eof and rs.bof then
	        strImages = ""
        else
	        strImages = trim(rs("Images") & "")
	    end if
		rs.close
		
		if strImages = "" then
		    rs.open
		    
		    rs.close
		end if
		
		
		
		
		set rs = nothing
		set rs2 = nothing
		set cn = nothing

	end if

%>
</b>
<font size=2 face=verdana>
</BODY>
</HTML>