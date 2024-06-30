<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

    function window_onload() {


	
	if (txtID.value=="")
		{
		document.write ("Unable to create files.");
		}
	else
		{
	    window.open("\\\\" & window.location.hostname & "\\Pulsar\\Excalibur\\temp\\Extracts\\" + txtID.value);
		window.opener='X';
		window.open('','_parent','')
		window.close();	
		}
}

//-->
</script>
</head>
<body LANGUAGE="javascript" onload="return window_onload()">

<center><font size="2" face="verdana">
	Generating Files.  Please wait...<br><br>
<img SRC="../images/progressbar.gif" WIDTH="150" HEIGHT="15">
	</font>
</center>

<%

	dim FileCount
	dim strDash
	dim strSKU
	dim strOutBuffer
	dim cm
	dim cn
	dim rs
	dim CurrentUser	

	if request("ProdID") <> "" then 'and Currentuser = "dwhorton" then

		FileCount = 0
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FolderExists("e:\temp\Extracts\" & request("ProdID")) Then
			fso.DeleteFolder("e:\temp\Extracts\" & request("ProdID"))
		end if
		fso.CreateFolder("e:\temp\Extracts\" & request("ProdID"))
	
		set fso = nothing
	

		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
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
		set	rs = server.CreateObject("ADODB.recordset")
	
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
		Set	rs = cm.Execute 
	
		set cm=nothing	
		if (rs.EOF and rs.BOF) then
			set rs = nothing
		    set cn=nothing
        	Response.Redirect "../NoAccess.asp?Level=0"
        else
            CurrentUserPartner = rs("PartnerID")
        end if 
        rs.Close

		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetProductPartner"
			

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("ProdID")
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
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListImagesForProductAll"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'Response.Flush
'		rs.Open "spListImagesForProductAll " & request("ProdID"),cn,adOpenForwardOnly
		do while not rs.EOF
			strDash = trim(rs("Dash") & "")
			strSKU = trim(lcase(rs("SKUNumber") & ""))
			
			if strDash = "" or strSKU = "" then
				strSKU = rs("ID") & ""
			else
				strDash = mid(strDash,2)
				strDash = left(strDash,len(strDash)-1)
				
				strSKU = replace(strSKU,"xx",strDash) 
			end if


			Response.Write "<BR>Generating file " & strSKU  & ".txt<BR>"

				strOutBuffer = ""
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spListDeliverablesInImage"
		

				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = rs("ID")
				cm.Parameters.Append p
	

				rs2.CursorType = adOpenForwardOnly
				rs2.LockType=AdLockReadOnly
				Set rs2 = cm.Execute 
				Set cm=nothing

				
				'rs2.Open "spListDeliverablesInImage " & rs("ID"),cn,adOpenForwardOnly
				'Response.Write "<TABLE border =1>"
				do while not rs2.EOF
					if ( rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore") ) and rs2("InImage") and ( trim(rs2("Images") & "") = "" or instr(", " & rs2("Images") & ",", ", " & rs("ID") & ",")>0  or instr( rs2("Images") , "(" & rs("ID") & "=")>0 )  then
					strOutbuffer = strOutbuffer & rs2("DeliverableName") & chr(9) ' & "***"
					strOutbuffer = strOutbuffer & rs2("Version") & chr(9) ' & "***"
					strOutbuffer = strOutbuffer & rs2("Revision") & chr(9) ' & "***"
					strOutbuffer = strOutbuffer & rs2("Pass") & chr(9) ' & "***"
					strOutbuffer = strOutbuffer & rs2("VendorVersion") & chr(9) '& "***"
					strOutbuffer = strOutbuffer & rs2("PartNumber") & chr(9) ' '& "***"
					strOutbuffer = strOutbuffer & rs2("PreinstallInternalRev") & chr(9) ' '& "***"
					strOutbuffer = strOutbuffer & vbcrlf
					
					
					
					
					
				'		Response.Write "<TR><TD>" & rs2("DeliverableName") & "</TD>"
				'		Response.Write "<TD>" & rs2("Version") & "</TD>"
				'		Response.Write "<TD>" & rs2("Revision") & "&nbsp;</TD>"
				'		Response.Write "<TD>" & rs2("Pass") & "&nbsp;</TD>"
				'		Response.Write "<TD>" & rs2("VendorVersion") & "&nbsp;</TD>"
				'		Response.Write "<TD>" & rs2("PartNumber") & "&nbsp;</TD>"
				'		Response.Write "</TR>"
					end if
					rs2.MoveNext
				loop
				rs2.close		
				'Response.Write "</TABLE>"


				Dim fs,f
				Set fs=Server.CreateObject("Scripting.FileSystemObject")
				Set f = fs.CreateTextFile("e:\temp\Extracts\" & request("ProdID") & "\" & strSKU & ".txt")
			
				f.WriteLine(strOutbuffer)


				f.close
				set f=nothing
				set fs=nothing
				


				'Response.Write strOutBuffer
			FileCount = FileCount + 1
			rs.MoveNext
		loop

		rs.Close


		set rs = nothing
		set rs2 = nothing
		set cn = nothing
	
		if Filecount = 0 then
			Response.Write "No Images found for the selected product."
		else
			Response.Write "Done"
		end if
	
	end if
	



%>
<input type="hidden" id="txtID" name="txtID" value="<%=request("ProdID")%>">
<input type="hidden" id="txtServer" name="txtServer" value="excaliburweb.cca.hp.com">
<!--Application("Excalibur_ServerName")-->
</body>
</html>
