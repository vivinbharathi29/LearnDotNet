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
		window.open (Application("DRDVDTextPaht") + txtID.value);
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
	Generating DRDVD Files.  Please wait...<br><br>
<img SRC="../images/progressbar.gif" WIDTH="150" HEIGHT="15">
	</font>
</center>

<%

	dim strDash
	dim strSKU
	dim strOutBuffer
	dim cm
	dim cn
	dim rs
	dim CurrentUser	
	dim CurrentUserEmail		
	dim CurrentDate
	dim strproductName
	dim DelCount
	
	if request("ProdID") <> "" then 'and Currentuser = "dwhorton" then

		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FolderExists("e:\temp\DRDVDText\" & request("ProdID")) Then
			fso.DeleteFolder("e:\temp\DRDVDText\" & request("ProdID"))
		end if
		fso.CreateFolder("e:\temp\DRDVDText\" & request("ProdID"))
	

		set fso = nothing
	

		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Application("PDPIMS_ConnectionString") 
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
		Set	rs = cm.Execute 
	
		set cm=nothing	
		if (rs.EOF and rs.BOF) then
			set rs = nothing
		    set cn=nothing
        	Response.Redirect "../NoAccess.asp?Level=0"
        else
			CurrentUserEmail = rs("Email")
            CurrentUserPartner = rs("PartnerID")
        end if 
        rs.Close

		'get the product name
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
		else
			strproductName = rs("Name") & ""
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
		cm.CommandText = "spListDeliverables4Product"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		CurrentDate = now()

		'Generate the header file
		strOutbuffer = strOutBuffer & "[Cue File]" & vbcrlf
		strOutbuffer = strOutBuffer & "Type=Export" & vbcrlf
		strOutbuffer = strOutBuffer & "Version=1" & vbcrlf
		strOutbuffer = strOutBuffer & "Date=" & CurrentDate & vbcrlf
		strOutbuffer = strOutBuffer & vbcrlf
				
		strOutbuffer = strOutBuffer & "TargetDirectory=" & vbcrlf
		strOutbuffer = strOutBuffer & "ExportName=" & vbcrlf
		strOutbuffer = strOutBuffer & "ProdName=" & strproductName & vbcrlf
		strOutbuffer = strOutBuffer & "FileSystem=1" & vbcrlf
		strOutbuffer = strOutBuffer & "SeparateDelivs=0" & vbcrlf
		strOutbuffer = strOutBuffer & vbcrlf
				
		strOutbuffer = strOutBuffer & "[Email]" & vbcrlf
		strOutbuffer = strOutBuffer & "NotifyFrom=HOUPortPreinDB@hp.com" & vbcrlf
		strOutbuffer = strOutBuffer & "NotifySuccess=" & CurrentUserEmail & vbcrlf
		strOutbuffer = strOutBuffer & "NotifyError=" & CurrentUserEmail & vbcrlf
		strOutbuffer = strOutBuffer & vbcrlf

		strOutbuffer = strOutBuffer & "[ExportList]" & vbcrlf
		
		'Response.Write "<TABLE border =1>"
		DelCount = 1
		do while not rs.EOF
			if  rs("DRDVD") and rs("InImage") then
				strOutbuffer = strOutbuffer & DelCount & "=|||" & rs("DeliverableName") & "|" ' & "***"
				strOutbuffer = strOutbuffer & "|" & "|" & rs("Version") ' & "***"
				strOutbuffer = strOutbuffer & "|" &  rs("Revision") ' & "***"
				strOutbuffer = strOutbuffer & "|" &  rs("Pass") & "|" ' & "***"
				strOutbuffer = strOutbuffer & vbcrlf	
				DelCount = DelCount + 1														
			end if
			rs.MoveNext
		loop
		rs.close		
		'Response.Write "</TABLE>"


		Dim fs,f
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		Set f = fs.CreateTextFile("e:\temp\DRDVDText\" & request("ProdID") & "\" & strproductName & ".txt")
			
		f.WriteLine(strOutbuffer)

		f.close
		set f=nothing
		set fs=nothing
				

		set rs = nothing
		set cn = nothing
	
	end if
	

%>
<input type="hidden" id="txtID" name="txtID" value="<%=request("ProdID")%>">
<input type="hidden" id="txtServer" name="txtServer" value="<%=Application("Excalibur_ServerName")%>">

</body>
</html>
