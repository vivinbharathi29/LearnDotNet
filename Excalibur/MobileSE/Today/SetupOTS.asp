<%@ Language=VBScript %>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=JavaScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	
	var e = document.getElementByID("txtUserID")
	if (e == "")
		document.write ("<BR>Invalid user ID supplied");
	else if (txtFound.value=="1")
		{
		window.open ("\\\\houhpqexcal03.auth.hpicorp.net\\temp\\OTSLinks\\" + e + ".artask");
		window.opener='X';
		window.open('','_parent','')
		window.close();	
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="return window_onload()">

<%
	response.Flush
	dim strOTSID
	dim cn
	dim rs
	dim cm
	dim p
	
	strOTSID = right("0000000" & request("OTSID"),7)

	strConnect = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = strConnect
	cn.CommandTimeout = 60
	cn.Open

	set rs = server.CreateObject("ADODB.Recordset")
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetOTSSummary"
		
	Set p = cm.CreateParameter("@OTSID", 200, &H0001,7)
	p.Value = strOTSID
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
'	rs.open "spGetOTSSummary '" & strOTSID & "'",cn,adOpenForwardOnly
	if rs.eof and rs.bof and request("OTSID") <> "" then
		%><font size=2 face=verdana><b>Unable to find the Observation Number you selected.</b></font><BR>
		  <INPUT type="hidden" id=txtFound name=txtFound value="0">
		<%
	else
		%><font size=2 face=verdana><b>Accessing OTS.  Please wait.</b></font><BR>
		  <INPUT type="hidden" id=txtFound name=txtFound value="1">		
		<%


		if request("USERID") <> "" then
			Dim fs,f
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			Set f = fs.CreateTextFile("e:\temp\OTSLinks\" & request("USERID") & ".artask")
			
			f.WriteLine("You do not appear to have OTS Installed.  To install OTS, please visit the OTS web site:")
			f.WriteLine("")
			f.WriteLine("http://houhpqots04.cce.hp.com/ots/")
			f.WriteLine("")
			f.WriteLine("Or, this URL will open the observation in the OTS web inerface:")
			f.WriteLine("http://houhpqots04.cce.hp.com/arsys/servlet/ViewFormServlet?server=ots01&formalias=OTSXObservationXTracking&view=webView&mode=Query&qual=%27Observation%20ID%27%3D%220165574%22")
			f.WriteLine("")
			f.WriteLine("")
			f.WriteLine("[ShortCut]")
			if request("OTSID") = "" then
    			f.WriteLine("Label = OTS:Main Entry")
    			f.WriteLine("Name = OTS:Main Entry")
    			f.WriteLine("Type = 1")
    			f.WriteLine("Server = houhpqots04.cce.hp.com")
    			f.WriteLine("Join = 0")
            else
			    f.WriteLine("Name = OTS:Observation Tracking")
			    f.WriteLine("Type = 0")
			    f.WriteLine("Server = houhpqots04.cce.hp.com")
			    f.WriteLine("Join=0")
				f.WriteLine("Ticket = " & strOTSID)
			end if
			f.close
			set f=nothing
			set fs=nothing
		end if

	end if

	rs.close	
	cn.Close
	set rs=nothing
	set cn=nothing
	
	set rs = server.CreateObject("ADODB.recordset")
	
	
%>

	<INPUT type="hidden" id=txtUserID name=txtUserID value="<%=Request("USERID")%>">
	<INPUT type="hidden" id=txtServer name=txtServer value="<%=Application("Excalibur_File_Server")%>">

</BODY>
</HTML>
