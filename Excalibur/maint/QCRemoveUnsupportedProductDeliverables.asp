<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (txtAutoClose.value == "1") {
        window.parent.opener = 'X';
        window.parent.open('', '_parent', '')
        window.parent.close();
    }

}

//-->
</SCRIPT>
</HEAD>
<BODY onload="return window_onload()">
<font face=verdana size=2>
<%

Server.ScriptTimeout = 5400

	dim cn
	dim cnQC
	dim strID
	dim rsProds
	dim rs 
	dim rsQC
	dim strSQL
	dim blnFound
	dim ProdID
	dim ProdName
	dim FamilyName
	dim strDelType
	dim strSomeVersions
	dim ProductLine
	dim SEPMEmail
	dim PDMEmail
	dim SETestLeadEmail
	dim TesterEmail
	dim strComponentName
	dim CountTotal
	dim CountComponentsUpdated
	dim CountProductLinksUpdated
	dim strDeveloperEmail
	dim strProductDeveloperEmail
	dim strDevManagerEmail
	dim strProductDevManagerEmail
	dim strFinalSQL
	dim strGeneric
	

	CountTotal = 0
	
	set cnQC = server.CreateObject("ADODB.Connection")

    if request("ITG")="1" then
    	'ITG Server
	    cnQC.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=SIMPACT;Server=gvs00800,2048;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=PSG_SIMPACT;PASSWORD=@PSG!Pwd2%simpact;" 'Application("QC_ConnectionString")
    else
        'Prod Server	
        cnQC.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=simpact;Server=gvv41555.houston.hp.com,2048;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=PSG_SIMPACT;PASSWORD=@PSG!Pwd2%simpact;" 'Application("QC_ConnectionString")
    end if
    
	cnQC.IsolationLevel=256
	cnQC.Open


	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")
	rs.open "spListSIProductComponents2Remove",cn
    CountTotal  = 0
    do while not rs.eof
        CountTotal = CountTotal  + 1
         
        rs2.open "spGetDeliverableDeveloper " & clng(rs("DelID")),cn
        if rs2.eof and rs2.bof then
            strProductDeveloperEmail = ""
            strProductDevManagerEmail = ""
        else
            strProductDeveloperEmail = rs2("DeveloperEmail") & ""
            strProductDevManagerEmail = rs2("DevManagerEmail") & ""
        end if
        rs2.close

        strSQL = "UpdatePlatformComponent " & rs("ProdID") & "," & rs("DelID") & ",0,'" & strProductDeveloperEmail & "','" & strProductDevManagerEmail & "',0 "
	    if request("autoclose") <> "1" then
    	    response.write strSQL & "<BR>"
	    end if	
        cnQC.Execute strSQL

		Response.Flush
		rs.MoveNext
	loop
	rs.Close

	set rs = nothing
    cn.close
    set cn = nothing	    
    cnQC.close
    set cnqc = nothing

%>
<BR>Done.

<%
	response.write "<BR><BR>Total Issues Corrected: " & CountTotal & "<BR>" 
%>
</font>
     <input id="txtAutoClose" type="hidden" value="<%=trim(request("autoclose"))%>">
</BODY>
</HTML>
