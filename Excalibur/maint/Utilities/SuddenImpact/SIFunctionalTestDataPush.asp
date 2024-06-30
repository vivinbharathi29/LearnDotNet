<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file="../../../includes/emailwrapper.asp" -->
<!-- #include file = "../../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
<style>
   td
   {
    font-family:Verdana;
    font-size:8pt;
    font-weight:normal;
   }
      body
   {
    font-family:Verdana;
    font-size:8pt;
    font-weight:normal;
   }
</style>
</HEAD>
<BODY>
<font face=verdana size=2><b>
Functional Test products pushed to Sudden Impact</b></font><br><br>
Note: Updates Product Name, PM, and Tester only (No Deliverables)
<br><br>
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
    dim blnPushDeliverables

	CountTotal = 0
	CountComponentsUpdated = 0
	CountProductLinksUpdated = 0
	
    if request("PushDeliverables") = "1" then
        blnPushDeliverables = true
    else
        blnPushDeliverables = false
    end if
	
	set cnQC = server.CreateObject("ADODB.Connection")

    if request("ITG")="1" then
    	'ITG Server
	    cnQC.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=SIMPACT;Server=gvs12016.auth.hpicorp.net,2048;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=PSG_SIMPACT;PASSWORD=@PSG!Pwd2%simpact;" 'Application("QC_ConnectionString")
    else
        'Prod Server	
        cnQC.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=simpact;Server=gvv11651.auth.hpicorp.net,2048;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=PSG_SIMPACT;PASSWORD=@PSG!Pwd2%simpact;" 'Application("QC_ConnectionString")
    end if
    
	cnQC.IsolationLevel=256
	cnQC.Open


	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	set rsProds = server.CreateObject("ADODB.recordset")
	rsProds.open "spListOTSVirtualProductsInExcalibur",cn
    response.write "<table bgcolor=ivory border=1 cellpadding=2 cellspacing=0 bordercolor=gainsboro>"
    response.write "<tr bgcolor=beige>"
    response.write "<td><b>ID</b></td>"
    response.write "<td><b>Product</b></td>"
    response.write "<td><b>PM</b></td>"
    response.write "<td><b>Tester</b></td>"
    response.write "</tr>"
	do while not rsProds.eof
		ProdID = rsProds("ID")
		SEPMEmail = left(rsProds("PMEmail") ,64)
		SETestLeadEmail = left(rsProds("TesterEmail") ,64)
		
		'Add/Update Product in SI
   		strFinalSQL = "UpdatePlatform " & rsProds("ID") & ",0,'ZM','" & rsprods("PrimaryProduct") & "','" & rsprods("ProductVersion") & "',null,'" & SEPMEmail & "','" & SETestLeadEmail & "',null,null,1,null,null"
		'Response.Write "<BR>" & strFinalSQL
   		cnQC.Execute strFinalSQL,rowsupdated

		'Add/Update to link the ODM
		strFinalSQL = "UpdateODMPlatform " & rsprods("PartnerID") & ",'" & rsProds("Partner") & "'," & rsProds("ID") & ",0,1"
		'Response.Write "<BR>" & strFinalSQL	& "<BR>"
		cnQC.Execute strFinalSQL,rowsupdated	
	    
	    response.write "<tr>"
	    response.write "<td>" & rsProds("ID") & "</td>"
	    response.write "<td>" & rsprods("ProductVersion") & "</td>"
	    response.write "<td>" & SEPMEmail & "</td>"
	    response.write "<td>" & SETestLeadEmail & "</td>"
	    response.write "</tr>"

		'Add/Update Components
		
        if blnPushDeliverables then 'Components
		    strSQl = "Select r.id as RootID,v.id, r.typeid, r.name, v.version, v.revision, v.pass, v.vendorversion, e1.email as DeveloperEmail, e2.email as DevManagerEmail, c.name as Category, c.id as CategoryID, 0 as Generic, v.partnumber, v.modelnumber, vd.id as VendorID, vd.name as Vendor " & _
                    "from deliverableroot r with (NOLOCK), deliverableversion v with (NOLOCK), DeliverableCategory c with (NOLOCK), employee e1 with (NOLOCK), employee e2 with (NOLOCK), vendor vd with (NOLOCK) " & _ 
                    "where v.deliverablerootid = r.id " & _
                    "and r.typeid in (1,2,3,4) " & _ 
                    "and C.id = r.categoryid " & _ 
                    "and e2.id = r.devmanagerid " & _  
                    "and e1.id = v.developerid " & _  
                    "and vd.id = v.vendorid " & _
                    "and OTSFVTOrganizationID = " & clng(rsProds("ID")) & ";"

		    set rs = server.CreateObject("ADODB.recordset")

		    rs.Open strSQL,cn,adOpenForwardOnly
		    errorcount=0
		
		    do while not rs.EOF
			    Response.Write "Sending " & rs("ID") & "<BR>"
			    Response.Flush
			    CountTotal = CountTotal+1

			    strVendorVersion = ucase(rs("VendorVersion"))
			    if trim(strVendorVersion) = "" then
				    strVendorVersion = "XX"
			    end if
		
			    if ucase(left(rs("Category"),6)) = "DRIVER" then
				    NewQCCategoryName = "Driver" 
			    elseif ucase(left(rs("Category") ,11)) = "APPLICATION" then
				    NewQCCategoryName = "Application" 
			    else
				    NewQCCategoryName = rs("Category") 
			    end if
		
			    if rs("TypeID") = 1 and rs("Generic") = 0 then
				    strComponentName = rs("Vendor") & " " & rs("Name")
			    else
				    strComponentName =  rs("Name")                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
			    end if

			    if rs("TypeID") = 1 and rs("Generic") = 0 then
				    strLocalization = rs("Version")
				    if rs("Revision") <> "" then
					    strLocalization = strLocalization & "," & rs("Revision")
				    end if
				    if rs("Pass") <> "" then
					    strLocalization = strLocalization & "," & rs("Pass")
				    end if
				    strLocalization = rs("ModelNumber") & " [" & strLocalization & "]"
			    else
				    strLocalization =  "XX"
			    end if
					
			    if rs("TypeID") = 1 and rs("Generic") = 0 then
				    strCPQVersion =  rs("PartNumber")
			    else
				    strCPQVersion = rs("Version")
				    if rs("Revision") <> "" then
					    strCPQVersion = strCPQVersion & "," & rs("Revision")
				    end if
				    if rs("Pass") <> "" then
					    strCPQVersion = strCPQVersion & "," & rs("Pass")
				    end if
			    end if
											
			    if (rs("TypeID") = 1 and rs("Generic") =0) and not isnull(rs("PartNumber")) then
				    strPartNumber ="'" &  rs("PartNumber") & "'"
			    else
				    strPartNumber = "null"
			    end if


                'Lookup Tester
                if rs("RootID") <> 0 then
        	        set rs2 = server.CreateObject("ADODB.recordset")
	    	        rs2.open "spGetDeliverableTester " & rs("RootID"),cn
    		        if rs2.eof and rs2.bof then
    		            strComponentTesterEmail = "null"
    		        else
    		            strComponentTesterEmail = "'" & rs2("Email") & "'"
    		        end if
		            rs2.close
		            set rs2=nothing
		        else
		            strComponentTesterEmail = "null"
		        end if
			
                if trim(rs("categoryID")) = "2" or trim(rs("categoryID")) = "3" or trim(rs("categoryID")) = "15" or trim(rs("categoryID")) = "71" or trim(rs("categoryID")) = "205" or trim(rs("categoryID")) = "36" or trim(rs("categoryID")) = "131" then
			        strGeneric = "1"
                else
			        strGeneric = rs("Generic")
			    end if
					
			    if trim(strGeneric) = "0" then
				    strDeveloperEmail = "'" & left(rs("DeveloperEmail"),50) & "'"
				    strDevManagerEmail = "'" & left(rs("DevManagerEmail"),50) & "'"
				    strProductDeveloperEmail = "null"
				    strProductDevManagerEmail= "null"
			    else
				    strDeveloperEmail = "null"
				    strDevManagerEmail = "null"
				    strProductDeveloperEmail  = "'" & left(rs("DeveloperEmail"),50) & "'"
				    strProductDevManagerEmail = "'" & left(rs("DevManagerEmail"),50) & "'"
			    end if

			    if instr(rs("Name"),"'")=0 then
                    'strFinalSQL = "UpdateComponent " & rs("ID") & ",0," & rs("TypeID") & "," & rs("categoryID") & ",'" & rs("Category")  & "'," &  strPartNumber & ", '" & strComponentName & "',null," & rs("vendorID") & ",'" & rs("Vendor") & "','" & strCPQVersion & "','" & strVendorVersion & "','" & strLocalization & "'," & strGeneric & "," & strDevManagerEmail & "," & strDeveloperEmail & ",null," & strDeveloperEmail & ",1,null"
		            strFinalSQL = "UpdateComponent " & rs("ID") & ",0," & rs("TypeID") & "," & rs("categoryID") & ",'" & rs("Category")  & "'," &  strPartNumber & ", '" & strComponentName & "',null," & rs("vendorID") & ",'" & rs("Vendor") & "','" & strCPQVersion & "','" & strVendorVersion & "','" & strLocalization & "'," & strGeneric & "," & strDevManagerEmail & "," & strDeveloperEmail & ",null," & strComponentTesterEmail & ",1,null"
				    response.Write strFinalSQL & "<BR>"
				    cnQC.Execute  strFinalSQL

                    strFinalSQL = "UpdatePlatformComponent " & ProdID & "," & rs("ID") & ",0,'" & strProductDeveloperEmail & "','" & strProductDevManagerEmail & "',1 "
				    response.write strFinalSQL & "<BR>"
				    cnQC.Execute strFinalSQL
			    end if
			
			    Response.Write "<BR>"
			    Response.Flush
			    rs.MoveNext
		    loop
		    rs.Close
		    set rs = nothing
	    
	        end if 'Components
	        rsProds.movenext
	    loop
	    rsProds.close
	    set rsProds = nothing
	
        response.write "</table>"
	cn.Close
	cnQC.Close
	set cn = nothing
	set cnQC = nothing



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


%>

</BODY>
</HTML>
