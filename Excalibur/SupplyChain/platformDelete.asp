<%@ Language=VBScript %>

<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/EmailQueue.asp" -->

<%
Dim rs, dw, cn, cmd

Dim m_ID	: m_ID = Request("ID")
Dim m_ProductVersionID : m_ProductVersionID = Request("PVID")
Dim m_PlatformID : m_PlatformID = Request("PlatformID")
Dim m_ProductBrandID : m_ProductBrandID = Request("PBID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim m_Error
Dim firstPfID : firstPfID =0
Dim firstPfName : firstPfName=""
Dim followMKT : followMKT = Request("followMKT")
Dim logoBadge : logoBadge=""
Dim pfName : pfName=""
Dim mktName: mktName=""
Dim curUserMail: curUserMail=""
Dim pvName: pvName=""
'##############################################################################	
'
' Create Security Object to get User Info
'
	m_Error = ""
	m_EditModeOn = False
		
    Dim Security
	Set Security = New ExcaliburSecurity
	
'	m_IsSysAdmin = Security.IsSysAdmin()
'	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
'	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)

	m_IsSysAdmin = Security.IsPulsarSystemAdmin()
	m_IsProgramCoordinator = Security.IsProgramCoordinatorPermissions()
	m_IsConfigurationManager = Security.IsConfigurationManagerPermissions()

	m_UserFullName = Security.CurrentUserFullName()
    curUserMail = Security.CurrentUserEmail()
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If
	
    ''m_EditModeOn = True 'remove this line when this ui is ready to have permission 

	If Not m_EditModeOn Then
		m_Error = "Insufficient User Privileges.  Access Denied"
    Else

        On Error Resume Next
        Err.Clear

        Set rs = Server.CreateObject("ADODB.RecordSet")
        Set cn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        Set dw = New DataWrapper
        Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
        
        ' 07/26/2016 - ADao - Change IRS_Platform_Alias, IRS_Platform, IRS_Alias synonyms to use actual tables 
        rs.Open "SELECT  count(1) as Total " &_
                "FROM	Feature FE INNER JOIN " &_
		        "Alias A ON FE.AliasID = A.AliasID INNER JOIN " &_
		        "Platform_Alias PA ON PA.AliasID = A.AliasID INNER JOIN " &_
		        "AvDetail AV ON AV.FeatureID = FE.FeatureID INNER JOIN " &_
		        "AvDetail_ProductBrand AVPB ON AVPB.AvDetailID = AV.AvDetailID " &_
                "WHERE	PA.PlatformID = " & m_PlatformID & " AND ltrim(rtrim(A.Name)) <> '' AND AVPB.ProductBrandID = " & m_ProductBrandID & " AND Status in ('A')", cn
        if rs("Total") > 0 then
            rs.close()
            m_Error = "Active Base Unit AVs are already set up in the SCM for this Base Unit Group and Brand.  Must set all Base Units to Obsolete before removing the Product's Base Unit Group"                         
        else
            rs.close()
            'detect if the platform is used in product drop
             rs.Open "SELECT  count(1) as Total " &_
                "FROM	productdrop_platform  " &_
		         "WHERE	productdrop_platform.PlatformID = " & m_PlatformID, cn
            if rs("Total") > 0 then
                rs.close()
                m_Error = "The Base Unit Group is already used in product drop(s).  Must remove it from product drop(s) before removing the Product's Base Unit Group"                         
            else
                
                if (followMKT =1) then 'check if it is first base unit group to be deleted.
                    rs.close()
                    
                    rs.open " ;WITH scmNo as(SELECT ISNULL(SCMNumber,0) as SCMNumber FROM product_brand where productversionid = " & m_ProductVersionID &  " AND ID = " & m_ProductBrandID & ")"  &_
                            " SELECT PlatformID, PHWebFamilyName FROM ( " &_ 
							       " SELECT TOP 1 pf.PlatformID, pf.PHWebFamilyName FROM Productversion_Platform pp " &_
							       " JOIN Platform pf on pp.PlatformID = pf.PlatformID " &_
							       " JOIN product_brand pb on pp.productbrandid = pb.id " &_
							       " JOIN scmNo s on ISNULL(pb.SCMNumber,0) = s.SCMNumber " &_
                                   " WHERE pp.ProductVersionID = " & m_ProductVersionID &  "ORDER BY pf.PlatformID ) p1 WHERE p1.PlatformID = "& m_PlatformID, cn, adOpenForwardOnly
                        
                    if not (rs.EOF and rs.BOF) then
	                    firstPfID =  rs("PlatformID")
                        firstPfName =  rs("PHWebFamilyName")
                    end if
                    rs.Close
                end if
                
                Set cmd = dw.CreateCommAndSP(cn, "usp_ProductVersion_RemovePlatform")
                cmd.CommandTimeout = 0
                dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, 0
                dw.CreateParameter cmd, "@p_PlatformID", adInteger, adParamInput, 8, 0
                dw.CreateParameter cmd, "@p_ID", adInteger, adParamInput, 8, m_ID
                cmd.Execute  
            end if    
        end if  
          
       If Err.Number <> 0 or cn.Errors.count > 0 Then        
          m_Error = Err.Description      
       else 
            if (followMKT =1) and (INT(m_PlatformID) = firstPfID) then
                rs.open " ;WITH scmNo as(SELECT ISNULL(SCMNumber,0) as SCMNumber FROM product_brand where productversionid = " & m_ProductVersionID &  " AND ID = " & m_ProductBrandID & ")"  &_
                        ", curFirstPf as ( " &_
	                        " SELECT TOP 1 pv.DotsName, pf.PlatformID, pf.PHWebFamilyName, pf.MktNameMaster FROM Productversion_Platform pp "&_
	                        " JOIN Platform pf on pp.PlatformID = pf.PlatformID "&_
                            " JOIN ProductVersion pv on pp.ProductVersionID = pv.ID " &_
                            " JOIN Product_Brand pb on pp.ProductBrandID = pb.ID " &_
                            " JOIN scmNo s on ISNULL(pb.SCMNumber,0) = s.SCMNumber " &_
                            " WHERE pp.ProductVersionID =" & m_ProductVersionID & " order by pf.PlatformID ) "&_
                " SELECT cp.DotsName, cp.PHWebFamilyName, cp.MktNameMaster, pb.LogoBadge FROM ProductVersion_Platform pp "&_
                " JOIN Product_Brand pb on pp.ProductBrandID = pb.Id "&_
                " JOIN curFirstPf cp on pp.PlatformID = cp.PlatformID WHERE ISNULL(pp.SeriesID,0) = 0 "&_
                " UNION " &_
                " SELECT cp.DotsName, cp.PHWebFamilyName, cp.MktNameMaster, s.LogoBadge FROM ProductVersion_Platform pp"&_
                " JOIN Product_Brand pb on pp.ProductBrandID = pb.Id "&_
                " JOIN curFirstPf cp on pp.PlatformID = cp.PlatformID "&_
                " JOIN Series s on pp.SeriesID = s.ID "
                if not (rs.EOF and rs.BOF) then
                    pvName = rs("DotsName")
                    pfName = rs("PHWebFamilyName")
                    mktName = rs("MktNameMaster")
    				logoBadge =  rs("LogoBadge")

                    Set	oMessage = New EmailQueue 		
			        oMessage.From = "pulsar.support@hp.com"
			        oMessage.To = "MobileExcalNotification-ProductNames@hp.com;" & curUserMail 
			        dim strOutput
			        strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			        strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Marketing Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & mktName & "</font></TD></TR>"
                    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo Badge C Cover</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & logoBadge & "</font></TD></TR>"
                    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Old PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & firstPfName & "</font></TD></TR>"
                    strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>New PHweb</b></font></TD></TR>"
			        strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & pfName & "</font></TD></TR>"
			        strOutput = strOutput & "</TABLE><BR><BR>"
                    oMessage.Subject = pvName & " series definitions updated in Pulsar" 
				    oMessage.HTMLBody = "<font face=Arial size=3 color=black><b>Series Updated</b></font>" & strOutput
    
                    oMessage.SendWithOutCopy
			        Set oMessage = Nothing 
                end if
			    rs.Close
            end if
       End If  
	End If

	Set Security = Nothing

'##############################################################################	



Dim AppRoot
AppRoot = Session("ApplicationRoot")
%>
<HTML>
<HEAD>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/SupplyChain/style.css" />
    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        function Body_Close() {

            var txtError = document.getElementById("txtError").value;
            if(txtError != "")
                alert(txtError);

            setTimeout(function () {
                //window.returnValue = 1;
                window.parent.Completed();
            }, 1000);          
           
        }  
    </script>
</HEAD>
<BODY onload="Body_Close();" style="height:50px; width:390px">
 <INPUT type="hidden" id="txtError" name="txtError" value="<%=m_Error%>">
</BODY>
</HTML>
