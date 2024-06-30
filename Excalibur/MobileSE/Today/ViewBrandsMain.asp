<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<TITLE>View Selected Brands</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>



<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<!--<font size=3 face=verdana><b>Selected Brands for </b></font>-->


<%


if request("ID") = ""  then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim strLoaded
    dim prodID 
    strID = request("ID")
	'response.write(request("ID"))
	strLoaded = ""

   	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
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
	

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
	end if
	rs.Close

end if

%>
    <table>
                              <thead>
                                <tr style="position: relative; top: expression(document.getElementById('DIV3').scrollTop-2);">
                                   
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Brands
                                    </td>

    <!--                                <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;SCMs
                                    </td>
                                    <td width="302" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Series
                                    </td>
                                     <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Generation
                                    </td>
                                     <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Form Factor
                                    </td>
                                    <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                                        border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Suffix&nbsp;
                                    </td>-->
                                </tr>
                            </thead>

        <tr>


			<!--<td width="160" style="vertical-align: top"><strong><font size="2">Selected&nbsp;Brands:</font></strong><font color="red" size="1">&nbsp;*</font></td>-->
				<td>
                   <!-- <select id="brandsID" name="brandsID" style="width: 200px;">
                         <option selected value=""></option>-->
                   <%  
                      
                        
                   
                       ' rs.open "spListbrands4Product 3885,1",cn,adOpenForwardOnly
                        rs.Open  "spListbrands4Product " & clng(strID) & ",1",cn
                     
                        do while not rs.eof
                     
                    %>     
                              
                        <option value="<%=rs("ID")%>"><%=rs("Name")%></option><br />
                      
                    <%
                        rs.movenext
                        loop
                     
                        rs.close    

                    %>

  


                   <!-- </select>-->

				</td>
        </tr>
    </table>

</BODY>
</HTML>
