<!--#include file="../_ScriptLibrary/jsrsServer.inc"--> 
<% jsrsDispatch( "ProfileStrings" ) %>


<script runat="server" language="vbscript">

function ProfileStrings(ID,RequestType,NewName,Value1,Value2,Value3,Value4,Value5,Value6,Value7,Value8,Value9,Value10,Value11,Value12,Value13,Value14,Value15,Value16,Value17,Value18,Value19,Value20,Value21,Value22,Value23,Value24,Value25,Value26,Value27,Value28,Value29,Value30,Value31,Value32,Value33,Value34,Value35,Value36,Value37,Value38,Value39,Value40,Value41,Value42,Value43,Value44,Value45,Value46,Value47,ProfileType,EmployeeID,DefaultSQL,ReportFilters,Value55) 

on error resume next 


dim cn 
dim rs 
dim i
dim strResult

Select Case RequestType
Case 1 'Rename
	strResult = ""
	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spRenameProfile"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.value =ID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProfileName", 200, &H0001, 50)
	p.value = left(NewName,50)
	cm.Parameters.Append p
	
	cm.Execute rowschanged

	if rowschanged = 1 then
		strResult = "1"
		cn.committrans
	else
		cn.rollbacktrans
	end if
	
	set cm = nothing
	set cn = nothing
	ProfileStrings = strResult
Case 2 'Delete
	strResult = ""
	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spDeleteProfile"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.value =ID
	cm.Parameters.Append p

	cm.Execute rowschanged

	if rowschanged = 1 then
		strResult = "1"
		cn.committrans
	else
		cn.rollbacktrans
	end if
	
	set cm = nothing
	set cn = nothing
	ProfileStrings = strResult

Case 3 'Update
	strResult = ""
	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spUpdateProfile"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@ID", 3,  &H0001)
	p.value =ID
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value5", 3, &H0001)
	p.value = Value5
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value6", 200, &H0001, 4)
	p.value = left(Value6,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value7", 3,  &H0001)
	p.value =Value7
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value8", 3,  &H0001)
	p.value =Value8
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value9", 200, &H0001, 20)
	p.value = left(Value9,20)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value10", 200, &H0001, 20)
	p.value = left(Value10,20)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value11", 200, &H0001, 30)
	p.value = left(Value11,30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value12", 200, &H0001, 80)
	p.value = left(Value12,80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value13", 200, &H0001, 255)
	p.value = left(Value13,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value14", 200, &H0001, 255)
	p.value = left(Value14,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value15", 200, &H0001, 2000)
	p.value = left(Value15,2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value16", 3,  &H0001)
	p.value =Value16
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value17", 11,  &H0001)
	p.value =Value17
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value18", 11,  &H0001)
	p.value =Value18
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value19", 11,  &H0001)
	p.value =Value19
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value20", 11,  &H0001)
	p.value =Value20
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value21", 11,  &H0001)
	p.value =Value21
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value22", 3, &H0001)
	p.value = Value22
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value23", 200, &H0001, 4)
	p.value = left(Value23,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value24", 3, &H0001)
	p.value = Value24
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value25", 200, &H0001, 4)
	p.value = left(Value25,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value27", 200, &H0001, 255)
	p.value = left(Value27,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value28", 200, &H0001, 25)
	p.value = left(Value28,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value29", 200, &H0001, 25)
	p.value = left(Value29,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value30", 200, &H0001, 25)
	p.value = left(Value30,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value31", 200, &H0001, 25)
	p.value = left(Value31,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value32", 200, &H0001, 25)
	p.value = left(Value32,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value33", 11,  &H0001)
	p.value =Value33
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value34", 11,  &H0001)
	p.value =Value34
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value35", 11,  &H0001)
	p.value =Value35
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value36", 11,  &H0001)
	p.value =Value36
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value37", 11,  &H0001)
	p.value =Value37
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value38", 11,  &H0001)
	p.value =Value38
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value41", 200, &H0001, 255)
	p.value = left(Value41,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value42", 200, &H0001, 255)
	p.value = left(Value42,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value44", 11,  &H0001)
	p.value =Value44
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value45", 200, &H0001, 2000)
	p.value = left(Value45,2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value46", 200, &H0001, 2000)
	p.value = left(Value46,2000)
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value47", 200, &H0001, 120)
	p.value = left(value47,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value48", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value49", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value50", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value51", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value52", 200, &H0001, 120)
	p.value = ""
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value53", 200, &H0001, 1000)
	p.value = ""
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value54", 200, &H0001, 2000)
	p.value = ""
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value1", 201, &H0001, 2147483647)
	p.value = Value1
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value2", 201, &H0001, 2147483647)
	p.value = Value2
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value3", 201, &H0001, 2147483647)
	p.value = Value3
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value4", 201, &H0001, 2147483647)
	p.value = Value4
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value26", 201, &H0001, 2147483647)
	p.value = Value26
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value39", 201, &H0001, 2147483647)
	p.value = Value39
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value40", 201, &H0001, 2147483647)
	p.value = Value40
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value43", 201, &H0001, 2147483647)
	p.value = Value43
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DefaultSQL", 201, &H0001, 2147483647)
	p.value = DefaultSQL
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ReportFilters", 201, &H0001, 2147483647)
	p.value = ReportFilters
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value55", 201, &H0001, 2147483647)
	p.value = Value55
	cm.Parameters.Append p
	
	cm.Execute rowschanged

	if rowschanged = 1 then
		strResult =  "1"
		cn.committrans
	else
		strResult = rowschanged
		cn.rollbacktrans
	end if
	
	set cm = nothing
	set cn = nothing
	ProfileStrings = strResult
		
Case 4 'Add
	strResult = ""
	set cn = server.createobject("ADODB.Connection") 
	set cm = server.CreateObject("ADODB.Command")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
	cn.open
	
	cn.begintrans
	
	set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn

	cm.CommandText = "spAddProfile"
	cm.CommandType =  &H0004
		
	Set p = cm.CreateParameter("@ProfileName", 200,  &H0001,50)
	p.value = left(NewName,50)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ProfileType", 3,  &H0001)
	p.value =ProfileType
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
	p.value =EmployeeID
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@Value5", 3, &H0001)
	p.value = Value5
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value6", 200, &H0001, 4)
	p.value = left(Value6,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value7", 3,  &H0001)
	p.value =Value7
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value8", 3,  &H0001)
	p.value =Value8
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value9", 200, &H0001, 20)
	p.value = left(Value9,20)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value10", 200, &H0001, 20)
	p.value = left(Value10,20)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value11", 200, &H0001, 30)
	p.value = left(Value11,30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value12", 200, &H0001, 80)
	p.value = left(Value12,80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value13", 200, &H0001, 255)
	p.value = left(Value13,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value14", 200, &H0001, 255)
	p.value = left(Value14,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value15", 200, &H0001, 2000)
	p.value = left(Value15,2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value16", 3,  &H0001)
	p.value =Value16
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value17", 11,  &H0001)
	p.value =Value17
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value18", 11,  &H0001)
	p.value =Value18
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value19", 11,  &H0001)
	p.value =Value19
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value20", 11,  &H0001)
	p.value =Value20
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value21", 11,  &H0001)
	p.value =Value21
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value22", 3, &H0001)
	p.value = Value22
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value23", 200, &H0001, 4)
	p.value = left(Value23,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value24", 3, &H0001)
	p.value = Value24
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value25", 200, &H0001, 4)
	p.value = left(Value25,4)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value27", 200, &H0001, 255)
	p.value = left(Value27,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value28", 200, &H0001, 25)
	p.value = left(Value28,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value29", 200, &H0001, 25)
	p.value = left(Value29,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value30", 200, &H0001, 25)
	p.value = left(Value30,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value31", 200, &H0001, 25)
	p.value = left(Value31,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value32", 200, &H0001, 25)
	p.value = left(Value32,25)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value33", 11,  &H0001)
	p.value =Value33
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value34", 11,  &H0001)
	p.value =Value34
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value35", 11,  &H0001)
	p.value =Value35
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value36", 11,  &H0001)
	p.value =Value36
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value37", 11,  &H0001)
	p.value = Value37
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value38", 11,  &H0001)
	p.value = Value38
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value41", 200, &H0001, 255)
	p.value = left(Value41,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value42", 200, &H0001, 255)
	p.value = left(Value42,255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value44", 11,  &H0001)
	p.value =Value44
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value45", 200, &H0001, 2000)
	p.value = left(Value45,2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value46", 200, &H0001, 2000)
	p.value = left(Value46,2000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value47", 200, &H0001, 120)
	p.value = left(value47,120)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value48", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value49", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value50", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value51", 11,  &H0001)
	p.value =0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value52", 200, &H0001, 120)
	p.value = ""
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value53", 200, &H0001, 1000)
	p.value = ""
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value54", 200, &H0001, 2000)
	p.value = ""
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value1", 201, &H0001, 2147483647)
	p.value = Value1
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value2", 201, &H0001, 2147483647)
	p.value = Value2
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value3", 201, &H0001, 2147483647)
	p.value = Value3
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value4", 201, &H0001, 2147483647)
	p.value = Value4
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value26", 201, &H0001, 2147483647)
	p.value = Value26
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Value39", 201, &H0001, 2147483647)
	p.value = Value39
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value40", 201, &H0001, 2147483647)
	p.value = Value40
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Value43", 201, &H0001, 2147483647)
	p.value = Value43
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DefaultSQL", 201, &H0001, 2147483647)
	p.value = DefaultSQL
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ReportFilters", 201, &H0001, 2147483647)
	p.value = ReportFilters
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@NewID", 3,  &H0002)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@Value55", 201, &H0001, 2147483647)
	p.value = Value55
	cm.Parameters.Append p

	cm.Execute rowschanged

	if rowschanged = 1 then
		strResult = cm("@NewID")
		cn.committrans
	else
		cn.rollbacktrans
	end if
	
	set cm = nothing
	set cn = nothing
	ProfileStrings = strResult
end select
end function 

</script> 

