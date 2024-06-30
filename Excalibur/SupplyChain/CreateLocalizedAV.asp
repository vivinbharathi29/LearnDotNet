<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim cn, cmd

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_ProductBrandID    : m_ProductBrandID = Request("BID")
Dim m_BaseGPG           : m_BaseGPG = Request("BaseGPG")
Dim m_FeatureID         : m_FeatureID = Request("FeatureID")
Dim m_UserName          : m_UserName = Request("UserName")
Dim m_ConfigCode        : m_ConfigCode = Request("ConfigCode")
Dim m_CountryCode       : m_CountryCode = Request("CountryCode")
Dim m_AVParentID        : m_AVParentID = Request("AVParentID")
Dim m_GeoID             : m_GeoID = Request("GeoID")
Dim m_ShareAV           : m_ShareAV = Request("ShareAV")
Dim m_Releases          : m_Releases = Request("Releases")
Dim m_RTPDate           : m_RTPDate = Request("RTPDate")
Dim m_EMDate            : m_EMDate = Request("EMDate")

function ScrubSQL(strWords) 

    dim badChars 
	dim newChars 
	dim i

	'strWords=replace(strWords,"'","''")
		
	badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
	newChars = strWords 
		
	for i = 0 to uBound(badChars) 
		newChars = replace(newChars, badChars(i), "") 
	next 
		
	ScrubSQL = newChars 
	
 end function

'##############################################################################	

set cn = server.CreateObject("ADODB.connection")
set cmd = Server.CreateObject("ADODB.Command")
'                      
cn.ConnectionString = Session("PDPIMS_ConnectionString")
cn.Open
cmd.CommandText = "usp_InsertLocalized_Pulsar"
cmd.CommandType = adCmdStoredProc
set cmd.ActiveConnection = cn

cmd.Parameters.Append cmd.CreateParameter("@PVID", adInteger, adParamInput, , clng(m_ProductVersionID))
cmd.Parameters.Append cmd.CreateParameter("@BID", adInteger, adParamInput, , clng(m_ProductBrandID))
cmd.Parameters.Append cmd.CreateParameter("@FeatureID", adInteger, adParamInput, , clng(ScrubSQL(m_FeatureID)))
cmd.Parameters.Append cmd.CreateParameter("@UserName", adVarChar, adParamInput, 100, ScrubSQL(m_UserName))                   
cmd.Parameters.Append cmd.CreateParameter("@ConfigCode", adVarChar, adParamInput, 5, ScrubSQL(m_ConfigCode))
cmd.Parameters.Append cmd.CreateParameter("@CountryCode", adVarChar, adParamInput, 15, m_CountryCode)
cmd.Parameters.Append cmd.CreateParameter("@AVParentID", adInteger, adParamInput, , clng(m_AVParentID))
cmd.Parameters.Append cmd.CreateParameter("@GeoID", adInteger, adParamInput, , clng(m_GeoID))
cmd.Parameters.Append cmd.CreateParameter("@NewAVParentID", adInteger, adParamOutput)
cmd.Parameters.Append cmd.CreateParameter("@ShareAV", adInteger, adParamInput, clng(m_ShareAV))
cmd.Parameters.Append cmd.CreateParameter("@ReleaseIDs", adVarChar, adParamInput, 250, ScrubSQL(m_Releases))
cmd.Parameters.Append cmd.CreateParameter("@RTPDate", adVarChar, adParamInput, 25, ScrubSQL(m_RTPDate))
cmd.Parameters.Append cmd.CreateParameter("@EMDate", adVarChar, adParamInput, 25, ScrubSQL(m_EMDate))

cmd.Execute

Dim NewAVParentID 
NewAVParentID = cmd.Parameters("@NewAVParentID")

Response.Write NewAVParentID

set cmd=nothing
set cn=nothing


%>

