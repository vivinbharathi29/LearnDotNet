 <%@Language="VBScript"%>
<%'*************************************************************************************
'* FileName		: schedule_datepicker.asp
'* Description	: Custom date picker so Schedule pages work with new modalDialog.
'* Creator		: Harris, Valerie
'* Created		: 09/29/2016 - PBI 26987
'************************************************************************************* %>
<HTML>
<HEAD>
<TITLE>
	Select A Date
</TITLE>
<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
<STYLE TYPE="TEXT/CSS">
<!--
	.MainTbl {border-left: 1px black solid; border-top: 1px black solid; border-right: 1px black solid; border-bottom: 1px black solid;}
	.TD {Font-family:tahoma, Arial, Verdana; font-weight:400; font-size: 8pt; color:#000000;}
	.INPUTComb {Font-family:tahoma, Arial, Verdana; font-weight:400; font-size: 7pt; color:#000000;}
	.CalDay {Font-family:tahoma, Arial, Verdana; font-weight:600; font-size: 8pt; color:#0000a0; text-align: center; width:25px; height:20px; background-color: #D5D1C8; border-left: 1px black offset; border-top: 1px black offset; border-right: 1px black inset; border-bottom: 1px black inset; cursor:hand}
	.OffCalDay {Font-family:tahoma, Arial, Verdana; font-weight:600; font-size: 7pt; color:buttonshadow; text-align: center; width:25px; height:20px; background-color: menu; border-left: 1px black offset; border-top: 1px blackoffset; border-right: 1px black inset; border-bottom: 1px black inset; cursor:hand}
	.Days {Font-family:Verdana,tahoma,Arial; font-weight:bold; font-size: 8pt; color:black; text-align: center;border-left: 2px black solid; border-top: 2px black solid; border-right: 2px black solid; border-bottom: 2px black solid;}
	.NoDay {Font-family:tahoma, Arial, Verdana; font-weight:600; font-size: 8pt; color:#0000a0; text-align: center; width:25px; height:20px; }

	.INPUTBUTTON {Font-family:Tahoma, Verdana, Arial; font-weight:400; font-size: 8pt; color:#0000a0; background-color: #D5D1C8; border-left: 1px #ffffff solid; border-top: 1px #ffffff solid; border-right: 1px #000000 solid; border-bottom: 1px #000000 solid; cursor:hand}
	A:link {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
	A:active {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
	A:visited {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #000000; text-decoration: none;}
	A:hover {Font-family: Tahoma, Arial, Verdana; Font-size: 12pt; Font-weight: 600; color: #00cc00; text-decoration: none;}
-->
</STYLE>

<SCRIPT Language=vbscript>

Option Explicit
	Dim American_lcid,Australian_lcid
	Dim curLocale,newLocale,strNewLocale
	Dim blnReload,strReDraw,objMonth
	Dim MyArray(2,117)
	Dim StrLocaleInfo,x
	'This part is just fo fun....
MyArray(0,0)="AFRIKAANS"
MyArray(1,0)="1078"
MyArray(0,1)="ALBANIAN"
MyArray(1,1)="1052"
MyArray(0,2)="ARABIC - U.A.E."
MyArray(1,2)="14337"
MyArray(0,3)="ARABIC - BAHRAIN"
MyArray(1,3)="15361"
MyArray(0,4)="ARABIC - ALGERIA"
MyArray(1,4)="5121"
MyArray(0,5)="ARABIC - EGYPT"
MyArray(1,5)="3073"
MyArray(0,6)="ARABIC - IRAQ"
MyArray(1,6)="2049"
MyArray(0,7)="ARABIC - JORDAN"
MyArray(1,7)="11265"
MyArray(0,8)="ARABIC - KUWAIT"
MyArray(1,8)="13313"
MyArray(0,9)="ARABIC - LEBANON"
MyArray(1,9)="12289"
MyArray(0,10)="ARABIC - LIBYA"
MyArray(1,10)="4097"
MyArray(0,11)="ARABIC - MOROCCO"
MyArray(1,11)="6145"
MyArray(0,12)="ARABIC - OMAN"
MyArray(1,12)="8193"
MyArray(0,13)="ARABIC - QATAR"
MyArray(1,13)="16385"
MyArray(0,14)="ARABIC - SAUDIA ARABIA"
MyArray(1,14)="1025"
MyArray(0,15)="ARABIC - SYRIA"
MyArray(1,15)="10241"
MyArray(0,16)="ARABIC - TUNISIA"
MyArray(1,16)="7169"
MyArray(0,17)="ARABIC - YEMEN"
MyArray(1,17)="9217"
MyArray(0,18)="BASQUE"
MyArray(1,18)="1069"
MyArray(0,19)="BELARUSIAN"
MyArray(1,19)="1059"
MyArray(0,20)="BULGARIAN"
MyArray(1,20)="1026"
MyArray(0,21)="CATALAN"
MyArray(1,21)="1027"
MyArray(0,22)="CHINESE"
MyArray(1,22)="4"
MyArray(0,23)="CHINESE - PRC"
MyArray(1,23)="2052"
MyArray(0,24)="CHINESE - HONG KONG"
MyArray(1,24)="3076"
MyArray(0,25)="CHINESE - SINGAPORE"
MyArray(1,25)="4100"
MyArray(0,26)="CHINESE - TAIWAN"
MyArray(1,26)="1028"
MyArray(0,27)="CROATIAN"
MyArray(1,27)="1050"
MyArray(0,28)="CZECH"
MyArray(1,28)="1029"
MyArray(0,29)="DANISH"
MyArray(1,29)="1030"
MyArray(0,30)="DUTCH"
MyArray(1,30)="1043"
MyArray(0,31)="DUTCH - BELGIUM"
MyArray(1,31)="2067"
MyArray(0,32)="ENGLISH"
MyArray(1,32)="9"
MyArray(0,33)="ENGLISH - AUSTRALIA"
MyArray(1,33)="3081"
MyArray(0,34)="ENGLISH - BELIZE"
MyArray(1,34)="10249"
MyArray(0,35)="ENGLISH - CANADA"
MyArray(1,35)="4105"
MyArray(0,36)="ENGLISH - IRELAND"
MyArray(1,36)="6153"
MyArray(0,37)="ENGLISH - JAMAICA"
MyArray(1,37)="8201"
MyArray(0,38)="ENGLISH - NEW ZEALAND"
MyArray(1,38)="5129"
MyArray(0,39)="ENGLISH - SOUTH AFRICA"
MyArray(1,39)="7177"
MyArray(0,40)="ENGLISH - TRINIDAD"
MyArray(1,40)="11273"
MyArray(0,41)="ENGLISH - UNITED KINGDOM"
MyArray(1,41)="2057"
MyArray(0,42)="ENGLISH - UNITED STATES"
MyArray(1,42)="1033"
MyArray(0,43)="ESTONIAN"
MyArray(1,43)="1061"
MyArray(0,44)="FARSI"
MyArray(1,44)="1065"
MyArray(0,45)="FINNISH"
MyArray(1,45)="1035"
MyArray(0,46)="FAEROESE"
MyArray(1,46)="1080"
MyArray(0,47)="FRENCH - STANDARD"
MyArray(1,47)="1036"
MyArray(0,48)="FRENCH - BELGIUM"
MyArray(1,48)="2060"
MyArray(0,49)="FRENCH - CANADA"
MyArray(1,49)="3084"
MyArray(0,50)="FRENCH - LUXEMBOURG"
MyArray(1,50)="5132"
MyArray(0,51)="FRENCH - SWITZERLAND"
MyArray(1,51)="4108"
MyArray(0,52)="GAELIC - SCOTLAND"
MyArray(1,52)="1084"
MyArray(0,53)="GERMAN - STANDARD"
MyArray(1,53)="1031"
MyArray(0,54)="GERMAN - AUSTRIAN"
MyArray(1,54)="3079"
MyArray(0,55)="GERMAN - LICHTENSTEIN"
MyArray(1,55)="5127"
MyArray(0,56)="GERMAN - LUXEMBOURG"
MyArray(1,56)="4103"
MyArray(0,57)="GERMAN - SWITZERLAND"
MyArray(1,57)="2055"
MyArray(0,58)="GREEK"
MyArray(1,58)="1032"
MyArray(0,59)="HEBREW"
MyArray(1,59)="1037"
MyArray(0,60)="HINDI"
MyArray(1,60)="1081"
MyArray(0,61)="HUNGARIAN"
MyArray(1,61)="1038"
MyArray(0,62)="ICELANDIC"
MyArray(1,62)="1039"
MyArray(0,63)="INDONESIAN"
MyArray(1,63)="1057"
MyArray(0,64)="ITALIAN - STANDARD"
MyArray(1,64)="1040"
MyArray(0,65)="ITALIAN - SWITZERLAND"
MyArray(1,65)="2064"
MyArray(0,66)="JAPANESE"
MyArray(1,66)="1041"
MyArray(0,67)="KOREAN"
MyArray(1,67)="1042"
MyArray(0,68)="LATVIAN"
MyArray(1,68)="1062"
MyArray(0,69)="LITHUANIAN"
MyArray(1,69)="1063"
MyArray(0,70)="MACEDONIAN"
MyArray(1,70)="1071"
MyArray(0,71)="MALAY - MALAYSIA"
MyArray(1,71)="1086"
MyArray(0,72)="MALTESE"
MyArray(1,72)="1082"
MyArray(0,73)="NORWEGIAN - BOKMÅL"
MyArray(1,73)="1044"
MyArray(0,74)="POLISH"
MyArray(1,74)="1045"
MyArray(0,75)="PORTUGUESE - STANDARD"
MyArray(1,75)="2070"
MyArray(0,76)="PORTUGUESE - BRAZIL"
MyArray(1,76)="1046"
MyArray(0,77)="RAETO-ROMANCE"
MyArray(1,77)="1047"
MyArray(0,78)="ROMANIAN"
MyArray(1,78)="1048"
MyArray(0,79)="ROMANIAN - MOLDOVA"
MyArray(1,79)="2072"
MyArray(0,80)="RUSSIAN"
MyArray(1,80)="1049"
MyArray(0,81)="RUSSIAN - MOLDOVA"
MyArray(1,81)="2073"
MyArray(0,82)="SERBIAN - CYRILLIC"
MyArray(1,82)="3098"
MyArray(0,83)="SETSUANA"
MyArray(1,83)="1074"
MyArray(0,84)="SLOVENIAN"
MyArray(1,84)="1060"
MyArray(0,85)="SLOVAK"
MyArray(1,85)="1051"
MyArray(0,86)="SORBIAN"
MyArray(1,86)="1070"
MyArray(0,87)="SPANISH - STANDARD"
MyArray(1,87)="1034"
MyArray(0,88)="SPANISH - ARGENTINA"
MyArray(1,88)="11274"
MyArray(0,89)="SPANISH - BOLIVIA"
MyArray(1,89)="16394"
MyArray(0,90)="SPANISH - CHILE"
MyArray(1,90)="13322"
MyArray(0,91)="SPANISH - COLUMBIA"
MyArray(1,91)="9226"
MyArray(0,92)="SPANISH - COSTA RICA"
MyArray(1,92)="5130"
MyArray(0,93)="SPANISH - DOMINICAN REPUBLIC"
MyArray(1,93)="7178"
MyArray(0,94)="SPANISH - ECUADOR"
MyArray(1,94)="12298"
MyArray(0,95)="SPANISH - GUATEMALA"
MyArray(1,95)="4106"
MyArray(0,96)="SPANISH - HONDURAS"
MyArray(1,96)="18442"
MyArray(0,97)="SPANISH - MEXICO"
MyArray(1,97)="2058"
MyArray(0,98)="SPANISH - NICARAGUA"
MyArray(1,98)="19466"
MyArray(0,99)="SPANISH - PANAMA"
MyArray(1,99)="6154"
MyArray(0,100)="SPANISH - PERU"
MyArray(1,100)="10250"
MyArray(0,101)="SPANISH - PUERTO RICO"
MyArray(1,101)="20490"
MyArray(0,102)="SPANISH - PARAGUAY"
MyArray(1,102)="15370"
MyArray(0,103)="SPANISH - EL SALVADOR"
MyArray(1,103)="17418"
MyArray(0,104)="SPANISH - URUGUAY"
MyArray(1,104)="14346"
MyArray(0,105)="SPANISH - VENEZUELA"
MyArray(1,105)="8202"
MyArray(0,106)="SUTU"
MyArray(1,106)="1072"
MyArray(0,107)="SWEDISH"
MyArray(1,107)="1053"
MyArray(0,108)="SWEDISH - FINLAND"
MyArray(1,108)="2077"
MyArray(0,109)="THAI"
MyArray(1,109)="1054"
MyArray(0,110)="TURKISH"
MyArray(1,110)="1055"
MyArray(0,111)="TSONGA"
MyArray(1,111)="1073"
MyArray(0,112)="UKRANIAN"
MyArray(1,112)="1058"
MyArray(0,113)="URDU - PAKISTAN"
MyArray(1,113)="1056"
MyArray(0,114)="VIETNAMESE"
MyArray(1,114)="1066"
MyArray(0,115)="XHOSA"
MyArray(1,115)="1076"
MyArray(0,116)="YIDDISH"
MyArray(1,116)="1085"
MyArray(0,117)="ZULU"
MyArray(1,117)="1077"

Sub window_onload()
	curLocale= GetLocale()
	SetLocale("1033") '1033
	'test.innerText = document.all("cMonth").value 'qsDate

	Set objMonth = document.all("cMonth")
	Call GetItOn(objMonth)
	For x = 0 to Ubound(MyArray,2)
		If cStr(MyArray(1,x)) = cStr(curLocale) Then
			StrLocaleInfo =  "Locale: " & MyArray(0,x) & " (" & MyArray(1,x) & ")"
		End If
	Next
	document.all("spnLocale").innerhtml = StrLocaleInfo
	document.all("spnLocale").style.display="none"
End sub

Sub window_onunload()
	SetLocale(curLocale)
End sub

Dim qsDate,strCal,TheDate,i,strFormatDate,LoadedDate
Dim strThisIsChecked,strSelectedOption

CONST cFORMATDATE = true

if isdate(parent.modalDialog.getArgument(null)) then
	qsDate = parent.modalDialog.getArgument(null)
else
	qsDate = Date
end if
LoadedDate=qsDate

Sub ReturnMe(objMe)
	Dim A,B,C,D,E,F,G,arrDate
	
	A = ObjMe.name ' Day
	B = chkDate.value ' What to Do
	C = document.all("cMonth").value ' Month
	D = document.all("cYear").value ' Year
	
	E = C & "/" & A & "/" & D
	E = cDate(E)
	
	Select Case B
		Case 0 'full
			F = FormatDateTime(E,1)
		Case 1,2
			'SetLocale(curLocale)
			F = FormatDateTime(E,2)
			F = cDate(F)
	End Select
	
	G = cstr(F)

	If B = 0 Then
		'window.returnvalue = F
         parent.cmdDateResult(F)
		 parent.modalDialog.cancel()
	Else
'MaHamilton: what is this?!
'see below for replacement logic
'
'		arrDate = split(G,"/")
'	
'		If Len(arrDate(0)) < 2 Then
'			arrDate(0) = "0" & arrDate(0)
'		end If
'	
'		if Len(arrDate(1)) < 2 Then
'			arrDate(1) = "0" & arrDate(1)
'		end if
'	
'		If B = 1 Then 'long
'			arrDate(2) = DatePart("yyyy",F)
'		Else 'short
'			If Len(arrDate(2))=4 Then
'				arrDate(2) = Right(arrDate(2),2)
'			Else
'				arrDate(2) = arrDate(2)
'			End If
'		End If
'		
'		window.returnvalue = arrDate(0) & "/" & arrDate(1) & "/" & arrDate(2) 
		dim myDate
		myDate = Right("00" & DatePart("m", F), 2) & "/" & Right("00" & DatePart("d", F), 2) & "/" & Right("0000" & DatePart("yyyy", F), 4) & ""
		'window.returnvalue = myDate
        parent.cmdDateResult(myDate)
		parent.modalDialog.cancel()
	End If

	window.close
end sub

Sub GetItOn(objMe)

	If objMe.name = "cYear" Then
		qsDate = document.all("cMonth").value & "/1/" & objMe.value
		StartMe		
	ElseIf objMe.name = "cMonth" Then
		qsDate = objMe.value & "/1/" & document.all("cYear").value
		StartMe
	ElseIf objMe.name = "prev" Then
		theDate = document.all("cMonth").value & "/1/" & document.all("cYear").value
		qsDate = dateadd("m",-1,theDate)
		StartMe
	ElseIf objMe.name = "prev2" Then
		theDate = document.all("cMonth").value & "/1/" & document.all("cYear").value
		qsDate = dateadd("yyyy",-1,theDate)
		StartMe
	ElseIf objMe.name = "next" Then
		theDate = document.all("cMonth").value & "/1/" & document.all("cYear").value
		qsDate = dateadd("yyyy",1,theDate)
		StartMe
	ElseIf objMe.name = "next2" Then
		theDate = document.all("cMonth").value & "/1/" & document.all("cYear").value
		qsDate = dateadd("m",1,theDate)
		StartMe
	Else
	
	End If

	'strThisIsChecked = chkLongDay.value

End Sub


Sub StartMe
    dim strThisisit
	If isDate(qsDate) then

		 'Start the Border Table
		strCal ="<center><table cellspacing=0 border=1 cellpadding=0 bordercolor=Black>" & vbcrlf
		strCal = strCal & " <tr>" & vbcrlf
		strCal = strCal & "  <td width=240>" & vbcrlf

		 DrawCalendarMonth (qsDate)

		 'Close the Border Table
		strCal = strCal & "  </td>" & vbcrlf
		strCal = strCal & "</table>" & vbcrlf
	else
		strCal = strCal & "<font face=verdana, arial>" & vbcrlf
		if qsDate="" then
			strCal = strCal & "You didn't enter a date. Append &quot;?date=&quot;, followed by the date to the URL.<br><br>Example: http://server/calDraw.asp?date=2/7/01<br>" & vbcrlf
		else
			strCal = strCal & "The Date you entered (" & qsDate & ") is not valid.<br>" & vbcrlf
		end if
	end if
End Sub

Sub DrawCalendarMonth(theDate)
	dim thisMonthFirstDay
	dim nextMonthFirstDay
	dim thisMonthLastDay
	dim lastMonthLastDay
	dim calBeginDate
	dim counter
	dim CurrentMonth
	
	if month(theDate) = month(LoadedDate) and year(theDate) = year(LoadedDate) then
		CurrentMonth=true
	else
		CurrentMonth = false
	end if

	'Set the Date variables
	thisMonthFirstDay=cDate(month(theDate) & "/1/" & year(theDate))
	nextMonthFirstDay=dateAdd("m",1,thisMonthFirstDay)
	thisMonthLastDay=dateadd("d",-1,nextMonthFirstDay)
	lastMonthLastDay=dateadd("d",-1,thisMonthFirstDay)
	calBeginDate=dateadd("d",1-weekday(thisMonthFirstDay),thisMonthFirstDay)

	'Start the Calendar Table
	strCal = strCal & " <table width=240 bordercolor=white border=1 style=""border-collapse:collapse"">" & vbcrlf
	strCal = strCal & " <tr bgcolor=white>" & vbcrlf
	strCal = strCal & " <td colspan=3 align=center><font face=verdana, arial size=2><b>" & Monthname(month(theDate),true) & " " & right(year(theDate),4) & "</b></font></td>" & vbcrlf
	strCal = strCal & " <td colspan=4>"
	strCal = strCal & " <SELECT class=""INPUTComb"" onChange=""vbScript:Call GetItOn(Me)"" id=""cMonth"" name=""cMonth"">"
	
	'Create the Months
	For i = 1 to 12
		strCal = strCal & "<OPTION Value=" & i
	If i = cdbl(month(theDate)) then
		strCal = strCal & " selected "
	Else
	End if
	
	strCal = strCal & ">" & monthname(i) & "</OPTION>"
	
	Next
	strCal = strCal & "</SELECT>"

	strCal = strCal & "<SELECT class=""INPUTComb"" onChange=""vbScript:Call GetItOn(Me)"" id=""cYear"" name=""cYear"">"
	
	'Create the Years.
	For i = 2099 to 2000 step -1
	
	strCal = strCal & "<OPTION Value=" & i
		If i = clng(right(year(theDate),4)) then
			strCal = strCal & " selected "
		Else
		End if

		strCal = strCal & ">" & i & "</OPTION>"

	Next

	strCal = strCal & "</SELECT>"
	
	strCal = strCal & "</td>" & vbcrlf
	strCal = strCal & "</tr>" & vbcrlf


	strCal = strCal & "    <tr bgcolor=#BED09E>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Su</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Mo</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Tu</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>We</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Th</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Fr</td>" & vbcrlf
	strCal = strCal & "     <td width=""14.28%"" align=center class=Days>Sa</td>" & vbcrlf
	strCal = strCal & "    </tr>" & vbcrlf

	 'Start the First Row of Days
	strCal = strCal & "    <tr>" & vbcrlf
dim x
x=0

	 ' If the First day of the month is not Sunday, draw previous month's ending days
	if weekday(thisMonthFirstDay)>1 then
		x=x+1
		For counter = day(calBeginDate) to day(lastMonthLastDay)
			DrawOtherDay (counter)
		Next
	end if

	 ' Draw each day of the specified month. After each Saturday, end the row & start a new one
	 For Counter=1 to day(thisMonthLastDay)
		DrawNormalDay counter,CurrentMonth
		If weekday(cDate(month(theDate) & "/" & counter & "/" & year(theDate))) = 7 then
			x=x+1
			strCal = strCal & "    </tr>" & vbcrlf
			strCal = strCal & "    <tr>" & vbcrlf
		End if
	 Next

	 ' If the Last day of the month is not Saturday, draw next month's beginning days
	 If weekday(thisMonthLastDay)<7 then
		x=x+1
		For counter = 1 to 7-weekday(thisMonthLastDay)
			DrawOtherDay (counter)
		Next
	 End If

	 'End the Last Row and the Calendar
	strCal = strCal & "    </tr>" & vbcrlf
	
	If x < 7 then
		'If weekday( month(qsDate) & "/" & day(thisMonthLastDay) & "/" & year(qsDate) ) =7 Then
		'Else
			strCal = strCal & "    <tr><td class=NoDay>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>" & vbcrlf
		'End IF
	End If

If cFORMATDATE = TRUE THEN

If document.all("chkDate").value <> "" Then
	strSelectedOption = document.all("chkDate").value
Else
	strSelectedOption = 1
End If

	strFormatDate = "<SELECT name=chkLongDay id=chkLongDay style=""display:none;width:110;font-family:verdana;font-size:9px;"" onChange=""vbScript:Call ChangeDateType(Me)"">" 
	strFormatDate = strFormatDate & "<OPTION value=0 " & SelectedText(strSelectedOption,0) & ">Full Date</OPTION>" 
	strFormatDate = strFormatDate & "<OPTION value=1 selected" & SelectedText(strSelectedOption,1) & ">Long Date (yyyy)</OPTION>" 
	strFormatDate = strFormatDate & "<OPTION value=2 " & SelectedText(strSelectedOption,2) & ">Short Date (yy)</OPTION>"
	strFormatDate = strFormatDate & "</SELECT>"

END IF

	strCal = strCal & "    <tr bgcolor=ivory valign=top>" &  vbcrlf
	strCal = strCal & "     <td style=""cursor:hand;"" onmouseover=""me.style.backgroundColor='#BED09E'"" onmouseout=""me.style.backgroundColor=''"" align=center><font style=""font-family:webdings;font-size:16px;"" id=prev2 name=prev2 onclick=""vbScript:Call GetItOn(Me)"">7</font></td>" & vbcrlf
	strCal = strCal & "     <td style=""cursor:hand;"" onmouseover=""me.style.backgroundColor='#BED09E'"" onmouseout=""me.style.backgroundColor=''"" align=center><font style=""font-family:webdings;font-size:16px;"" id=prev name=prev onclick=""vbScript:Call GetItOn(Me)"">3</font></td>" & vbcrlf
	strCal = strCal & "     <td style=""cursor:hand;"" colspan=3 align=center onmouseover=""me.style.backgroundColor='#BED09E'"" onmouseout=""me.style.backgroundColor=''"" onclick=""vbScript:Call JumpToToday()"">" & "<font size=2 face=verdana><b>Today</b></font>" & "</td>" & vbcrlf  'strFormatDate
	strCal = strCal & "     <td style=""cursor:hand;"" onmouseover=""me.style.backgroundColor='#BED09E'"" onmouseout=""me.style.backgroundColor=''"" align=center><font style=""font-family:webdings;font-size:16px;"" id=next2 name=next2 onclick=""vbScript:Call GetItOn(Me)"" >4</font></td>" & vbcrlf
	strCal = strCal & "     <td style=""cursor:hand;"" onmouseover=""me.style.backgroundColor='#BED09E'"" onmouseout=""me.style.backgroundColor=''"" align=center><font style=""font-family:webdings;font-size:16px;"" id=next name=next onclick=""vbScript:Call GetItOn(Me)"">8</font></td>" & vbcrlf
	strCal = strCal & "    </tr>" & vbcrlf

	strCal = strCal & "    <tr><td colspan=7><span style=""font-family:verdana;font-size:9px;"" name=spnLocale style=""Display:none"" id=spnLocale>" & StrLocaleInfo & "</span></td></tr>"
	strCal = strCal & "   </table>" & vbcrlf
	
document.all("calendar").innerhtml = strCal
x=0 
end sub

Sub JumpToToday()
	qsDate = now
	LoadedDate = now
	StartMe
end sub

Sub ChangeDateType(objMe)
	document.all("chkDate").value = objMe.value
End Sub

Function SelectedText(val1,val2)
	If clng(val1) = clng(val2) Then
		SelectedText = " SELECTED "
	Else
		SelectedText = " "
	End If
End Function

Sub DrawNormalDay(DayNumber,CurrentMonth)
	' Draws a day cell - date is in current month
	' The response.write's are separate lines for clarity only
	strCal = strCal & "<td   name=" & DayNumber & " id=" & DayNumber & " onclick=""vbScript:Call ReturnMe(Me)"" class=CalDay onmouseover=""me.style.backgroundColor='gold'"" "
	strCal = strCal & "onmouseout=""me.style.backgroundColor='#D5D1C8'""> "
	if CurrentMonth and DayNumber = Day(LoadedDate) then
		strCal = strCal & "<font color=red>" & DayNumber & "</font></td>" & vbcrlf
	else
		strCal = strCal & DayNumber & "</td>" & vbcrlf
	end if
End Sub

Sub DrawOtherDay(DayNumber)
	' Draws a day cell - date is in previous or next month
	' The response.write's are separate lines for clarity only
	strCal = strCal & "<td class=OffCalDay name=" & DayNumber & " id=" & DayNumber & ">" & DayNumber & "</td>" & vbcrlf
End Sub

</SCRIPT>

</HEAD>
<BODY bgcolor=Ivory>
<span ID=test></span>
<BR>
	<DIV NAME=calendar ID=calendar></DIV>
	<input type=hidden name=chkDate id=chkDate value="1">
<SCRIPT Language="vbscript">
Call StartMe
</SCRIPT>
</BODY>
</HTML>



