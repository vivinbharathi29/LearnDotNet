function newDate(passedValue) 
{   
   var firstSlash=passedValue.indexOf("/")
   var lastSlash=passedValue.lastIndexOf("/")
   var month=passedValue.substr(0,firstSlash)-1
   var day=passedValue.substring(firstSlash+1,lastSlash)
   var year1= new Number(passedValue.substr(lastSlash+1))  
   if ( year1 > 70 )
   		var year=year1 + 1900		
   else	
		var year=year1 + 2000	
   var newDate = new Date(year,month,day)
   return newDate
}

function popCalendar(cntrlname,someDate)
{
	if (someDate == "")
	{
		newdate = new Date();
		someDate = newdate.toGMTString();
	}
	if (someDate == "1/1/1900")
	{
		newdate = new Date();
		someDate = newdate.toGMTString();
	}

	var url = "/ipulsar/PopupCalendar.aspx?cntrl="+cntrlname + "&currentDate=" + escape(someDate);
	var popbox = window.open(url, "Calendar", "resizable=no," +
							"toolbar=no,"+
							"scrollbars=no,"+
							"directories=no,"+
							"menubar=no,"+
							"titlebar=no,"+
							"width=195,"+
							"top=400,"+
							"left=400,"+
							"alwaysRaised=yes,"+
							"dependent=yes,"+
							"height=195"); 
	popbox.opener = self;
}

function ASPpopCalendar(cntrlname,someDate)
{
	if (someDate == "")
	{
		newdate = new Date();
		someDate = newdate.toGMTString();
	}
	thefile = "../library/widget/popupcalendar.asp?cntrl=" + cntrlname + "&currentDate=" + escape(someDate);
	popbox=window.open(thefile,"Calendar","resizable=yes,toolbar=no,scrollbars=no,directories=no,menubar=no,width=250,height=260");
	//popbox=window.open(thefile,"Calendar","resizable=yes,toolbar=no,scrollbars=no,directories=no,menubar=no,width=250,height=240");			
	if(popbox !=null)
	{
		if (popbox.opener==null)
		{
			popbox.opener=self;
		}
	}
}	

// validate the dates entered (MM/DD/YYYY)
function checkDate (theField, s, emptyOK) {   
	// Next line is needed on NN3 to avoid "undefined is not a number" error
	// in equality comparison below.
	if (checkDate.arguments.length == 2) 
		emptyOK = true;
	if ((isEmpty(theField.value)) || isWhitespace(theField.value)) 
		if (emptyOK == true) {
			return true;
		} else {
			return warnEmpty (theField, s);
		}
	// add a space so an empty string won't have 2 spaces in the message
	if (s.length > 0)
		s += " "

	//if (isValidDate(theField.value) == 1) 
	if (!isValidDateCheck(theField.value, false))
		return warnInvalid (theField, "Please enter a valid date for the " + s + "field in the form of MM/DD/YYYY.\nYear must be between between 1970 and 2099.");
	else 
		return true;
}

// validate the dates entered (MM/DD/YYYY)
/*
function checkValidDate (theField, s, emptyOK) {   
	// Next line is needed on NN3 to avoid "undefined is not a number" error
	// in equality comparison below.
	if (checkValidDate.arguments.length == 2) 
		emptyOK = true;
	if ((isEmpty(theField.value)) || isWhitespace(theField.value)) 
		if (emptyOK == true) {
			return true;
		} else {
			return warnEmpty (theField, s);
		}
	// add a space so an empty string won't have 2 spaces in the message
	if (s.length > 0)
		s += " "

	//if (isValidDate(theField.value) == 1) 
	if (!isValidDateCheck(theField.value, false)) {
		return warnInvalid (theField, "Please enter a valid date for the " + s + "field in the form of MM/DD/YYYY.\nYear must be between between 1970 and 2099.");
	} else 
		return true;
}
*/
function isValidDateCheck(dateStr, showError) {
	// showError = true, false (show exact error as to why it failed)
	// Checks for the following valid date formats:
	// MM/DD/YY   MM/DD/YYYY   MM-DD-YY   MM-DD-YYYY
	// Also separates date into month, day, and year variables
	
	//var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{2}|\d{4})$/;
	// To require a 4 digit year entry, use this line instead:
	var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/;
	
	var matchArray = dateStr.match(datePat); // is the format ok?
	if (matchArray == null) {
		if (showError)
			alert("Date is not in a valid format.")
		return false;
	}
	month = matchArray[1]; // parse date into variables
	day = matchArray[3];
	year = matchArray[4];
	if (month < 1 || month > 12) { // check month range
		if (showError)
			alert("Month must be between 1 and 12.");
		return false;
	}
	if (day < 1 || day > 31) {
		if (showError)
			alert("Day must be between 1 and 31.");
		return false;
	}
	if ((month==4 || month==6 || month==9 || month==11) && day==31) {
		if (showError)
			alert("Month "+month+" doesn't have 31 days.")
		return false
	}
	if (month == 2) { // check for february 29th
		var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
		if (day>29 || (day==29 && !isleap)) {
			if (showError)
				alert("February " + year + " doesn't have " + day + " days.");
			return false;
		}
	}
	if (year < 1970 || year > 2099) {
		if (showError)
			alert("Year must be between 1970 and 2099.")
		return false
	}
	return true;  // date is valid
}

function isValidDate(passedValue) {
	var err=0

	//accept only numbers and /
	var checkOK = "0123456789/";
	for (i = 0;  i < passedValue.length;  i++)
	{
		ch = passedValue.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
			if (ch == checkOK.charAt(j))
			  break;
		if (j == checkOK.length)
		{
			err=1;
			break;
		}
	}
	   
	var firstSlash=passedValue.indexOf("/")
	var lastSlash=passedValue.lastIndexOf("/")

	//verify month
	var month=passedValue.substr(0,firstSlash)
	if (month<1 || month>12) err = 1

	//verify day
	var day=passedValue.substring(firstSlash+1,lastSlash)
	if (day<1 || day>31) err = 1
	
	//verify year
	var year=passedValue.substr(lastSlash+1)
	if (year.length != 4) err = 1
	if (year < 0 || year >9999 ) err = 1
	
	//check days in 30 day months
	if (month==4 || month==6 || month==9 || month==11){
	if (day==31) err=1
	}

	//check Feb.
	if (month==2){
		var g = parseInt(year/4)
		if (isNaN(g)) {
			err=1
		}
		if (day>29) err=1
		if (day==29 && ((year/4)!=parseInt(year/4))) err=1
	}
	//return error value, 0=okay, 1=fail
	return err
}

