<%@ Language=VBScript %>
  var SelectedRow;
  function Export(strID) {
    if (txtCurrentFilter.value == "")
      window.open(window.location.href + "?FileType=" + strID);
    else
      window.open(window.location.pathname + "?" + txtCurrentFilter.value + "&FileType=" + strID);
  }
  function button1_onclick(NewColor) {
    //MyTable.borderColor = NewColor;
    Row1.bgColor = NewColor;
    Row2.bgColor = NewColor;
  }
  function button2_onclick() {
    COL1.style.width = 0;
    //COL1.width=10;
  }

  /* *** LY BEGINNING OF CHANGE - ADD DATE RANGE DIALOG TO BATTERY REPORT *** */
  function ChooseDates(StartDate, EndDate) {
    var strID = window.showModalDialog("../Query/daterange.asp?StartDate=" + StartDate + "&EndDate=" + EndDate, "", "dialogWidth:370px;dialogHeight:250px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
    if (typeof (strID) != "undefined") {
      window.location.href = "HardwareMatrix_BatteryReport.asp?lstCategory=71&DateStart=" + strID[0] + "&DateEnd=" + strID[1];
    }
  }
  /*  *** LY END OF CHANGE - ADD DATE RANGE DIALOG TO BATTERY REPORT *** */

  function Commodity_onclick() {
    var RowElement;
    RowElement = window.event.srcElement;
    while (RowElement.className != "Row") {
      RowElement = RowElement.parentElement;
    }
    if (RowElement.style.backgroundColor == "cornflowerblue")//lightgoldenrodyellow
      RowElement.style.backgroundColor = "";
    else
      RowElement.style.backgroundColor = "cornflowerblue"; //lightgoldenrodyellow
    if (typeof (SelectedRow) != "undefined") {
      if (SelectedRow != RowElement)
        SelectedRow.style.backgroundColor = "";
    }
    SelectedRow = RowElement;
  }
  function Commodity_onmouseover() {
    var RowElement;
    RowElement = window.event.srcElement;
    while (RowElement.className != "Row") {
      RowElement = RowElement.parentElement;
    }
    if (RowElement.style.backgroundColor == "") {
      RowElement.style.backgroundColor = "#99ccff";
      RowElement.style.cursor = "hand";
    }
  }
  function Commodity_onmouseout() {
    var RowElement;
    RowElement = window.event.srcElement;
    while (RowElement.className != "Row") {
      RowElement = RowElement.parentElement;
    }
    if (RowElement.style.backgroundColor == "#99ccff") {
      RowElement.style.backgroundColor = "";
    }
  }
  function DisplayVersion(VersionID) {
    var strResult;
    strResult = window.showModalDialog("../../WizardFrames.asp?Type=1&ID=" + VersionID, "", "dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
    if (typeof (strResult) != "undefined") {
      window.location.reload(true);
    }
    if (typeof (SelectedRow) != "undefined") {
      SelectedRow.style.backgroundColor = "";
    }
  }
  function SwitchFilterView(strType) {
    if (strType == 1) {
      QuickLinks.style.display = "none";
      FilterBox.style.display = "";
    }
    else if (strType == 2) {
      QuickLinks.style.display = "";
      FilterBox.style.display = "none";
    }
  }
  function window_onload() {
    lblLoad.style.display = "none";
    if (txtScrollToRow.value != "")
      document.all("Row" + txtScrollToRow.value).scrollIntoView();
    window.name = "HardwareMatrix";
    //self.focus();
  }
  function Test(ID) {
    DisplayArea.innerHTML = divQuickReports.innerHTML;
  }