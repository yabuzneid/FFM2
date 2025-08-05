$(document).ready(function() {
	//DATE PICKERS
	Date.format = 'mm/dd/yyyy';
	$("input.pickdate").datePicker({startDate:'01/01/2000',showYearNavigation:false});
	
	//VALIDATE DATE SUBMIT
	$("input#CPH_MainContent_BTN_Submit").click(function() {
		return false;
	
	});
	
	//Jobnumber report submit
	$("input#submitjobnumber").click(function(){
		console.log("click");

		submitURL = "jobreport.aspx?startdate=" + $("#CPH_MainContent_startdate").val() + "&enddate=" + $("#CPH_MainContent_enddate").val() + "&jobnumber=" + $("#jobnumber").val();
		window.location = submitURL;
		
	});


	
});

function ValidateDates() {
	
	var StartDate = new Date(document.getElementById("CPH_MainContent_TXT_StartDate").value);
	var EndDate = new Date(document.getElementById("CPH_MainContent_TXT_EndDate").value);
	var StartTime = document.getElementById("CPH_MainContent_DDL_StartTime").value;
	var EndTime = document.getElementById("CPH_MainContent_DDL_EndTime").value;
	var StartTimeArr = StartTime.split(":");
	var EndTimeArr = EndTime.split(":");
	var ShowAlert = false;
	
	if (StartDate > EndDate) {
		ShowAlert = true;
	} 
	if (StartDate = EndDate) {
		if (StartTimeArr[0] > EndTimeArr[0]) {
			ShowAlert = true;
		} else if (StartTimeArr[0] == EndTimeArr[0]) {
			if (StartTimeArr[1] > EndTimeArr[1]) {
				ShowAlert = true;
			}
		}
	}
	
	if (ShowAlert == true) {
		alert("You can not enter a negative time period. End Time must be AFTER Start Time.");
		return false;
	} else {
		return true;
	}
	
}