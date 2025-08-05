$(document).ready(function() {
	WinHeight = $(document).height()-265;
	$("div#holder").css("min-height",WinHeight);

});


function SubmitAjax(DataString, PostURL, Message) {

	var ReturnString = $.ajax({  
		type: "POST",  
		url: PostURL,  
		data: DataString,
		async:false,
		success: function(returnmsg) {
			
			if (returnmsg.substr(0,5) == "ERROR") {
				alert("We are Sorry, but an error has occured. Technical details are below \n \n " + returnmsg);
				return "error";
			} else if (returnmsg == "%EXPIRED%") {
				document.location.href = "Default.aspx?action=expired";
				return false;
			}
		}
	}).responseText; 

	return ReturnString;	
}

function ValidateTimeEntry() {

	var EmployeeID = document.getElementById("CPH_MainContent_DDL_Employee").value;
	
	if (EmployeeID == -1) {
		var Confirmation = confirm("You are adding shifts for ALL EMPLOYEES, are you sure you want to proceed?");
		
		if (Confirmation) {
			return ValidateTimeEntryDates();
		} else {
			return false;
		}
		
	} else {
		return ValidateTimeEntryDates();
	}

}

function ValidateTimeEntryDates() {
	
	
	var MultiDay = document.getElementById("CPH_MainContent_RB_MultiDay").checked;
	
	if (MultiDay != true) {
		var StartDate = new Date(document.getElementById("CPH_MainContent_TXT_StartDate").value);
		var EndDate = new Date(document.getElementById("CPH_MainContent_TXT_EndDate").value);
		var StartTime = document.getElementById("CPH_MainContent_DDL_StartTime").value;
		var EndTime = document.getElementById("CPH_MainContent_DDL_EndTime").value;
	} else {
		var StartDate = new Date(document.getElementById("CPH_MainContent_TXT_StartDateMulti").value);
		var EndDate = new Date(document.getElementById("CPH_MainContent_TXT_EndDateMulti").value);
		var StartTime = document.getElementById("CPH_MainContent_DDL_StartTimeMulti").value;
		var EndTime = document.getElementById("CPH_MainContent_DDL_EndTimeMulti").value;
	}
	
	var StartTimeArr = StartTime.split(":");
	var EndTimeArr = EndTime.split(":");
	var ShowAlert = false;
	
	//alert(StartDate + "::" + StartTime + ":::::" + EndDate + "::" + EndTime);
	
	var DateDifference = EndDate - StartDate;
	
	if (MultiDay == true) {
		if (DateDifference < 0) {
			ShowAlert = true;
		} 
		
		if (StartTimeArr[0] > EndTimeArr[0]) {
			ShowAlert = true;
		} else if (StartTimeArr[0] == EndTimeArr[0]) {
			if (StartTimeArr[1] > EndTimeArr[1]) {
				ShowAlert = true;
			}
		}
	} else { 
	
		if (DateDifference < 0) {
			ShowAlert = true;
		} 

		//console.log("start date: " + StartDate + " end date: " + EndDate + " start time: " + StartTime + " end time: " + EndTime);
		//console.log("difference: " + DateDifference);
		
		if (DateDifference == 0) {
			if (StartTimeArr[0] > EndTimeArr[0]) {
				ShowAlert = true;
			} else if (StartTimeArr[0] == EndTimeArr[0]) {
				if (StartTimeArr[1] > EndTimeArr[1]) {
					ShowAlert = true;
				}
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