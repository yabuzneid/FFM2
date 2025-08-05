$(document).ready(function() {

	//NOTES FIELD EDIT
	$('div.weeknotes').editable(function(value, settings) { 
		var ArrID = $(this).attr("id").split("_");
		var UserID = ArrID[1];
		var WeekStart = ArrID[2];
		var DataString = "employeeid=" + UserID + "&weekstart=" + WeekStart + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx?action=savenotes","");
		return(value);
		}, {
		type      : 'textarea',
		height    : '40',
		width     : '350',
		indicator : 'Saving...',
		tooltip   : 'Click to edit...',
		cancel    : 'Cancel',
		submit    : 'Save Notes',
		placeholder : ''
	});

	//DEFINE MASKED INPUT TYPE
	$.editable.addInputType('masked', {
	    element : function(settings, original) {
	        /* Create an input. Mask it using masked input plugin. Settings  */
	        /* for mask can be passed with Jeditable settings hash.          */
	        var input = $('<input />').mask(settings.mask);
	        $(this).append(input);
	        return(input);
	    }
	});

	//FIGURE OUT IF TEXT EDIT SHOULD BE LOCKED OR NOT
	$('td.edit').click(function(){
		
		var OTVal = $(this).siblings("td.edit_overtime").html();
		var FieldTypeArr = $(this).attr("id").split("_");
		var FieldType = FieldTypeArr[0];
		
		if ((FieldType == 'jobnumber')||(FieldType == 'state')) {
			if ((OTVal == 0) || (OTVal == null)) {
	            $(this).trigger("dblClick");
			}
		} else {
			$(this).trigger("dblClick");
		}
		
	});

	//REGULAR TEXT EDIT
	$('td.edit').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		//$(this).unbind(settings.event);
		return(value);
		}, {
		height    : '16',
		indicator : 'Saving...',
		tooltip   : 'Click to edit...',
		cancel    : 'Cancel',
		submit    : 'OK',
		placeholder : '',
		event		: 'dblClick'
	});

	//EDIT JOB NUMBER - SHOW FORM
	$('td.edit_jobnumber').livequery('click',function() {

		var TimeEntryID = $(this).attr("id");
		TimeEntryID = TimeEntryID.substring(10);
		$(this).attr("alt", $(this).html());

		var DataString = "timeentryid=" + TimeEntryID +"&jobnumber=" + $(this).html() + "&id=" + $(this).attr("id");
		var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx?action=showjobnumberform","")

		$(this).html(Response);
		$(this).removeClass("edit_jobnumber");
		
	});

	//EDIT JOB NUMBER - SUBMIT
	$('input.jobnum_submit').livequery('click',function() {

		var TimeEntryID = $(this).attr('name') 
		TimeEntryID = TimeEntryID.substring(14);

		var jobnum1 = $('input#jobnum_1_' + TimeEntryID).val();
		var jobnum2 = $('select#jobnum_2_' + TimeEntryID).val();
		var jobnum3 = $('input#jobnum_3_' + TimeEntryID).val();
		var jobnum4 = $('input#jobnum_4_' + TimeEntryID).val();

		var IsValid = true;

		if (jobnum1== "01") { 
			if (jobnum2 == "") { IsValid = false; }
			if (jobnum3.length != 6) { IsValid = false; }
		}

		if (jobnum1== "02") { 
			if (jobnum2 == "") { IsValid = false; }
			if (jobnum3.length != 4) { IsValid = false; }
		}


		if (IsValid == true) {
			var DataString = "timeentryid=" + TimeEntryID + "&jobnum1=" + jobnum1 + "&jobnum2=" + jobnum2 + "&jobnum3=" + jobnum3 + "&jobnum4=" + jobnum4 + "&id=" + $(this).attr("id");
			var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx?action=savejobnumberform","");

			$(this).parent().parent().addClass("edit_jobnumber");
			$(this).parent().parent().removeClass("invalidjobnumber");
			$(this).parent().parent().html(Response);
			$(this).parent().parent().attr("alt","");
		} else {
			alert("Job Number is not in the correct format. \n \n Correct format is: \n 01 - XX - XXXXXX ( - XXX) \n or \n 02 - XX - XXXX  \n for general ledger employees");
		}
		
	});

	//EDIT JOB NUMBER - CANCEL
	$('input.jobnum_cancel').livequery('click',function() {
		var oldJobNum = $(this).parent().parent().attr("alt");
		$(this).parent().parent().addClass("edit_jobnumber");
		$(this).parent().parent().html(oldJobNum);
		$(this).parent().parent().attr("alt","");
	});


	
	//FIGURE OUT IF DATE SHOULD BE LOCKED OR NOT
	$('span.edit_date').click(function(){
		
		var OTVal = $(this).parent().parent().prevAll(".shift").children(".edit_overtime").html();
		
		var FieldTypeArr = $(this).attr("id").split("_");
		var FieldType = FieldTypeArr[0];
		
		if ((FieldType == 'starttime')) {
			if ((OTVal == 0) || (OTVal == null)) {
	            $(this).trigger("dblClick");
			}
		} else {
			$(this).trigger("dblClick");
		}
		
	});
	
	//DATE EDIT
	$('.edit_date').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx","");
		if (Response.substr(0,5) != "ERROR") {
			return(value);
		}
		}, {
		type      : 'masked',
        mask      : '99/99/9999 99:99',
		indicator : 'Saving...',
		tooltip   : 'Click to edit...',
		cancel    : 'Cancel',
		submit    : 'OK',
		placeholder : '',
		event 		: 'dblClick'
	});
	
	//EDIT TYPE
	$('td.edit_type').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, { 
		data   : " {'':'','KC':'KC','KA':'KA','G':'G'}",
		type   : 'select',
		submit : 'OK',
		placeholder : ''
	});
	
	//EDIT SHIFT TYPE
	$('td.edit_shifttype').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, { 
		    data: " {'':'','Regular Shift':'Regular Shift','Travel':'Travel','Holiday':'Holiday','Personal':'Personal','Vacation':'Vacation','Sick':'Sick','PTO':'PTO','FMLA':'FMLA','FFCRA':'FFCRA','Bereavement':'Bereavement'}",
		type   : 'select',
		submit : 'OK',
		placeholder : ''
	});

	//EDIT TRAVEL
	$('td.edit_travel').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, { 
		data   : " {'':'','N/A':'N/A','Company':'Company','Personal':'Personal'}",
		type   : 'select',
		submit : 'OK',
		placeholder : ''
	});
	
	
	//FIGURE OUT IF PW SHOULD BE LOCKED OR NOT
	$('td.edit_yesno').click(function(){
		
		var OTVal = $(this).siblings("td.edit_overtime").html();
		var FieldTypeArr = $(this).attr("id").split("_");
		var FieldType = FieldTypeArr[0];
		
		if (FieldType == 'pw') {
			if ((OTVal == 0) || (OTVal == null)) {
	            $(this).trigger("dblClick");
			}
		} else {
			$(this).trigger("dblClick");
		}
		
	});
	
	//EDIT PERDIEM, PW, INJURED
	$('td.edit_yesno').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, { 
		data   : " {'':'','Yes':'Yes','No':'No'}",
		type   : 'select',
		submit : 'OK',
		placeholder : '',
		event	: 'dblClick'
	});	
	
	//EDIT INJURED
	$('td.edit_injured').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, { 
		data   : " {'':'','Yes':'Yes','No':'No'}",
		type   : 'select',
		submit : 'OK',
		placeholder : ''
	});	
	
	//OVERTIME EDIT
	$('td.edit_overtime').editable(function(value, settings) { 
		
		
		//CheckOTEditable($(this));
		
		var TimeEntryID = $(this).attr("alt");
		
		var JobNumber = $(this).siblings("td#jobnumber_" + TimeEntryID).html();
		var PW = $(this).siblings("td#pw_" + TimeEntryID).html();
		var State = $(this).siblings("td#state_" + TimeEntryID).html();
		var StartDate = $(this).parent().siblings().children().children("span#starttime_" + TimeEntryID).html().substr(0,10);
		
		var ArrOldKey = $(this).attr("id").split("_")
		
		var UserID = ArrOldKey[1];
		var OldStartDate = ArrOldKey[3];
		
		
		
		var NewKey = "ot_" + UserID + "_" + JobNumber + "_" + StartDate + "_" + PW + "_" + State;
				
		var DataString = "id=" + NewKey + "&value=" + value;
		SubmitAjax(DataString,"ax-editpayroll.aspx","")
		return(value);
		}, {
		indicator : 'Saving...',
		tooltip   : 'Click to edit...',
		cancel    : 'Cancel',
		submit    : 'OK',
		placeholder : ''
	});
	
	//COMMENTS EDIT
	$('td.edit_comments').editable(function(value, settings) { 
		var DataString = "id=" + $(this).attr("id") + "&value=" + value;
		SubmitAjax(DataString,"ax-edittimesheet.aspx","")
		return(value);
		}, {
		type      : 'textarea',
		height    : '50',
		width     : '100',
		indicator : 'Saving...',
		tooltip   : 'Click to edit...',
		cancel    : 'Cancel',
		submit    : 'OK',
		placeholder : ''
	});
	
	//APPROVE BUTTON
	$("input.approvebutton").livequery('click', function() {
		var ArrApproveID = $(this).attr("id").split("_");
		var ApproveID = ArrApproveID[1];

		var JobNum = $("td#jobnumber_" + ApproveID).html();
		
		if ((ValidateJobNum(JobNum) == true) || ($("td#jobnumber_" + ApproveID).length == 0)) {
			var DataString = "";
			var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx?action=approve&id=" + ApproveID,"");
			$(this).parent().html(Response);
		} else {
			alert("Job Number is not in the correct format. \n \n Correct format is: \n 01 - XX - XXXXXX ( - XXX) \n or \n 02 - XX - XXXX  \n for general ledger employees");
		}


	});

	//VALIDATE JOB NUMBER
	function ValidateJobNum(JobNum) {


		var JobNumArr = String(JobNum).split("-");
		var IsValid = true;

		if (JobNumArr.length > 1) {	

			if (JobNumArr[0] == "01") { 
				if (JobNumArr[1] == "") { IsValid = false; }
				if (JobNumArr[2].length != 6) { IsValid = false; }
			}

			if (JobNumArr[0] == "02") { 
				if (JobNumArr[1] == "") { IsValid = false; }
				if (JobNumArr[2].length != 4) { IsValid = false; }
			}

		} else {
			IsValid = false;
		}

		return IsValid;
	}
	
	//WORKER DROPDOWN
	$("select#worker_ddl").change(function() {
		document.location.href = $("select#worker_ddl option:selected").val();
	});
	
	//UNAPPROVE
	$("a.unapprove").livequery('click',function() {
		var ArrApproveID = $(this).attr("id").split("_");
		var ApproveID = ArrApproveID[1];
		var DataString = "";
		var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx?action=unapprove&id=" + ApproveID,"");
		$(this).parent().parent().html(Response);
		return false;
	});
	
	//SAVE NOTES
	$("input.saveweeknotes").click(function(){
		var ArrID = $(this).attr("id").split("_");
		var UserID = ArrID[1];
		var WeekStart = ArrID[2];
		var Value = $("textarea#weeknotes").val();
		var DataString = "employeeid=" + UserID + "&weekstart=" + WeekStart + "&value=" + Value;
		var Response = SubmitAjax(DataString,"ax-edittimesheet.aspx?action=savenotes");
		
		if (Response == "SUCCESS") {
		
		} else {
			alert(Response);
		}

	});	

	// ADDITIONAL TIME ENTRY INFO

	$(".tei-trigger").click(function(){
		var teiID = $(this).attr("id");
		teiID = teiID.substring(12);
		console.log(teiID);
		$("#tei-holder-" + teiID).toggle();
		return false;
	});

});