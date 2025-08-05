$(document).ready(function() {

	WinHeight = $(document).height()-535;
	WinWidth = $(document).width();
	$("div#payrollholder").css("min-height","700");
	$("div#footer").css("width",WinWidth);
	


	$("div#holder").css("display","none");
	//PAYROLL EDIT
	$("td.edit_payroll").editInPlace({
	    url: "ax-editpayroll.aspx",
	    params: "",
		bg_over:"",
		on_blur:"cancel",
		show_buttons: true,
		default_text:"",
		
		callback: function(original_element, html, original){
			var DataString = "id=" + original_element + "&value=" + html;
			return SubmitAjax(DataString,"ax-editpayroll.aspx","");
        }

	});
	
	//PAYROLL NOTES EDIT
	$("td.edit_payroll_notes").editInPlace({
	    url: "ax-editpayroll.aspx",
	    params: "",
		bg_over:"",
		on_blur:"cancel",
		show_buttons: true,
		default_text:"",
		field_type:"textarea",
		textarea_cols:15,
		textarea_rows:2,
		
		callback: function(original_element, html, original){
			var DataString = "id=" + original_element + "&value=" + html;
			return SubmitAjax(DataString,"ax-editpayroll.aspx","");
        }
	});	
	
	
	//INIT TABLE SORTER
	$("table.payroll").tablesorter({
	headers: { 
		5: { sorter: false}, 
		6: {sorter: false},
		7: { sorter: false}, 
		8: {sorter: false}, 
		9: { sorter: false}, 
		10: {sorter: false}, 
		11: { sorter: false}, 
		12: {sorter: false}, 
		13: { sorter: false}, 
		14: {sorter: false}, 
		15: { sorter: false}, 
		16: {sorter: false}, 
		17: { sorter: false}, 
		18: {sorter: false}, 
		19: { sorter: false}, 
		20: {sorter: false}, 
		21: { sorter: false}, 
		22: {sorter: false}, 
		23: { sorter: false}, 
		24: {sorter: false}, 
		25: { sorter: false}, 
		26: {sorter: false}, 
		27: { sorter: false}, 
		28: {sorter: false}, 
		29: { sorter: false}, 
		30: {sorter: false}, 
		31: { sorter: false}, 
		32: {sorter: false}, 
		33: { sorter: false}, 
		34: {sorter: false}, 
		35: { sorter: false},
		36: { sorter: false},
		37: { sorter: false}, 
		38: {sorter: false},
		39: { sorter: false}, 
		40: {sorter: false}, 
		41: { sorter: false}, 
		42: {sorter: false}, 
		43: { sorter: false}, 
		44: {sorter: false}, 
		45: { sorter: false}, 
		46: {sorter: false}, 
		47: { sorter: false}, 
		48: {sorter: false}, 
		49: { sorter: false}, 
		50: {sorter: false}, 
		51: { sorter: false}, 
		52: {sorter: false}, 
		53: { sorter: false}, 
		54: {sorter: false}, 
		55: { sorter: false}, 
		56: {sorter: false}, 
		57: { sorter: false}, 
		58: {sorter: false}, 
		59: { sorter: false}, 
		60: {sorter: false}, 
		61: { sorter: false}, 
		62: {sorter: false}, 
		63: { sorter: false}, 
		64: {sorter: false}, 
		65: { sorter: false}, 
		66: {sorter: false}, 
		67: { sorter: false}		
	}
	}); 
	
	
});