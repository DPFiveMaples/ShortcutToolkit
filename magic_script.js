$(document).ready(function(){
/*hide "extras" from the start*/
	$(".foreign").hide();
	$(".secure").hide();
	$(".merge").hide();
	$(".NCOA").hide();
	$(".dedupe").hide();
	$(".fcspr").hide();
	$(".ndcscf").hide();
	$(".pslabels").hide();
	$(".clientstock").hide();
	$(".letter").hide();
	$(".reply").hide();
	$(".envelope").hide();
	$(".comingle").hide();
	$(".dropselse").hide();
	
	
/*checkbox conditions*/	
	$("#SecureCheck").click(function(){
		if($("#SecureCheck").is(":checked")){
			$(".secure").show();
			$(".notsecure").hide();
		} else {
			$(".secure").hide();
			$(".notsecure").show();
		}

	});
	$("#MergeCheck").click(function(){
		if($("#MergeCheck").is(":checked")){
			$(".merge").show();
			$(".notmerge").hide();
		} else {
			$(".merge").hide();
			$(".notmerge").show();
		}

	});
		$("#LetterCheck").click(function(){
		if($("#LetterCheck").is(":checked")){
			$(".letter").show();
		} else {
			$(".letter").hide();
		}

	});
		$("#ReplyCheck").click(function(){
		if($("#ReplyCheck").is(":checked")){
			$(".reply").show();
		} else {
			$(".reply").hide();
		}

	});
		$("#EnvelopeCheck").click(function(){
		if($("#EnvelopeCheck").is(":checked")){
			$(".envelope").show();
		} else {
			$(".envelope").hide();
		}

	});
	$("#ForeignCheck").click(function(){
		if($("#ForeignCheck").is(":checked")){
			$(".foreign").show();
		} else {
			$(".foreign").hide();
		}
	});
	$("#NCOACheck").click(function(){
		if($("#NCOACheck").is(":checked")){
			$(".NCOA").show();
		} else {
			$(".NCOA").hide();
		}
	});
		$("#DeDupeCheck").click(function(){
		if($("#DeDupeCheck").is(":checked")){
			$(".dedupe").show();
		} else {
			$(".dedupe").hide();
		}
	});
		$("#FCSPRCheck").click(function(){
		if($("#FCSPRCheck").is(":checked")){
			$(".fcspr").show();
			$(".sorted").hide();
		} else {
			$(".fcspr").hide();
			$(".sorted").show();
		}
	});
		$("#NDCSCFCheck").click(function(){
		if($("#NDCSCFCheck").is(":checked")){
			$(".ndcscf").show();
		} else {
			$(".ndcscf").hide();
		}
	});
		$("#PSLabelsCheck").click(function(){
		if($("#PSLabelsCheck").is(":checked")){
			$(".pslabels").show();
		} else {
			$(".pslabels").hide();
		}
	});
		$("#ClientStockCheck").click(function(){
		if($("#ClientStockCheck").is(":checked")){
			$(".clientstock").show();
			$(".fmstock").hide();
		} else {
			$(".clientstock").hide();
			$(".fmstock").show();
		}
	});
	
		$("#DropsElseCheck").click(function(){
		if($("#DropsElseCheck").is(":checked")){
			$(".dropselse").show();
		} else {
			$(".dropselse").hide();
		}
	});
	
		$("#ComingleCheck").click(function(){
		if($("#ComingleCheck").is(":checked")){
			$(".comingle").show();
			$(".notcomingle").hide();
		} else {
			$(".comingle").hide();
			$(".notcomingle").show();
		}
	});

});	

function myFunction() {
    return "Are you sure you want to leave?";
}
	



