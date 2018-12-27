// Sidebar Js
$('.side-control').click(function(){
	if (screen.width < 767) {
	}
	else {
		if($(this).hasClass('active')){
			$(this).removeClass('active');
			$(this).parents('.page-container').removeClass('sb-collaps');
			$('.sidebar-menu .main-lnk a span').removeClass('arrow-down glyphicon-chevron-down');
		}
		else{
			$(this).addClass('active');
			$('.main-lnk').removeClass('open');
			$(this).parents('.page-container').addClass('sb-collaps');
		}
	}
});
// Sidebar Height Issues
var sidebarHeight = $('.page-sidebar').outerHeight();
$('.page-content').css({'min-height' : sidebarHeight});
$(window).resize(function(){
	var sidebarHeight = $('.page-sidebar').outerHeight();
	$('.page-content').css({'min-height' : sidebarHeight});
	var w = $(window).width();
	if(w <= 979){
		$('.main-lnk.open').removeClass('open');
		$('.main-lnk .arrow-down.glyphicon-chevron-down').removeClass('arrow-down glyphicon-chevron-down');	
	}
});
// Sidebar Expand / Collaps Js
$('.sidebar-menu .main-lnk').click(function(){
	if($(this).hasClass('open')){
		$('.main-lnk.open').removeClass('open');
		$('.glyphicon-chevron-left.arrow.arrow-down').removeClass('arrow-down glyphicon-chevron-down');
		$(this).find('.arrow.arrow-down').addClass('glyphicon-chevron-left');
	}
	else{
		$('.sidebar-menu .main-lnk.open').removeClass('open');
		$(this).addClass('open');	
		$('.arrow-down.glyphicon-chevron-down').removeClass('arrow-down glyphicon-chevron-down');
		$(this).find('.arrow').addClass('arrow-down glyphicon-chevron-down');
	}
	var sidebarHeight1 = $('.page-sidebar').outerHeight();
	$('.page-content').css({'min-height' : sidebarHeight1 });
});
$('.page-container .sidebar-menu .main-lnk').click(function(){
	if($(this).parents(".page-container").hasClass('sb-collaps')){
		$(this).removeClass('open');
	}
});

$(window).resize(function(){
	var w = $(window).width();
	if(w <= 979) {
		//alert('767');
		$('.page-container').removeClass('sb-collaps');
	}
});

// Date Time Picker JS
$('.datepicker').datetimepicker({
	showToday:true,
	pickTime:false,
	//disabledDates:[ ("2015-02-25"),("2015-02-26"),("2015-02-27")],
});	