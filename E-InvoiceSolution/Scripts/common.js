$(document).ready(function () {

    $(document).on('keydown', '.NumberField', function (e) { -1 !== $.inArray(e.keyCode, [46, 8, 9, 27, 13, 110]) || (/65|67|86|88/.test(e.keyCode) && (e.ctrlKey === true || e.metaKey === true)) && (!0 === e.ctrlKey || !0 === e.metaKey) || 35 <= e.keyCode && 40 >= e.keyCode || (e.shiftKey || 48 > e.keyCode || 57 < e.keyCode) && (96 > e.keyCode || 105 < e.keyCode) && e.preventDefault() });

});

function mAlert(message) {
    $("#spMessage").html(message);
    $('.alert').show().fadeIn();
    $("html, body").animate({ scrollTop: 30 }, 500);
}

function mAlertClose() {
    $('.alert').fadeOut('slow');
}

// Fuction for scroll body at given element
function scrollToControl(control) {
    // alert($('#' + control).offset().top);
    $('html,body').animate({ scrollTop: $('#' + control).offset().top }, 500);
}

function showLoader() {
    $('#Loader').show();
}
function hideLoader() {
    $('#Loader').hide();
}
function SlideUpDown(DivControl) {
    $('#' + DivControl).slideToggle(function () {
        $('#' + DivControl).css({ 'overflow': 'initial' });
    });
    // 
}
var ansConfirm = "";
function mConfirm(title, message, functionname) {
    ansConfirm = "";
    //do {
    var modalHTML = "<center><div class=alert alert-info dvMessage style='width:30%;background-color:white;box-shadow: 0 5px 15px rgba(0, 0, 0, 0.5);font-family:open_sansregular'>" +
         "<div style='text-align:left;margin:0px 5px;border-bottom:1px solid gray;'><h4 class='modal-title'>" + title + "</h4></div>" +
         "<div class='modal-body' style='padding: 12px;position: relative;font-size: 14px;text-align:left;color:#666666;font-weight:normal;' id='dvMessageText'>" + message + "</div>" +
         " <div class='modal-footer' style='border-top: 1px solid transparent;padding: 12px;text-align: right;'>" +
         "<div class='btn-group'><button data-dismiss='modal' type='button' class='btn btn-success' onclick=\"ansConfirm=true;closeConfirm();" + functionname + "\">Ok</button></div>" +
         " &nbsp;<div class='btn-group'><button data-dismiss='modal' type='button' class='btn btn-info' onclick='ansConfirm=false;closeConfirm();'>Cancel</button></div></div>"
    $('#Confirm').html(modalHTML).show(300);

}

// This function will close confirm box
function closeConfirm() {
    $('#Confirm').html('').hide(500);
}
