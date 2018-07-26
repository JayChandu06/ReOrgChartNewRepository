$(document).ready(function () {

    const sidebarWrapper = $('.page-sidebar-wrapper');

    let loginButton = $('#login-form').find('button'),
        registerButton = $('#register-form').find('button'),
        toggleBtn = $('.quicklinks.layout-toggle');

    sidebarWrapper.find('ul li').on('click',function (e) {
        sidebarWrapper.find('ul li').each(function () {
            $(this).removeClass('open active');
        });
        $(this).addClass('open active')
    });

    loginButton.on('click',function (e) {
        e.preventDefault();
        window.location.href = 'index.html';
    });
    registerButton.on('click',function (e) {
        e.preventDefault();
        window.location.href = 'index.html';
    });
    toggleBtn.on('click', function () {
        alert("Left Icon");
        let header = $('.header-seperation');
        let bodyContent = $('.page-content');
        if(header.hasClass('full') || bodyContent.hasClass('full')){
            header.removeClass('full');
            bodyContent.removeClass('full');
        }else{
            header.addClass('full');
            bodyContent.addClass('full');
        }
    })
    //toggleBtn.on('click', function () {
    //    let header = $('.header-seperation');
    //    let bodyContent = $('.page-content');
    //    let sideBar = $('.page-sidebar-wrapper');

    //    if (header.hasClass('full') || bodyContent.hasClass('full') || sideBar.hasClass('full')) {
    //        header.removeClass('full');
    //        bodyContent.removeClass('full');
    //        sideBar.removeClass('full');
    //    } else {
    //        header.addClass('full');
    //        bodyContent.addClass('full');
    //        sideBar.addClass('full');
    //    }
    //});

    
        //$('#example').dataTable({
        //    "sPaginationType": "two_button",
        //    "bFilter": false,
        //    "bLengthChange": false
        //});
    
    $('#excel,#item_pic').on('change', function (e) {
        var fileName = e.target.files[0].name;
        $(this).siblings('label').find('p').html(fileName);
    });

});
