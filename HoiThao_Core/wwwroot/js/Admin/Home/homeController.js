$.formattedDate = function (dateToFormat) {
    var dateObject = new Date(dateToFormat);

    var day = dateObject.getDate();
    var month = dateObject.getMonth() + 1;
    var year = dateObject.getFullYear();
    day = day < 10 ? "0" + day : day;
    month = month < 10 ? "0" + month : month;
    var formattedDate = day + "/" + month + "/" + year;
    return formattedDate;
};

$.formattedDateTime = function (dateToFormat) {
    var dateObject = new Date(dateToFormat);

    if (dateObject.getHours() >= 12) {
        var hour = parseInt(dateObject.getHours()) - 12;
        var amPm = "PM";
    } else {
        var hour = dateObject.getHours();
        var amPm = "AM";
    }
    var time = hour + ":" + dateObject.getMinutes() + ":" + dateObject.getSeconds() + " " + amPm;

    var day = dateObject.getDate();
    var month = dateObject.getMonth() + 1;
    var year = dateObject.getFullYear();
    day = day < 10 ? "0" + day : day;
    month = month < 10 ? "0" + month : month;
    var formattedDate = day + "/" + month + "/" + year + " " + time;
    return formattedDate;
};


$.stringToDate = function (_date, _format, _delimiter) {
    var formatLowerCase = _format.toLowerCase();
    var formatItems = formatLowerCase.split(_delimiter);
    var dateItems = _date.split(_delimiter);
    var monthIndex = formatItems.indexOf("mm");
    var dayIndex = formatItems.indexOf("dd");
    var yearIndex = formatItems.indexOf("yyyy");
    var month = parseInt(dateItems[monthIndex]);
    month -= 1;
    var formatedDate = new Date(dateItems[yearIndex], month, dateItems[dayIndex]);
    return formatedDate;
};


$.getMyFormatDate = function (date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var d = date;
    var hours = d.getHours();
    var ampm = hours >= 12 ? 'PM' : 'AM';
    return months[d.getMonth()] + ' ' + d.getDate() + " " + d.getFullYear() + ' ' + hours + ':' + d.getMinutes() + ' ' + ampm;
}

var homeController = {
    init: function () {
        homeController.registerEvent();
    },

    registerEvent: function () {
        $('#ddlStatusS').off('change').on('change', function () {
            $('#frmSearch').submit();
        });

        $('#btnSearch').click(function () {
            $('#frmSearch').submit();

        });

        $('.paymentTr').click(function () {
            if ($(this).hasClass("hoverClass"))
                $(this).removeClass("hoverClass");
            else {
                $('tr.hoverClass').removeClass("hoverClass");
                $(this).addClass("hoverClass");
            }

        });

        //$('#txtNameS').off('change keydown paste input').on('change keydown paste input', function () {
        //    $('#frmSearch').submit();
        //});

        $('.txtCheckin').off('keydown').on('keydown', function (e) {
            var d = new Date();
            if (e.which === 113) {//F2

                var k = $(this).data('kid');

                $.ajax({
                    url: '/Home/UpdateCheckin',
                    type: 'POST',
                    data: { id: k },
                    dataType: 'json',
                    success: function (response) {
                        if (response.status) {
                            bootbox.alert({
                                size: "small",
                                title: "Checkin Infomation",
                                message: "Checkin success.",
                                callback: function () {
                                    $('#refresh').submit();

                                }
                            });


                        }


                    }
                });

            }
        });

        $('.btn-delete').off('click').on('click', function () {
            var id = $(this).data('k');
            bootbox.confirm('Do you want to delete this item?', function (result) {
                if (result) {
                    $.ajax({
                        url: '/Home/Delete',
                        type: 'DELETE',
                        data: { id: id },
                        dataType: 'json',
                        success: function (response) {
                            if (response.status) {
                                bootbox.alert('Delete success.');
                                $('#refresh').submit();
                            }
                            else {
                                bootbox.alert('Delete fail.');
                            }

                        }
                    });
                }
            });
        });

        //$('.btn-PrintReceipt').off('click').on('click', function () {
        //    var id = $(this).data('kid');
        //    window.open('/Home/PrintReceipt?id=' + id);

        //});

        //$('.btn-PrintBadge').off('click').on('click', function () {
        //    var id = $(this).data('k');

        //    $.get('/Home/PrintBadge', { id: id }, function (data) {
        //        console.log(data);
        //        homeController.CreatePdf();
        //    });

        //});
        $('.aseanTr').off('click').on('click', function () {
            var k = $(this).data('kid');
            homeController.getDetail(k);
        });
        
    },
    getDetail: function (id) {
        $.ajax({
            url: '/Home/GetDetail',
            data: {
                id: id
            },
            dataType: 'json',
            type: 'GET',
            success: function (response) {
                if (response.status) {
                    var data = response.data;

                    if (data.checkin === null)
                        ck = "";
                    else {
                        ck = $.formattedDateTime(new Date(parseInt(data.checkin.substr(6))));
                    }

                    if (data.checkin === null)
                        ckHid = "";
                    else {
                        ckHid = $.formattedDate(new Date(parseInt(data.checkin.substr(6))));
                    }

                    if (data.dangky === null)
                        dk = "";
                    else {
                        dk = $.formattedDateTime(new Date(parseInt(data.dangky.substr(6))));
                    }
                    $('#Hotel').text(data.Hotel);
                    $('#HotelCheckin').text(data.HotelCheckin);
                    $('#HotelChceckout').text(data.HotelCheckout);
                    $('#HotelPrice').text(data.HotelPrice);
                    $('#HotelBookingInf').text(data.HotelBookingInf);
                    $('#At').text(data.at);

                    $('#Dep').text(data.dt);

                    $('#Address').text(data.email);
                    $('#Department').text(data.email);
                    $('#Institution').text(data.email);
                    $('#Note').text(data.dfno);

                }
                else {
                    bootbox.alert({
                        size: "small",
                        title: "Detail User Infomation",
                        message: response.message
                    });
                }
            }
        });
    },
};
homeController.init();