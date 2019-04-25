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

    }
};
homeController.init();