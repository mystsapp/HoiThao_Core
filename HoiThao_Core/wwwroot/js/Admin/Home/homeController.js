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

        $('.btn-PrintReceipt').off('click').on('click', function () {
            var id = $(this).data('kid');
            window.open('/Home/PrintReceipt?id=' + id);

        });

        //$('.btn-PrintVAT').off('click').on('click', function () {
        //    var id = $(this).data('kid');
        //    window.open('/Home/PrintVAT?id=' + id);

        //});
        $('.btn-PrintBadge').off('click').on('click', function () {
            var id = $(this).data('k');
            //$.ajax({
            //    url: '/Home/PrintBadge',
            //    type: 'POST',
            //    data: { id: id },
            //    dataType: 'json',
            //    success: function (response) {
            //        console.log(response);
            //        homeController.CreatePdf();
            //    }
            //});
            $.get('/Home/PrintBadge', { id: id }, function (data) {
                console.log(data);
                homeController.CreatePdf();
            });

        });
        
    },

    CreatePdf: function () {
        $.ajax({
            url: '/Home/CreatePDF',
            type: 'GET',
            dataType: 'json'
        });
    }
};
homeController.init();