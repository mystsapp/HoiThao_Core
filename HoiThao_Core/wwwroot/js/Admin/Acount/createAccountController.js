var createAccountController = {
    init: function () {
        createAccountController.registerEvent();
    },

    registerEvent: function () {

        //Remove the dummy row if data present.
        //if ($("#tblAccount tr").length > 2) {
        //    console.log($("#tblAccount tr").length);
        //    $("#tblAccount tr:eq(1)").remove();
        //} else {
        //    var row = $("#tblAccount tr:last-child");
        //    row.find(".Edit").hide();
        //    row.find(".Delete").hide();
        //    row.find("span").html('&nbsp;');
        //}

        $('.Add').off('click').on('click', function () {
            $('#modalUpdate').modal('show');
        });

        //Add event handler.
        //$("#btnAdd").on("click", function () {
        $("body").on("click", "#btnAdd", function () {
            var txtUsername = $("#txtUsername").val();
            var txtPassword = $("#txtPassword").val();
            var txtHoTen = $("#txtHoTen").val();
            var ckTrangThai = $('#ckTrangThai').prop('checked');

            $.ajax({
                type: "POST",
                url: "/Accounts/InsertAccount",
                data: { username: txtUsername, password: txtPassword, hoten: txtHoTen, trangthai: ckTrangThai },
                dataType: "json",
                success: function (r) {
                    $('#modalUpdate').modal('hide');

                    var row = $("#tblAccount tr:last-child");
                    if ($("#tblAccount tr:last-child span").eq(0).html() !== "&nbsp;") {
                        row = row.clone();
                    }

                    createAccountController.AppendRow(row, r.username, r.password, r.hoten, r.trangthai);
                    $("#txtUsername").val("");
                    $("#txtPassword").val("");
                    $("#txtHoTen").val("");
                    $('#ckTrangThai').prop('checked', false);
                }
            });
        });
    },

    AppendRow: function (row, username, password, hoten, trangthai) {
        //Bind CustomerId.
        $(".username", row).find("span").html(username);
        $(".username", row).find("input").val(username);

        //Bind Name.
        //$(".password", row).find("span").html(password);
        //$(".password", row).find("input").val(password);

        //Bind Country.
        $(".hoten", row).find("span").html(hoten);
        $(".hoten", row).find("input").val(hoten);

        $(".trangthai", row).find("span").html(trangthai);
        $(".trangthai", row).find("input").val(trangthai);

        row.find(".Edit").show();
        row.find(".Delete").show();
        $("#tblAccount").append(row);
    }

};
createAccountController.init();