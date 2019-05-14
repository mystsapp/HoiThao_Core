var createAseanController = {
    init: function () {
        createAseanController.registerEvent();
    },

    registerEvent: function () {

        //Remove the dummy row if data present.
        //if ($("#tblAsean tr").length > 2) {
        //    console.log($("#tblAsean tr").length);
        //    $("#tblAsean tr:eq(1)").remove();
        //} else {
        //    var row = $("#tblAsean tr:last-child");
        //    row.find(".Edit").hide();
        //    row.find(".Delete").hide();
        //    row.find("span").html('&nbsp;');
        //}

        $('.Add').off('click').on('click', function () {
            $('#modalUpdate').modal('show');
        });
    },

    AppendRow: function (row, customerId, name, country) {
        //Bind CustomerId.
        $(".CustomerId", row).find("span").html(customerId);

        //Bind Name.
        $(".Name", row).find("span").html(name);
        $(".Name", row).find("input").val(name);

        //Bind Country.
        $(".Country", row).find("span").html(country);
        $(".Country", row).find("input").val(country);

        row.find(".Edit").show();
        row.find(".Delete").show();
        $("#tblCustomers").append(row);
    }

};
createAseanController.init();