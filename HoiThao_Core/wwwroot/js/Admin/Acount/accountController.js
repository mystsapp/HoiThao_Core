var accountController = {
    init: function () {
        accountController.registerEvent();
    },

    registerEvent: function () {
        $('.btn-delete').off('click').on('click', function (e) {
            e.preventDefault();
            var link = $(this).attr("href");
            bootbox.confirm({
                title: "Delete Confirm?",
                message: "Are you sure?",
                buttons: {
                    cancel: {
                        label: '<i class="fa fa-times"></i> Cancel'
                    },
                    confirm: {
                        label: '<i class="fa fa-check"></i> Confirm'
                    }

                },
                callback: function (result) {
                    if (result) {
                        document.location.href = link;
                    }
                }

            });
        });
    }
};
accountController.init();