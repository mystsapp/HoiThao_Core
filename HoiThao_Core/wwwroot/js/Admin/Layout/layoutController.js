var layoutController = {
    init: function () {
        layoutController.registerEvent();
    },
    registerEvent: function () {
        $('#HotelList').off('click').on('click', function (event) {
            event.preventDefault();
            $('#hotelModal').modal('show');

            layoutController.loadDdlHotel();

            $('#btnReport').off('click').on('click', function () {
                //            $('#hotelModal').modal('show');
                var hotel = $('#ddlHotel').val();
                //alert(hotel);
                //layoutController.ExportHotel(hotel);
                $('#hidHotel').val(hotel);
                $('#frmHotel').on('submit');
                $('#hotelModal').modal('hide');
            });
        });
    },

    loadDdlHotel: function () {
        $('#ddlHotel').html('');
        var option = '';
        // option = option + '<option value=select>Select</option>';

        $.ajax({
            url: '/Home/GetAllHotel',
            type: 'GET',
            dataType: 'json',
            success: function (response) {
                
                var data = JSON.parse(response.data);

                for (var i = 0; i < data.length; i++) {
                    // set the key/property (input element) for your object
                    var ele = data[i];
                    if (ele === null)
                        ele = 'Other';
                    option = option + '<option value="' + ele + '">' + ele + '</option>'; //chinhanh1
                    // add the property to the object and set the value
                    //params[ele] = $('#' + ele).val();
                }
                $('#ddlHotel').html(option);

            }
        });
        //homeController.resetForm();
    }
};
layoutController.init();