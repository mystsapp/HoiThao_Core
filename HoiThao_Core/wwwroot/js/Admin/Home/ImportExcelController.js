var importExcelController = {
    init: function () {
        importExcelController.registerEvent();
    },

    registerEvent: function () {
        $('#btnUpload').on('click', function () {
            var fileExtension = ['xls', 'xlsx'];
            var filename = $('#fUpload').val();
            if (filename.length === 0) {
                alert("Please select a file.");
                return false;
            }
            else {
                var extension = filename.replace(/^.*\./, '');
                if ($.inArray(extension, fileExtension) === -1) {
                    alert("Please select only excel files.");
                    return false;
                }
            }
            var fdata = new FormData();
            var fileUpload = $("#fUpload").get(0);
            var files = fileUpload.files;
            fdata.append(files[0].name, files[0]);
            $.ajax({
                type: "POST",
                //url: "/ImportExcel?handler=Import",
                url: "/Home/UploadExcel", //OnPostImport
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("XSRF-TOKEN",
                        $('input:hidden[name="__RequestVerificationToken"]').val());
                },
                data: fdata,
                contentType: false,
                processData: false,
                success: function (response) {
                    if (response.status)
                        alert('Upload success.');
                    else {
                        //$('#dvData').html(response);
                        alert('Some error occured while uploading');
                    }
                },
                error: function (e) {
                    $('#dvData').html(e.responseText);
                }
            });
        });
    }
};
importExcelController.init();