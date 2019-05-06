jQuery(function ($) {

    $("#btnPrint").off('click').on('click', function () {
        //Print ele2 with default options
        $("#printReceipt").print({
            //Use Global styles
            globalStyles: false,
            //Add link with attrbute media=print
            mediaPrint: false,
            //Custom stylesheet
            stylesheet: "http://fonts.googleapis.com/css?family=Inconsolata",
            //Print in a hidden iframe
            iframe: false,
            //Don't print this
            noPrintSelector: ".avoid-this"
            //Add this at top
            //prepend: "Hello World!!!<br/>",
            //Add this on bottom
            //append: "<span><br/>Buh Bye!</span>",
            //Log to console when printing is done via a deffered callback
           // deferred: $.Deferred().done(function () { console.log('Printing done', arguments); })
        });
    });

    $("#btnPrintVAT").off('click').on('click', function () {
        //Print ele2 with default options
        $("#printVAT").print();
    });

   // printController.registerEvent();
});

var printController = {
    //init: function () {
    //    homeController.registerEvent();
    //},

    registerEvent: function () {
        $("#btnPrint1").on('click', function () {
            //Print ele2 with default options
            printController.test1();
            
        });

    },
    test1: function () {
        alert('c');
    }
}