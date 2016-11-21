'use strict';

var IndexPage = (function () {
    var _urlExcelProviders;

    return {
        init: function (urlExcelProviders) {
            _urlExcelProviders = urlExcelProviders;
        }
    };


    function getRequestExcelProviders() {
        return $.post(_urlExcelProviders);
    }

    function LoadProviders() {

    }

   
})();