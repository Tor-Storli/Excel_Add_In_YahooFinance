async function GetDataFromWebService() {

    var ticker = document.querySelector('#txtTicker').value;
    var range = document.querySelector('#selRange').selectedOptions[0].value;

    const responseArray =
        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "GetData.asmx/GetDataFromWebApi",
            dataType: "json",
            data: "{'ticker':'" + ticker + "', 'range':'" + range + "' }",
            success: function (data) {
                return data.d;
            }
        });

    return responseArray;
};
