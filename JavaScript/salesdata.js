

function get_order_details() {
    try {
        let result = [];
        $.ajax({
            type: "POST",
            url: "SalesHandler.asmx/GetOrderData",
            contentType: "application/json",
            datatype: "json",
            async: false,
            success: function (responseFromServer) {
                //alert(responseFromServer.d);
                let d = JSON.parse(responseFromServer.d);
                result = d.Records;

            },
            error: function (xhr, status, error) {
                alert(error + status);
            }
        });

        return result;
    } catch (e) {
        alert(e);
    }
}

function detailedordertrack(record) {
    try {

        $('#detailedtrack').css('display', '');
        $('#orderno').html('#' + record.OrderNo);

        $('.step0').removeClass('active');
        $('#arrivedate').html('--/--/----');

        if (record.Status == 'Completed') {
            $('.step0').addClass('active');
        }
        else if (record.Status == 'Processing') {
            $('#st01').addClass('active');
            $('#st02').addClass('active');
            $('#arrivedate').html('25/08/2022');
        }
        else if (record.Status == 'Cancelled') {
           
        }

    } catch (e) {
        alert(e);
    }
}