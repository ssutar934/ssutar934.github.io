

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
                /*alert(error + status);*/
            }
        });



        result = [
            { OrderNo: '21581', Product: 'MacBook Pro', Quantity: '1', Status: 'Completed' },
            { OrderNo: '21582', Product: 'iPhone', Quantity: '5', Status: 'Completed' },
            { OrderNo: '21583', Product: 'Google Pixel 7', Quantity: '3', Status: 'Processing' },
            { OrderNo: '21584', Product: 'iPad mini ', Quantity: '2', Status: 'Cancelled' },
            { OrderNo: '21585', Product: 'Asus VivoBook', Quantity: '1', Status: 'Processing' },
            { OrderNo: '21586', Product: 'Panasonic', Quantity: '3', Status: 'Cancelled' },
            { OrderNo: '21587', Product: 'HP SPECTRE X360', Quantity: '2', Status: 'Processing' },
            { OrderNo: '55321', Product: 'ASUS ROG ZEPHYRUS', Quantity: '1', Status: 'Processing' },
            { OrderNo: '55322', Product: 'Lenova Yoga 7i', Quantity: '1', Status: 'Cancelled' },
            { OrderNo: '22323', Product: 'Dell XPS', Quantity: '2', Status: 'Completed' }
        ]


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