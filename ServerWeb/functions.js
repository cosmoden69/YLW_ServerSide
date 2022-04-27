function sendPost(address, params) 
{
	$.ajax({
            type: "POST",
            url: address,
            contentType: "Application/json; charset=utf-8",
            dataType: "json",
            data: params,
            success: function (data) {
                alert(data.d);
            }
        });
}
function sendPost2(address, params) 
{
  alert(params);
  var xmlhttp = new ActiveXObject("MSXML2.ServerXMLHTTP");
  xmlhttp.open("POST", address, false);
  xmlhttp.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  xmlhttp.send(params);
  var result = xmlhttp.responseXML;
  alert(result);
}
function WebServiceCall(String address) {
    var con = new XMLHttpRequest();
    con.open("GET", address, false);
    con.send();
    window.close();
}
