function update(innerhtml) {
  var xhttp = new ActiveXObject("Microsoft.XMLHTTP");
  xhttp.onreadystatechange = function() {
   if (xhttp.readyState == 4) {
      document.getElementById("clin_input_table").innerHTML = xhttp.responseText;
   }
  };
  xhttp.open("GET", "innerhtml.html", true);
  xhttp.send();
}
