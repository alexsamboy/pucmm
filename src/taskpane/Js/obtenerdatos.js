var myHeaders = new Headers();
myHeaders.append("Content-Type", "application/json");

var raw = JSON.stringify({
  "correo": "cbueno@pucmm.edu.do"
});

let dataUser;

var requestOptions = {
  method: 'POST',
  headers: myHeaders,
  body: raw,
  redirect: 'follow'
};

fetch("https://prod-29.westus.logic.azure.com:443/workflows/b99ce35126ee4b278be693e44ee31bf2/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BNBlKqFf8vaViNf7nlHudzobTdzxa4n0DcgXr8oXe_c", requestOptions)
  .then(response => response.json())
  .then(result => dataUser = result)
  .catch(error => console.log('error', error));

  console.log(dataUser);