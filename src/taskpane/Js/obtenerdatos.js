var myHeaders = new Headers();
myHeaders.append("Authorization", "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6Ik9yY0Z1dG1aeFBMbDF2NGpwZVRsNFF6ckY1NjlESlNhTmFnRXpsd19hVDQiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83M2M5YTQxOS04NjNkLTQyMjYtYTgzZi03YTIwMGFkNjliZTkvIiwiaWF0IjoxNjU5MDIxMTk5LCJuYmYiOjE2NTkwMjExOTksImV4cCI6MTY1OTAyNTk1MSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkUyWmdZTWc5RyszK1k0ZU0wRVdMUFEreTh5S2Z2V1hwTFh2amJ2MVU1L2lNQTlGMVc5VUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik8zNjV3cFVzZXIiLCJhcHBpZCI6Ijg2NGZhODY0LTdiOGYtNGFlMC05NWE5LTM5YTU4N2ZkZjYwYyIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoiUMOpcmV6IFNhbWJveSIsImdpdmVuX25hbWUiOiJNYW51ZWwgQWxleGFuZGVyIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTkwLjExMy43Ny41NSIsIm5hbWUiOiJNYW51ZWwgQWxleGFuZGVyIFDDqXJleiBTYW1ib3kiLCJvaWQiOiI2MThkNjg4OS1lNTZlLTQwNDktODBjZS01ZWZhZGE2NWJkNjAiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtNDE0MDAyNzMyNy0yOTgwMzkyMTQ4LTExMDQwNzk4MDgtMjU2NyIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzN0ZGRTg5RjUzMEZBIiwicmgiOiIwLkFRNEFHYVRKY3oyR0prS29QM29nQ3RhYjZRTUFBQUFBQUFBQXdBQUFBQUFBQUFBT0FQOC4iLCJzY3AiOiJlbWFpbCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6ImVrOV9qeFNWQi1HVjBoZUhLQjdMbWROZk9GMVF2QzJFWUQ0M1NKM0dLLVkiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiI3M2M5YTQxOS04NjNkLTQyMjYtYTgzZi03YTIwMGFkNjliZTkiLCJ1bmlxdWVfbmFtZSI6Im1hbnVlbHBlcmV6QHB1Y21tLmVkdS5kbyIsInVwbiI6Im1hbnVlbHBlcmV6QHB1Y21tLmVkdS5kbyIsInV0aSI6ImF5MkZXLXpfVzBHeGlnT3JtSTJBQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjc1OTM0MDMxLTZjN2UtNDE1YS05OWQ3LTQ4ZGJkNDllODc1ZSIsImNmMWMzOGU1LTM2MjEtNDAwNC1hN2NiLTg3OTYyNGRjZWQ3YyIsIjExNjQ4NTk3LTkyNmMtNGNmMy05YzM2LWJjZWJiMGJhOGRjYyIsIjRhNWQ4ZjY1LTQxZGEtNGRlNC04OTY4LWUwMzViNjUzMzljZiIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoibFZuWW5JNnNjU3otTURnNWJjQnEyejlaWWxKZXpWeTNSYlZZdkpVVVZBMCJ9LCJ4bXNfdGNkdCI6MTM5ODI5ODA1NX0.D71CUUQKQfLnKrumBA06SpbSc39c36zBdTHef_HYqc5xmBwXDqwMCbZgOolKw-yAj7mcdBePo8pbP96kGIETBN1XSKI_qzuoA6CE6tt_0RIvttp_EYrFlfEZlZQZBoylFOysqlyW74hm5iPYBB5XpzRJTAz4Q2f-eljx0k4CBSQEg3gswxmyG_sgyyDu9A1lE5jJsLUwT-4KZcIDC63tZx8NFqpLatEaIZMLoJt-0URd9u_shn6AHplhVjAyRcDd7q-a4YFmai_eaJLI-sGNw_p8Gw3xf7xM8psVt9xrAZBC3X8kdkNoJ2NwFu7ioDnyT19C-2nk5hAeVu4G6INqOA");

var requestOptions = {
  method: 'GET',
  headers: myHeaders,
  redirect: 'follow'
};

let urlApi="";
urlApi+=`https://graph.microsoft.com/v1.0/users/`;
  urlApi+= `manuelperez@pucmm.edu.do`;
  urlApi+=`?$select=displayName,jobTitle,officeLocation,businessPhones,mail`;


fetch(urlApi, requestOptions)
  .then(response => response.json())
  .then(result => mifuncion(result))
  .catch(error => console.log('error', error));

  function mifuncion(result){
    console.log(result.displayName)
  }