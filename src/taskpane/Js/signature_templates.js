// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.


var myHeaders = new Headers();
myHeaders.append("Authorization", "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6Ik9yY0Z1dG1aeFBMbDF2NGpwZVRsNFF6ckY1NjlESlNhTmFnRXpsd19hVDQiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83M2M5YTQxOS04NjNkLTQyMjYtYTgzZi03YTIwMGFkNjliZTkvIiwiaWF0IjoxNjU5MDIxMTk5LCJuYmYiOjE2NTkwMjExOTksImV4cCI6MTY1OTAyNTk1MSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkUyWmdZTWc5RyszK1k0ZU0wRVdMUFEreTh5S2Z2V1hwTFh2amJ2MVU1L2lNQTlGMVc5VUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6Ik8zNjV3cFVzZXIiLCJhcHBpZCI6Ijg2NGZhODY0LTdiOGYtNGFlMC05NWE5LTM5YTU4N2ZkZjYwYyIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoiUMOpcmV6IFNhbWJveSIsImdpdmVuX25hbWUiOiJNYW51ZWwgQWxleGFuZGVyIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTkwLjExMy43Ny41NSIsIm5hbWUiOiJNYW51ZWwgQWxleGFuZGVyIFDDqXJleiBTYW1ib3kiLCJvaWQiOiI2MThkNjg4OS1lNTZlLTQwNDktODBjZS01ZWZhZGE2NWJkNjAiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtNDE0MDAyNzMyNy0yOTgwMzkyMTQ4LTExMDQwNzk4MDgtMjU2NyIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzN0ZGRTg5RjUzMEZBIiwicmgiOiIwLkFRNEFHYVRKY3oyR0prS29QM29nQ3RhYjZRTUFBQUFBQUFBQXdBQUFBQUFBQUFBT0FQOC4iLCJzY3AiOiJlbWFpbCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6ImVrOV9qeFNWQi1HVjBoZUhLQjdMbWROZk9GMVF2QzJFWUQ0M1NKM0dLLVkiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiI3M2M5YTQxOS04NjNkLTQyMjYtYTgzZi03YTIwMGFkNjliZTkiLCJ1bmlxdWVfbmFtZSI6Im1hbnVlbHBlcmV6QHB1Y21tLmVkdS5kbyIsInVwbiI6Im1hbnVlbHBlcmV6QHB1Y21tLmVkdS5kbyIsInV0aSI6ImF5MkZXLXpfVzBHeGlnT3JtSTJBQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjc1OTM0MDMxLTZjN2UtNDE1YS05OWQ3LTQ4ZGJkNDllODc1ZSIsImNmMWMzOGU1LTM2MjEtNDAwNC1hN2NiLTg3OTYyNGRjZWQ3YyIsIjExNjQ4NTk3LTkyNmMtNGNmMy05YzM2LWJjZWJiMGJhOGRjYyIsIjRhNWQ4ZjY1LTQxZGEtNGRlNC04OTY4LWUwMzViNjUzMzljZiIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoibFZuWW5JNnNjU3otTURnNWJjQnEyejlaWWxKZXpWeTNSYlZZdkpVVVZBMCJ9LCJ4bXNfdGNkdCI6MTM5ODI5ODA1NX0.D71CUUQKQfLnKrumBA06SpbSc39c36zBdTHef_HYqc5xmBwXDqwMCbZgOolKw-yAj7mcdBePo8pbP96kGIETBN1XSKI_qzuoA6CE6tt_0RIvttp_EYrFlfEZlZQZBoylFOysqlyW74hm5iPYBB5XpzRJTAz4Q2f-eljx0k4CBSQEg3gswxmyG_sgyyDu9A1lE5jJsLUwT-4KZcIDC63tZx8NFqpLatEaIZMLoJt-0URd9u_shn6AHplhVjAyRcDd7q-a4YFmai_eaJLI-sGNw_p8Gw3xf7xM8psVt9xrAZBC3X8kdkNoJ2NwFu7ioDnyT19C-2nk5hAeVu4G6INqOA");

var requestOptions = {
  method: 'GET',
  headers: myHeaders,
  redirect: 'follow'
};

/*if (is_valid_data(user_info.greeting))
  {
    let urlApi = "";
  }*/
  let urlApi = "";
  urlApi+=`https://graph.microsoft.com/v1.0/users/`;
  urlApi+= user_info.email;
  urlApi+=`?$select=displayName,jobTitle,officeLocation,businessPhones,mail`;

fetch(urlApi, requestOptions)
  .then(response => response.json())
  .then(result => get_template_A_str(result))
  .catch(error => console.log('error', error));




function get_template_A_str(result)
{
  let str = "";
 
  str += `<table border="0" cellpadding="5" cellspacing="5"><tbody><tr><td><font size="+1" color="#17365d"><strong>`;
  str += result.displayName;
  str += `</strong></font><br>`;
  str += result.jobTitle;
  str += `<br><font size="+1" color="#17365d"><strong>`;
  str += result.officeLocation;
  str += `</strong></font><br> 809-535-0111, ext. :`;
  str += result.businessPhones[0];
  str += `<br><font color="#17365d">`;
  str += result.mail;
  str += `</font></td></tr><tr><td><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="258" height="87"><img src="https://alexsamboy.github.io/pucmm/assets/marca-pucmm.jpg" width="258" height="87" alt=""/></td><td width="15"></td><td style="padding:0 0 0 15px;border-left-style:solid;border-left-width:1pt;border-left-color:#7F7F7F;"><p><font><strong>Campus de Santiago:</strong><br>Autopista Duarte km. 1½, Santiago, R.D.<br><br><strong>Campus de Santo Domingo:</strong><br>Av. Abraham Lincoln. esq. Av. Bolívar, Santo Domingo, R.D. </font> </p></td></tr></tbody></table></td></tr><tr><td height="70" align="left" valign="middle" ><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="601" height="71"><img src="https://alexsamboy.github.io/pucmm/assets/banner.jpg" width="601" height="71" alt=""/></td><td width="15"></td><td><a href="http://www.facebook.com/pucmm/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/facebook.png" alt="" width="24" height="24"/></a> <a href="http://twitter.com/pucmm/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/twitter.png" alt="" width="24" height="25"/></a> <a href="http://www.instagram.com/pucmm" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/instagram.png" alt="" width="24" height="25"/></a> <a href="http://www.youtube.com/pucmmtv/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/youtube.png" alt="" width="24" height="25"/></a> <a href="https://www.linkedin.com/edu/school?id=12020" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/linkedin.png" alt="" width="24" height="25"/></a></td></tr></tbody></table></td></tr><tr><td><img src="https://alexsamboy.github.io/pucmm/assets/Green.gif" width="14" height="14" alt=""/> <font color="#6C6C6C">No me imprimas si no es necesario.</font></td></tr><tr><td><p style="margin: 0"><font color="#6C6C6C"> NOTA DE CONFIDENCIALIDAD: La información transmitida, incluidos los archivos adjuntos, está dirigida solo a la persona o entidad a la que ha sido remitida y puede contener información confidencial y/o privilegiada. Cualquier difusión u otro uso de la misma, o tomar cualquier acción basada en esta información por personas o entidades distintas al destinatario, está prohibido. Si ha recibido este mensaje por error, por favor contactar al remitente y destruya cualquier copia de esta información.<br><br>CONFIDENTIALITY NOTE: The information transmitted, including attachments, is intended only for the person or entity to which it is addressed and may contain confidential and/or privileged material. Any review, retransmission, dissemination, or other use of, or taking of any action in reliance upon this information by persons or entities other than the intended recipient is prohibited. If you received this in error, please contact the sender and destroy any copies of this information.</font></p></td></tr></tbody></table>`;

  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://hurodriguez.github.io/pucmm/firma/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;
  
  return str;
}