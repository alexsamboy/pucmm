// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.


function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }
 function midata(result){
  str += `<table border="0" cellpadding="5" cellspacing="5"><tbody><tr><td><font size="+1" color="#17365d"><strong>`;
  str += user_info.name;
  str += `</strong></font><br>`;
<<<<<<< HEAD
  str += is_valid_data(user_info.job) ? user_info.job;
  str += `<br><font size="+1" color="#17365d"><strong>`;
  str += `Comunicaciones`;
  str += `</strong></font><br>`;
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
=======
  str += is_valid_data(user_info.job) ? user_info.job : "";
  str += `<br><font size="+1" color="#17365d"><strong>`;
  str += `Dirección de Comunicaciones`;
  str += `</strong></font><br>`;
  str += is_valid_data(user_info.phone) ? user_info.phone + ", ext. " : "";
>>>>>>> parent of 99a6e20 (Update signature_templates.js)
  str += `<br><font color="#17365d">`;
  str += user_info.email;
  str += `</font></td></tr><tr><td><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="258" height="87"><img src="https://alexsamboy.github.io/pucmm/assets/marca-pucmm.jpg" width="258" height="87" alt=""/></td><td width="15"></td><td style="padding:0 0 0 15px;border-left-style:solid;border-left-width:1pt;border-left-color:#7F7F7F;"><p><font><strong>Campus de Santiago:</strong><br>Autopista Duarte km. 1½, Santiago, R.D.<br><br><strong>Campus de Santo Domingo:</strong><br>Av. Abraham Lincoln. esq. Av. Bolívar, Santo Domingo, R.D. </font> </p></td></tr></tbody></table></td></tr><tr><td height="70" align="left" valign="middle" ><table border="0" cellpadding="0" cellspacing="0"><tbody><tr><td width="601" height="71"><img src="https://alexsamboy.github.io/pucmm/assets/banner.jpg" width="601" height="71" alt=""/></td><td width="15"></td><td><a href="http://www.facebook.com/pucmm/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/facebook.png" alt="" width="24" height="24"/></a> <a href="http://twitter.com/pucmm/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/twitter.png" alt="" width="24" height="25"/></a> <a href="http://www.instagram.com/pucmm" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/instagram.png" alt="" width="24" height="25"/></a> <a href="http://www.youtube.com/pucmmtv/" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/youtube.png" alt="" width="24" height="25"/></a> <a href="https://www.linkedin.com/edu/school?id=12020" target="_blank"><img src="https://alexsamboy.github.io/pucmm/assets/linkedin.png" alt="" width="24" height="25"/></a></td></tr></tbody></table></td></tr><tr><td><img src="https://alexsamboy.github.io/pucmm/assets/Green.gif" width="14" height="14" alt=""/> <font color="#6C6C6C">No me imprimas si no es necesario.</font></td></tr><tr><td><p style="margin: 0"><font color="#6C6C6C"> NOTA DE CONFIDENCIALIDAD: La información transmitida, incluidos los archivos adjuntos, está dirigida solo a la persona o entidad a la que ha sido remitida y puede contener información confidencial y/o privilegiada. Cualquier difusión u otro uso de la misma, o tomar cualquier acción basada en esta información por personas o entidades distintas al destinatario, está prohibido. Si ha recibido este mensaje por error, por favor contactar al remitente y destruya cualquier copia de esta información.<br><br>CONFIDENTIALITY NOTE: The information transmitted, including attachments, is intended only for the person or entity to which it is addressed and may contain confidential and/or privileged material. Any review, retransmission, dissemination, or other use of, or taking of any action in reliance upon this information by persons or entities other than the intended recipient is prohibited. If you received this in error, please contact the sender and destroy any copies of this information.</font></p></td></tr></tbody></table>`;

  return str;
 }
 return midata;
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