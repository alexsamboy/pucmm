// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on both Outlook on web and Outlook on Windows.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows only. Insert signature into appointment or message.
 * Outlook on Windows can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Por favor confimer su firma..",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Establecer firma",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "marca-pucmm.jpg";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong> firma por defecto" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:"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"
      /*"iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC"*/,
    logoFileName: logoFileName,
  };
}


/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://alexsamboy.github.io/pucmm/assets/marca-pucmm.jpg' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
