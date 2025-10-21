

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
console.log("first line");

/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/



function onNewAppointmentComposeHandler(event) {
  setMessage(event);
}


function setMessage(event) {
  var item = Office.context.mailbox.item;
  var htmlBody = "<p> <b> Statement of Achievement </b> <br/> " 
                  + " <b> Meeting Type (informational or decision): </b> <br/> "
                  +" <b> Agenda: </b> <br/> <b>Facilitator: </b> <br/> <b> Note Taker: </b> </p>";

  item.body.prependAsync(htmlBody, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Inserted HTML to body");
    } else {
      console.log("Error: " + asyncResult.error.message);
    }
    event.completed();
  });
}
